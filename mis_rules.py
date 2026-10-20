"""Shared MIS data rules (S1–S10).

One rule set for every entry path — the section form (_do_save), the bulk
Excel upload and the Checker's approval — so a value that the form refuses
can no longer arrive through another door.
"""
import math
import re
from datetime import datetime

from form_defs import SECTION_FIELDS

FIELDS = {f["key"]: f for fs in SECTION_FIELDS.values() for f in fs}
SECTION_OF = {f["key"]: s for s, fs in SECTION_FIELDS.items() for f in fs}

_DATE_FMTS = ("%d/%m/%Y", "%d.%m.%Y", "%d-%m-%Y", "%d/%m/%y", "%d.%m.%y", "%d-%m-%y",
              "%d-%b-%Y", "%d-%b-%y", "%d %b %Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S")
_NA_WORDS = {"NA", "N/A", "N.A.", "N.A", "NIL", "-", "NOT APPLICABLE"}
DATE_OR_NA = r"^(\d{2}/\d{2}/\d{4}|NA)?$"


def num(v):
    try:
        return float(str(v).replace(",", "").strip())
    except (ValueError, TypeError):
        return None


def norm_date(v):
    """Coerce a date value to 'DD/MM/YYYY' or 'NA'. Returns (value, ok); blank -> ("", True)."""
    if v is None or str(v).strip() == "":
        return "", True
    if hasattr(v, "strftime"):
        return v.strftime("%d/%m/%Y"), True
    sv = str(v).strip()
    if sv.upper() in _NA_WORDS:
        return "NA", True
    for fmt in _DATE_FMTS:
        try:
            d = datetime.strptime(sv, fmt).date()
        except ValueError:
            continue
        if 1950 <= d.year <= 2100:
            return d.strftime("%d/%m/%Y"), True
    return sv, False


def check_field(f: dict, val):
    """Same rules as the in-app section widgets. Returns (normalised_value, problem_or_None)."""
    t = f.get("type")
    if t in ("number", "int"):
        x = num(val)
        if x is None:
            return val, "must be a number"
        if t == "int" and x != int(x):
            return val, "must be a whole number"
        if f.get("min") is not None and x < float(f["min"]):
            return val, f"is below the minimum allowed ({f['min']})"
        if f.get("max") is not None and x > float(f["max"]):
            return val, (f"is above the maximum allowed ({f['max']}) "
                         f"— check the unit (e.g. Lakhs vs Rupees)")
        return str(int(x)) if t == "int" else str(val).replace(",", "").strip(), None
    if t == "select":
        opts = f.get("opts") or []
        sv = str(val).strip()
        yn = {"Y": "Yes", "N": "No"}.get(sv.upper(), sv)
        match = next((o for o in opts if o.lower() in (sv.lower(), yn.lower())), None)
        return (match, None) if match else (val, f"must be one of: {', '.join(opts)}")
    if t == "date":
        nv, ok = norm_date(val)
        return (nv, None) if ok else (val, "must be a date as DD/MM/YYYY")
    return val, None


# ── auto-calculated fields ────────────────────────────────────────────────────
def compute_autos(vals: dict) -> dict:
    """Return {auto_key: value_str} for every auto field whose inputs are present."""
    out = {}
    work = dict(vals)
    for s in sorted(SECTION_FIELDS):
        for f in SECTION_FIELDS[s]:
            expr = f.get("auto")
            if not expr:
                continue
            refs = re.findall(r"f\d+", expr)
            xs = {r: num(work.get(r)) for r in refs}
            if any(x is None for x in xs.values()):
                continue
            try:
                v = eval(re.sub(r"f\d+", lambda m: f"xs['{m.group(0)}']", expr), {"xs": xs})
            except ZeroDivisionError:
                continue
            out[f["key"]] = work[f["key"]] = f"{round(v, 4):.4f}".rstrip("0").rstrip(".")
    return out


def clear_hidden(vals: dict) -> dict:
    """Blank conditional fields whose Yes/No trigger hides them (as _do_save does)."""
    out = dict(vals)
    for k, f in FIELDS.items():
        sw = f.get("show_when")
        if sw and not all(str(out.get(a, "") or "") == str(b) for a, b in sw.items()):
            if str(out.get(k, "") or "").strip():
                out[k] = ""
    return out


# ── cross-field consistency ───────────────────────────────────────────────────
def _lbl(k):
    return FIELDS[k]["label"]


def _sum(vals, keys):
    xs = [num(vals.get(k)) for k in keys]
    return None if any(x is None for x in xs) else sum(xs)


_TT_FLEET = ("f112", "f113", "f114", "f115", "f116", "f117")
_TT_BOTTOM = ("f118", "f119", "f120")


def cross_field_errors(vals: dict) -> list:
    """[(keys_involved, message)] for every rule whose inputs are all present."""
    g = lambda k: num(vals.get(k))
    errs = []

    def rule(keys, bad, msg):
        if all(g(k) is not None for k in keys) and bad():
            errs.append((set(keys), msg))

    rule(("f1", "f2", "f142", "f3"), lambda: g("f3") + 0.5 < g("f1") + g("f2") + g("f142"),
         f"'{_lbl('f3')}' must be at least MS + HSD + ATF")
    rule(("f77", "f78"), lambda: g("f78") > g("f77"),
         f"'{_lbl('f78')}' cannot exceed '{_lbl('f77')}'")
    rule(("f72", "f73"), lambda: g("f73") > g("f72"),
         f"'{_lbl('f73')}' cannot exceed '{_lbl('f72')}'")
    rule(("f160", "f161"), lambda: g("f161") > g("f160"),
         f"'{_lbl('f161')}' cannot exceed '{_lbl('f160')}'")
    rule(("f10", "f19", "f11"), lambda: g("f19") + g("f11") > g("f10") + 1,
         f"Solar used + exported cannot exceed '{_lbl('f10')}'")
    rule(("f14", "f16"), lambda: g("f14") > g("f16"),
         f"'{_lbl('f16')}' cannot be below '{_lbl('f14')}'")
    rule(("f110", "f111"), lambda: g("f111") < g("f110"),
         f"'{_lbl('f111')}' cannot be below '{_lbl('f110')}'")
    if g("f110") is not None:
        fleet, bottom = _sum(vals, _TT_FLEET), _sum(vals, _TT_BOTTOM)
        if fleet is not None and fleet > g("f110"):
            errs.append(({"f110", *_TT_FLEET}, f"DCT + TPT TT counts ({fleet:g}) exceed '{_lbl('f110')}'"))
        if bottom is not None and bottom > g("f110"):
            errs.append(({"f110", *_TT_BOTTOM}, f"Bottom-loading TT counts ({bottom:g}) exceed '{_lbl('f110')}'"))
    rule(("f110", "f121"), lambda: g("f121") > g("f110"),
         f"'{_lbl('f121')}' cannot exceed '{_lbl('f110')}'")
    rule(("f122", "f124"), lambda: g("f124") > g("f122"),
         f"'{_lbl('f124')}' cannot exceed '{_lbl('f122')}'")
    rule(("f98", "f154", "f155"), lambda: g("f154") + g("f155") > g("f98") + 0.5,
         f"'{_lbl('f98')}' cannot be less than R2→loading start + loading→R3")
    rule(("f84",), lambda: 0 < g("f84") <= 24,
         f"'{_lbl('f84')}' is monthly hours — e.g. 8 h/day × 30 days = 240, not 8")
    for k, f in FIELDS.items():
        if ("Lakhs" in f["label"]) and g(k) is not None and g(k) >= 100000:
            errs.append(({k}, f"'{f['label']}' = {g(k):g} looks like Rupees — enter in Lakhs"))
    return errs


# ── unit guard against the location's previous month ─────────────────────────
def _unit_factors(k):
    f = FIELDS[k]; lab = f["label"]
    if f["type"] not in ("number", "int") or f.get("auto"):
        return ()
    if f.get("max") == 100 or "%" in lab:
        return (2,)
    if "Lakhs" in lab:
        return (2, 5)
    if "(Rs)" in lab:
        return (5,)
    if any(u in lab for u in ("(KL)", "(MT)", "Rs/MT", "KWH", "KVA")):
        return (3,)
    if "SQM" in lab:
        return (2,)
    return ()


def unit_guard_errors(vals: dict, prev: dict) -> list:
    """Flag values off by an exact unit factor (x100, x1000, x100000) from last month."""
    errs = []
    if not prev:
        return errs
    for k in FIELDS:
        facs = _unit_factors(k)
        a, b = num(vals.get(k)), num(prev.get(k))
        if not facs or not a or not b or a <= 0 or b <= 0:
            continue
        r = math.log10(a / b)
        for e in facs:
            if abs(abs(r) - e) <= 0.3:
                errs.append(({k}, f"'{FIELDS[k]['label']}' = {a:g} is about {10 ** e:g}× "
                                  f"{'more' if r > 0 else 'less'} than last month ({b:g}) — check the unit"))
                break
    return errs


def prev_month_key(month_year: str) -> str:
    d = datetime.strptime(month_year, "%b-%Y")
    y, m = (d.year, d.month - 1) if d.month > 1 else (d.year - 1, 12)
    return datetime(y, m, 1).strftime("%b-%Y")


def validate_all(vals: dict, prev: dict = None, only_keys=None, skip_keys=()) -> list:
    """Every problem in a full MIS record (values, cross-field, unit guard) as messages.
    only_keys: report only rules touching these fields (e.g. the section being saved).
    skip_keys: fields not applicable to this location type."""
    msgs = []
    for k, f in FIELDS.items():
        if f.get("auto") or k in skip_keys or (only_keys and k not in only_keys):
            continue
        v = vals.get(k)
        if v is None or str(v).strip() == "":
            continue
        _, p = check_field(f, v)
        if p:
            msgs.append(f"{f['label']}: '{v}' {p}")
    for keys, m in cross_field_errors(vals) + unit_guard_errors(vals, prev or {}):
        if keys & set(skip_keys):
            continue
        if only_keys and not (keys & set(only_keys)):
            continue
        msgs.append(m)
    return msgs
