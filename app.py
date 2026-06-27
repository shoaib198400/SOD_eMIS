"""HPCL SOD MIS — Main Streamlit Application. Build: 20260614-1."""

import base64
import os
from datetime import date

import pandas as pd
import streamlit as st

# ── Corporate-network SSL bypass for httpx (used by google-genai / Gemini) ────
try:
    import httpx as _httpx
    _orig_client = _httpx.Client.__init__
    def _client_no_ssl(self, *a, **kw):
        kw['verify'] = False
        _orig_client(self, *a, **kw)
    _httpx.Client.__init__ = _client_no_ssl

    _orig_async = _httpx.AsyncClient.__init__
    def _async_no_ssl(self, *a, **kw):
        kw['verify'] = False
        _orig_async(self, *a, **kw)
    _httpx.AsyncClient.__init__ = _async_no_ssl
except Exception:
    pass
# ─────────────────────────────────────────────────────────────────────────────

import sheets
from form_defs import SECTION_FIELDS, SECTION_NAMES, S5_FIELDS

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="HPCL SOD — MIS Portal",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

HPCL_BLUE = "#001F5E"
HPCL_RED  = "#E53935"
HPCL_GOLD = "#C6A64A"

SECTIONS = [
    (1,  "Operations"),      (2,  "Facilities & Planning"),
    (3,  "S&D"),             (4,  "Biofuel"),
    (5,  "M&I"),             (6,  "HSE"),
    (7,  "Operational Efficiency"), (8,  "EM Lock"),
    (9,  "Transportation"),  (10, "Others"),
]

STATUS_META = {
    "NOT_STARTED":    ("⚪", "#8c9db5", "Not Started"),
    "IN_PROGRESS":    ("🔵", "#0d6efd", "In Progress"),
    "PENDING_REVIEW": ("🟡", "#f59e0b", "Pending Review"),
    "SUBMITTED":      ("✅", "#198754", "Submitted"),
    "LOCKED":         ("🔒", "#495057", "Locked"),
    "REJECTED":       ("❌", "#dc3545", "Rejected"),
}


# ── Image loader (cached) ─────────────────────────────────────────────────────

@st.cache_resource
def _assets() -> dict:
    d = os.path.join(os.path.dirname(__file__), "assets")
    out = {}
    for key, fname, mime in [
        ("left_panel",    "left_panel.png",    "image/png"),
        ("dh_logo",       "logo.png",          "image/png"),
        ("hpcl_dh",       "hpcl_dh.jpg",       "image/jpeg"),
        ("side_logo",          "side_logo.png",          "image/png"),
        ("side_panel_banner",  "side_logo.png",          "image/png"),
        ("title_banner",       "title_banner.png",       "image/png"),
        ("login_bg",      "login_bg.png",      "image/png"),
    ]:
        try:
            with open(os.path.join(d, fname), "rb") as f:
                out[key] = f"data:{mime};base64,{base64.b64encode(f.read()).decode()}"
        except Exception:
            out[key] = None
    return out


def _img(key: str, h: int, style: str = "") -> str:
    src = _assets().get(key)
    return f'<img src="{src}" style="height:{h}px;{style}">' if src else ""


def _fy() -> str:
    t = date.today()
    y = t.year
    return f"{y}-{str(y+1)[2:]}" if t.month >= 4 else f"{y-1}-{str(y)[2:]}"


# ── Gemini AI Assistant ───────────────────────────────────────────────────────

def _gemini_response(question: str, user: dict, history: list) -> str:
    """Send a question to Gemini Flash and return the answer."""
    api_key = st.secrets.get("GEMINI_API_KEY", "")
    if not api_key:
        return "⚠️ Gemini API key not configured. Please ask the administrator to add it to secrets.toml."
    try:
        from google import genai
        from google.genai import types

        role_scope = {
            "Admin":   "You can answer questions about all zones and all 111 locations across India.",
            "Zone":    f"The user manages {user.get('zone', 'their zone')}. Scope answers to that zone unless asked otherwise.",
            "Maker":   f"The user works at {user.get('locName', 'their depot')} ({user.get('userId', '')}). Scope answers to their location.",
            "Checker": f"The user works at {user.get('locName', 'their depot')} ({user.get('userId', '')}). Scope answers to their location.",
        }.get(user.get("role", ""), "")

        system_prompt = f"""You are MIS Assistant for HPCL SOD (Supply Operations & Distribution).
You help users query and understand MIS (Monthly Information System) data for petroleum depots.

User role: {user.get('role', 'Unknown')}
{role_scope}

Key MIS parameters (135 total, most important listed):
- f1: MS (Motor Spirit/Petrol) throughput in MT
- f2: HSD (High Speed Diesel) throughput in MT
- f3: Total throughput MT (f1+f2)
- f4: Throughput target MT
- f5: MEB (Maintenance Expenditure Budget) actual in Rupees Lakhs
- f6: MEB % vs Budget
- f7: OPEX Rs/MT (Operating Expenditure per metric tonne)
- f12: Electricity expenditure Rupees Lakhs
- f21: SEC KWH/MT (Specific Energy Consumption, auto-calculated)
- f22: Stock Loss details (MS, HSD, ATF product-wise %)
- f24: AIM Holds count
- f26: Auto-Reco % (Automatic Reconciliation)
- f38/f39: CAPEX actual vs AOP (Annual Operating Plan) in Rupees Lakhs
- f46/f47: MDP target vs actual
- f50: EBP % (Ethanol Blending Programme)
- f54: M&I Index (Maintenance & Inspection)
- f55: PM % (Planned Maintenance percentage)
- f59: HSE Index (Health, Safety & Environment)
- f61: SWC KL/MT (Specific Water Consumption, auto-calculated)

MIS Rules:
- Submission deadline: 5th of each month
- Compliant = submitted on or before the 5th
- Role hierarchy: Maker fills data, Checker reviews, Zone monitors, HQO Admin oversees all India
- 111 locations, 16 zones across India

Be concise and professional. If asked for specific live data you do not have access to in this
conversation, say so clearly and tell the user which parameter (f-number) to look up.
Format tables using markdown when comparing multiple locations or months."""

        client = genai.Client(api_key=api_key)

        # Build contents list from history + new question
        contents = []
        for msg in history:
            role = "user" if msg["role"] == "user" else "model"
            contents.append(types.Content(role=role, parts=[types.Part(text=msg["content"])]))
        contents.append(types.Content(role="user", parts=[types.Part(text=question)]))

        response = client.models.generate_content(
            model="gemini-1.5-flash",
            contents=contents,
            config=types.GenerateContentConfig(system_instruction=system_prompt),
        )
        return response.text

    except Exception as e:
        err = str(e)
        if "API_KEY_INVALID" in err or "invalid" in err.lower() or "400" in err:
            return "⚠️ Gemini API key appears to be invalid. Please generate a fresh key from aistudio.google.com and update secrets.toml."
        return f"⚠️ Could not reach Gemini: {err}"


# ── Phase-3 form helpers ──────────────────────────────────────────────────────

def _sk(month_year: str, key: str) -> str:
    """Session-state key for a form field."""
    return f"draft_{month_year.replace('-', '_')}_{key}"


def _compute_auto(expr: str, vals: dict):
    """Evaluate expressions like 'f17+f18+f19' using current field values."""
    try:
        s = expr
        for k in sorted(vals, key=len, reverse=True):
            if k not in s:
                continue
            v = vals[k]
            if v is None:
                return None
            s = s.replace(k, str(float(v)))
        return float(eval(s))  # noqa: S307 — expr is from form_defs, not user input
    except Exception:
        return None


def _load_draft_to_session(user_id: str, month_year: str, section_num: int):
    """Populate session state from saved draft (once per section/month per rerun cycle)."""
    month_clean = month_year.replace("-", "_")
    flag = f"draft_loaded_{month_clean}_s{section_num}"
    if st.session_state.get(flag):
        # If flag is set but all fields are empty, the values were lost — force reload
        _has_any = any(
            st.session_state.get(_sk(month_year, f["key"])) not in (None, "")
            for f in SECTION_FIELDS.get(section_num, [])
            if not f.get("auto")
        )
        if _has_any:
            return
        st.session_state.pop(flag, None)
    draft = sheets.load_draft(user_id, month_year)
    for f in SECTION_FIELDS.get(section_num, []):
        if f.get("auto"):
            continue
        sk  = _sk(month_year, f["key"])
        cur = st.session_state.get(sk)
        # Skip only when the session already holds a real (non-empty) value —
        # don't skip on None/""  because Streamlit initialises widgets to None
        # and we must overwrite that with the saved draft value.
        if cur is not None and cur != "":
            continue
        raw = draft.get(f["key"])
        if raw is None:
            continue
        try:
            if f["type"] == "int":
                st.session_state[sk] = int(raw)
            elif f["type"] == "number":
                st.session_state[sk] = float(raw)
            elif f["type"] == "date":
                from datetime import datetime as _dts
                st.session_state[sk] = _dts.strptime(str(raw).strip(), "%d/%m/%Y").date()
            else:
                st.session_state[sk] = str(raw)
        except (TypeError, ValueError):
            pass
    st.session_state[flag] = True


def _field_help(field: dict) -> str:
    """Build tooltip text in the official Data Rules format for hover display."""
    if field.get("auto"):
        return f"Auto-calculated  ·  {field.get('hint','')}"

    rule_parts = []
    rule_parts.append("required" if field.get("req") else "optional")

    t = field["type"]
    if t == "int":
        rule_parts.append("integer")
        mn = field.get("min")
        mx = field.get("max")
        if mn is not None and mn > 0:
            rule_parts.append(f"min={mn} (positive)")
        elif mn is not None:
            rule_parts.append(f"min={mn}")
        if mx is not None:
            rule_parts.append(f"max={mx}")
    elif t == "number":
        dec = field.get("dec", 2)
        rule_parts.append("number")
        mn = field.get("min")
        mx = field.get("max")
        if mn is not None and mn > 0:
            rule_parts.append("only positive numbers")
        elif mn is not None:
            rule_parts.append(f"min={mn}")
        if mx is not None:
            rule_parts.append(f"max={mx}")
        rule_parts.append(f"upto {dec} decimals")
    elif t == "select":
        opts = field.get("opts") or []
        rule_parts.append(f"select: {' / '.join(opts)}")
    elif t == "textarea":
        rule_parts.append("text; max 750 characters")
    elif t == "date":
        rule_parts.append("date; DD/MM/YYYY")

    rule_str = "; ".join(rule_parts)
    hint = field.get("hint", "")
    if hint:
        return f"{rule_str}  ·  {hint}"
    return rule_str


def _render_field(field: dict, sk: str, disabled: bool, all_vals: dict):
    """Render one MIS form field widget."""
    ftype   = field["type"]
    is_auto = bool(field.get("auto"))
    hint    = field.get("hint", "")
    tooltip = _field_help(field)
    label   = field["label"] + (" *" if field.get("req") else "")

    if is_auto:
        val = all_vals.get(field["key"])
        dec = field.get("dec") or 0
        if val is not None:
            disp = f"{val:.{dec}f}"
        else:
            disp = "—"
        st.markdown(
            f'<div class="auto-box">'
            f'<div class="auto-label">{field["label"]}</div>'
            f'<div class="auto-val">{disp}</div>'
            f'<div class="auto-hint">🔄 {hint}</div>'
            f'</div>',
            unsafe_allow_html=True,
        )
        return

    if ftype == "select":
        opts = field.get("opts") or []
        if sk in st.session_state and st.session_state[sk] not in opts:
            del st.session_state[sk]
        st.selectbox(label, opts, key=sk, disabled=disabled, help=tooltip)

    elif ftype == "date":
        from datetime import date as _date_t
        # Coerce string "DD/MM/YYYY" stored in session/draft to a date object
        _sv = st.session_state.get(sk)
        if isinstance(_sv, str) and _sv:
            try:
                from datetime import datetime as _dt_t
                st.session_state[sk] = _dt_t.strptime(_sv, "%d/%m/%Y").date()
            except Exception:
                st.session_state.pop(sk, None)
        st.date_input(label, value=None, format="DD/MM/YYYY",
                      key=sk, disabled=disabled, help=tooltip)

    elif ftype == "textarea":
        st.text_area(label, key=sk, disabled=disabled, help=tooltip,
                     max_chars=750, height=110)

    elif ftype == "int":
        min_v = int(field["min"]) if field["min"] is not None else None
        max_v = int(field["max"]) if field["max"] is not None else None
        # Coerce string values (e.g. from Excel upload) to int
        _sv = st.session_state.get(sk)
        if isinstance(_sv, str):
            try:
                st.session_state[sk] = int(float(_sv)) if _sv.strip() else None
            except Exception:
                st.session_state.pop(sk, None)
        st.number_input(label, value=None, step=1,
                        min_value=min_v, max_value=max_v,
                        key=sk, disabled=disabled, help=tooltip)

    else:  # "number" → float
        dec   = field.get("dec") or 2
        min_v = float(field["min"]) if field["min"] is not None else None
        max_v = float(field["max"]) if field["max"] is not None else None
        step  = round(10.0 ** (-dec), dec)
        # Coerce string values (e.g. from Excel upload) to float
        _sv = st.session_state.get(sk)
        if isinstance(_sv, str):
            try:
                st.session_state[sk] = round(float(_sv), dec) if _sv.strip() else None
            except Exception:
                st.session_state.pop(sk, None)
        st.number_input(label, value=None, step=step,
                        format=f"%.{dec}f",
                        min_value=min_v, max_value=max_v,
                        key=sk, disabled=disabled, help=tooltip)


def _do_save(user_id: str, month_year: str, section_num: int,
             fields: list, is_locked: bool) -> dict:
    """Collect session-state values and persist draft."""
    if is_locked:
        return {"ok": False, "msg": "Section is locked."}
    from datetime import date as _date_cls
    field_data, all_req_filled = {}, True
    for f in fields:
        sk  = _sk(month_year, f["key"])
        val = st.session_state.get(sk)
        if isinstance(val, _date_cls):
            val = val.strftime("%d/%m/%Y")
        if val is not None and val != "":
            field_data[f["key"]] = val
        elif f.get("req") and not f.get("auto"):
            all_req_filled = False
    # Cross-field: complied recommendations cannot exceed total
    if section_num == 1 and all_req_filled:
        try:
            if int(field_data.get("f161") or 0) > int(field_data.get("f160") or 0):
                all_req_filled = False
        except (ValueError, TypeError):
            pass

    return sheets.save_draft(user_id, month_year, section_num,
                             field_data, mark_complete=all_req_filled)


# ── Global CSS ────────────────────────────────────────────────────────────────

def _base_css():
    st.markdown("""
    <style>
    /* ── Global font (system stack — no external fetch, no render block) ── */
    html, body, [class*="css"], [data-testid], .stApp, section.main {
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI',
                     Roboto, Helvetica, Arial, sans-serif !important;
    }

    #MainMenu, footer, header { visibility: hidden; }
    section.main { padding:0 !important; }
    section.main > .block-container { padding:0 !important; max-width:100% !important; }

    /* ── Global text sizing (~25% smaller than original) ── */
    body, p, span, div, li { font-size:12px !important; line-height:1.55 !important; }

    /* ── Sidebar text ── */
    [data-testid="stSidebar"] * { font-size:11px !important; color:#ffffff !important; }
    [data-testid="stSidebar"] .stButton button { font-size:11px !important; font-weight:600 !important; color:#ffffff !important; }

    /* ── PRIMARY buttons (Login, Save, Submit, etc.) ── */
    .stButton button[kind="primary"],
    button[data-testid="baseButton-primary"],
    .stFormSubmitButton button {
        background: linear-gradient(135deg, #E53935 0%, #C62828 100%) !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 10px !important;
        font-size: 15px !important;
        font-weight: 700 !important;
        padding: 10px 22px !important;
        letter-spacing: 0.3px !important;
        box-shadow: 0 4px 14px rgba(198,40,40,0.28) !important;
        transition: all 0.2s ease !important;
    }
    .stButton button[kind="primary"] *,
    button[data-testid="baseButton-primary"] *,
    .stFormSubmitButton button * {
        color: #ffffff !important;
    }
    .stButton button[kind="primary"]:hover,
    button[data-testid="baseButton-primary"]:hover,
    .stFormSubmitButton button:hover {
        background: linear-gradient(135deg, #EF5350 0%, #D32F2F 100%) !important;
        box-shadow: 0 6px 20px rgba(198,40,40,0.38) !important;
        transform: translateY(-1px) !important;
    }

    /* ── SECONDARY buttons (Back, Cancel, etc.) ── */
    .stButton button[kind="secondary"],
    button[data-testid="baseButton-secondary"] {
        background: #1565C0 !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 10px !important;
        font-weight: 600 !important;
        padding: 9px 18px !important;
        box-shadow: 0 2px 8px rgba(21,101,192,0.25) !important;
    }
    .stButton button[kind="secondary"] *,
    button[data-testid="baseButton-secondary"] * {
        color: #ffffff !important;
    }
    .stButton button[kind="secondary"]:hover,
    button[data-testid="baseButton-secondary"]:hover {
        background: #1976D2 !important;
        color: #ffffff !important;
        box-shadow: 0 4px 14px rgba(21,101,192,0.38) !important;
    }

    /* ── Input labels — target label itself + inner p/span Streamlit wraps ── */
    [data-testid="stNumberInput"] label,
    [data-testid="stNumberInput"] label p,
    [data-testid="stNumberInput"] label span,
    [data-testid="stTextInput"]   label,
    [data-testid="stTextInput"]   label p,
    [data-testid="stTextInput"]   label span,
    [data-testid="stSelectbox"]   label,
    [data-testid="stSelectbox"]   label p,
    [data-testid="stSelectbox"]   label span,
    [data-testid="stTextArea"]    label,
    [data-testid="stTextArea"]    label p,
    [data-testid="stTextArea"]    label span,
    [data-testid="stDateInput"]   label,
    [data-testid="stDateInput"]   label p,
    [data-testid="stDateInput"]   label span {
        font-size:13px !important; font-weight:800 !important;
        letter-spacing:0.2px !important; color:#001a6e !important;
        text-transform:none !important; margin-bottom:4px !important;
    }

    /* ── Checkbox — darker border so it's clearly visible ── */
    [data-testid="stCheckbox"] > label > div:first-child {
        border:2px solid #001a6e !important;
        border-radius:4px !important;
    }
    [data-testid="stCheckbox"] > label > div:first-child:hover {
        border-color:#0033A0 !important;
        box-shadow:0 0 0 2px rgba(0,51,160,0.18) !important;
    }

    /* ── File uploader — fix overlapping "uploadUpload" text ── */
    [data-testid="stFileUploader"] section {
        background: #f0f5ff !important; border: 2px dashed #0033A0 !important;
        border-radius: 10px !important; padding: 16px !important;
    }
    [data-testid="stFileUploader"] section > button {
        background: #001a6e !important; color: white !important;
        border-radius: 8px !important; font-weight: 700 !important;
        border: none !important;
    }
    [data-testid="stFileUploader"] section > button > div { display:none !important; }
    [data-testid="stFileUploader"] section > button::after {
        content: "Browse Files" !important;
        color: white !important; font-weight: 700 !important;
    }
    [data-testid="stFileUploader"] section > span { display:none !important; }

    /* ── Input boxes ── */
    .stTextInput > div > div > input,
    .stNumberInput > div > div > input {
        border: 1.5px solid #d0d8ec !important; border-radius: 9px !important;
        padding: 10px 14px !important; font-size: 15px !important;
        background: #f7f9fc !important; transition: all 0.15s !important;
    }
    .stTextInput > div > div > input:focus,
    .stNumberInput > div > div > input:focus {
        border-color: #001F5E !important; background: white !important;
        box-shadow: 0 0 0 3px rgba(0,31,94,0.12) !important;
    }

    /* ── Selectbox ── */
    .stSelectbox > div > div {
        font-size:15px !important; border-radius:9px !important;
        border: 1.5px solid #d0d8ec !important;
    }

    /* ── Textarea ── */
    .stTextArea textarea {
        font-size:15px !important; border-radius:9px !important;
        border:1.5px solid #d0d8ec !important; padding:10px 14px !important;
    }

    /* ── Expanders ── */
    [data-testid="stExpander"] summary,
    .streamlit-expanderHeader {
        font-size:13px !important; color:#001F5E !important;
        font-weight:600 !important;
    }

    /* ── Metrics ── */
    [data-testid="stMetric"] {
        background: white !important; border-radius: 12px !important;
        padding: 14px 18px !important;
        box-shadow: 0 2px 10px rgba(0,31,94,0.08) !important;
        border-top: 3px solid #001F5E !important;
    }
    [data-testid="stMetric"] label {
        font-size:12px !important; font-weight:700 !important;
        color:#666 !important; text-transform:uppercase !important; letter-spacing:0.5px !important;
    }
    [data-testid="stMetric"] [data-testid="stMetricValue"] {
        font-size:28px !important; font-weight:800 !important; color:#001F5E !important;
    }

    /* ── Dataframe / table (HTML table rendering) ── */
    [data-testid="stDataFrame"] td,
    [data-testid="stDataFrame"] th {
        font-size:13px !important; padding:8px 14px !important;
        text-align:center !important; vertical-align:middle !important;
    }
    [data-testid="stDataFrame"] th {
        background: linear-gradient(135deg,#001060 0%,#002b8f 100%) !important;
        color: white !important; font-weight:700 !important;
        font-size:12px !important; letter-spacing:0.5px !important;
        text-transform:uppercase !important;
    }
    [data-testid="stDataFrame"] tbody tr:nth-child(even) td {
        background:#f0f4ff !important;
    }
    [data-testid="stDataFrame"] tbody tr:hover td {
        background:#dce6f7 !important;
    }

    /* ── data_editor AG Grid styling (covers all Streamlit AG Grid themes) ── */
    .ag-root-wrapper,
    .ag-theme-alpine .ag-root-wrapper,
    .ag-theme-streamlit .ag-root-wrapper {
        border-radius:10px !important; overflow:hidden !important;
        border:1.5px solid #b0c8f8 !important;
        box-shadow:0 2px 12px rgba(0,26,110,0.10) !important;
    }
    .ag-header,
    .ag-theme-alpine .ag-header,
    .ag-theme-streamlit .ag-header {
        background:linear-gradient(90deg,#001a6e 0%,#0033A0 100%) !important;
        border-bottom:2px solid #4a80d0 !important;
    }
    .ag-header-row,
    .ag-theme-alpine .ag-header-row,
    .ag-theme-streamlit .ag-header-row { background:transparent !important; }
    .ag-header-cell,
    .ag-theme-alpine .ag-header-cell,
    .ag-theme-streamlit .ag-header-cell {
        background:transparent !important;
        border-right:1px solid rgba(255,255,255,0.18) !important;
    }
    .ag-header-cell-text,
    .ag-theme-alpine .ag-header-cell-text,
    .ag-theme-streamlit .ag-header-cell-text {
        color:#ffffff !important; font-weight:700 !important;
        font-size:11.5px !important; letter-spacing:0.3px !important;
    }
    .ag-header-cell-label,
    .ag-theme-alpine .ag-header-cell-label,
    .ag-theme-streamlit .ag-header-cell-label { justify-content:center !important; }
    .ag-header-icon, .ag-header-icon svg,
    .ag-theme-alpine .ag-header-icon, .ag-theme-alpine .ag-header-icon svg,
    .ag-theme-streamlit .ag-header-icon, .ag-theme-streamlit .ag-header-icon svg {
        color:#cfe0ff !important; fill:#cfe0ff !important;
    }
    .ag-sort-indicator-icon,
    .ag-theme-alpine .ag-sort-indicator-icon,
    .ag-theme-streamlit .ag-sort-indicator-icon { color:#cfe0ff !important; }
    /* Row zebra & hover */
    .ag-row-odd, .ag-theme-alpine .ag-row-odd, .ag-theme-streamlit .ag-row-odd
        { background:#f0f5ff !important; }
    .ag-row-even, .ag-theme-alpine .ag-row-even, .ag-theme-streamlit .ag-row-even
        { background:#ffffff !important; }
    .ag-row:hover, .ag-theme-alpine .ag-row:hover, .ag-theme-streamlit .ag-row:hover
        { background:#dce8ff !important; }
    /* Center data cells */
    .ag-cell,
    .ag-theme-alpine .ag-cell,
    .ag-theme-streamlit .ag-cell {
        text-align:center !important; font-size:13px !important;
        display:flex !important; align-items:center !important;
        justify-content:center !important;
    }
    .ag-full-width-row, .ag-theme-alpine .ag-full-width-row,
    .ag-theme-streamlit .ag-full-width-row { background:#f8faff !important; }
    .ag-header-cell-resize::after, .ag-theme-alpine .ag-header-cell-resize::after,
    .ag-theme-streamlit .ag-header-cell-resize::after { background:#7aabff !important; }

    /* ── Alerts ── */
    [data-testid="stAlert"] {
        font-size:14px !important; border-radius:10px !important;
        border-left-width: 4px !important;
    }

    /* ── Captions ── */
    .stCaption, [data-testid="stCaptionContainer"] {
        font-size:13px !important; color:#666 !important;
    }

    /* ── Progress bar ── */
    .stProgress > div > div {
        background:#dce6f7 !important; border-radius:6px !important; height:10px !important;
    }
    .stProgress > div > div > div {
        background: linear-gradient(90deg, #E53935, #1565C0) !important;
        border-radius:6px !important;
    }

    /* ── Section cards ── */
    .sec-card {
        background: white; border-radius: 12px; padding: 14px 18px; margin-bottom: 10px;
        box-shadow: 0 2px 10px rgba(0,31,94,0.08); display:flex;
        justify-content:space-between; align-items:center;
        border-left: 4px solid #dee2e6;
        transition: box-shadow 0.2s;
    }
    .sec-card:hover { box-shadow: 0 4px 18px rgba(0,31,94,0.14); }
    .sec-badge {
        font-size:12px !important; font-weight:700; color:white;
        padding:4px 12px; border-radius:20px; white-space:nowrap;
        letter-spacing:0.3px;
    }

    /* ── File uploader ── */
    [data-testid="stFileUploader"] label { font-size:14px !important; font-weight:600 !important; }
    [data-testid="stFileUploader"] p    { font-size:14px !important; }

    /* ── Tabs — base (overridden per-page in _dashboard_css) ── */
    .stTabs [data-baseweb="tab"] {
        font-size:14px !important; font-weight:600 !important;
        padding: 8px 18px !important;
    }

    /* ── Auto-calc field boxes ── */
    .auto-box {
        background: linear-gradient(135deg, #f0f4ff, #e8eefa);
        border: 1px solid #c8d4f0; border-radius: 10px;
        padding: 10px 14px; margin-bottom: 8px;
    }
    .auto-label { font-size:12px; font-weight:700; color:#001F5E; text-transform:uppercase; letter-spacing:0.4px; }
    .auto-val   { font-size:22px; font-weight:800; color:#001F5E; margin:2px 0; }
    .auto-hint  { font-size:11px; color:#6b7a99; }

    </style>
    """, unsafe_allow_html=True)

    # Inject JS via components.html so it actually executes (st.markdown strips scripts)
    import streamlit.components.v1 as _comp
    _comp.html("""
    <script>
    (function applyLogoutStyle() {
        var pd = window.parent.document;
        function styleLogout() {
            pd.querySelectorAll('button').forEach(function(btn) {
                var txt = (btn.innerText || btn.textContent || '').trim();
                if (txt.includes('Logout')) {
                    btn.style.setProperty('background', '#E53935', 'important');
                    btn.style.setProperty('color', 'white', 'important');
                    btn.style.setProperty('border-radius', '24px', 'important');
                    btn.style.setProperty('border', '2px solid #dc3545', 'important');
                    btn.style.setProperty('font-weight', '700', 'important');
                    btn.style.setProperty('padding', '4px 20px', 'important');
                }
            });
        }
        styleLogout();
        var obs = new MutationObserver(styleLogout);
        obs.observe(pd.body, { childList: true, subtree: true });
    })();
    </script>
    """, height=0)


def _login_css(bg_b64: str = ""):
    bg_css = (
        f'url("{bg_b64}") left top / cover no-repeat fixed'
        if bg_b64 else
        "linear-gradient(160deg,#001F5E 0%,#003087 100%)"
    )
    st.markdown(f"""
    <style>
    /* ── Hide Streamlit chrome ── */
    [data-testid="stHeader"], [data-testid="stSidebar"],
    [data-testid="collapsedControl"], #MainMenu, footer {{ display:none !important; }}

    /* ── Background on html/body ── */
    html, body {{
        background: {bg_css} !important;
        overflow: hidden !important;
        height: 100vh !important;
        margin: 0 !important; padding: 0 !important;
    }}

    /* ── Every Streamlit wrapper layer: transparent, no layout ── */
    .stApp,
    [data-testid="stAppViewContainer"],
    [data-testid="stAppViewContainer"] > section,
    section.main,
    [data-testid="stMain"] {{
        background: transparent !important;
        overflow: visible !important;
        height: 100vh !important;
        padding: 0 !important; margin: 0 !important;
    }}

    /* ── Login card — fixed to the RIGHT of the background image text.
       Left edge anchored at 68vw so it never overlaps the "SOD e-MIS Portal"
       title text regardless of viewport width or browser zoom level. ── */
    div[class*="block-container"] {{
        position: fixed !important;
        left: 68vw !important;
        right: 1vw !important;
        top: 4vh !important;
        bottom: 10vh !important;
        width: auto !important;
        max-width: 400px !important;
        min-width: 270px !important;
        background: #ffffff !important;
        border-top: 5px solid #E53935 !important;
        border-radius: 14px !important;
        overflow-y: auto !important;
        overflow-x: hidden !important;
        padding: 0 18px 12px !important;
        margin: 0 !important;
        z-index: 200 !important;
        box-shadow: 0 12px 56px rgba(0,20,80,0.38) !important;
    }}

    /* ── Suppress default vertical-block gap ── */
    div[data-testid="stVerticalBlock"] {{
        gap: 0 !important; padding: 0 !important;
    }}
    .element-container {{ margin-bottom: 2px !important; }}

    /* ── Inner horizontal blocks (Remember me / Forgot row) ── */
    div[data-testid="stHorizontalBlock"] {{
        height: auto !important; min-height: 0 !important; gap: 4px !important;
    }}

    /* ── Form input labels ── */
    [data-testid="stForm"] label {{
        font-size: 11px !important; font-weight: 700 !important;
        color: #001F5E !important; letter-spacing: 0.8px !important;
        text-transform: uppercase !important;
    }}

    /* ── Input fields ── */
    [data-testid="stForm"] input {{
        border: 2px solid #c0cce8 !important; border-radius: 8px !important;
        padding: 10px 14px !important; font-size: 13px !important;
        color: #001F5E !important; background: #ffffff !important;
        font-weight: 500 !important;
    }}
    [data-testid="stForm"] input:focus {{
        border-color: #E53935 !important;
        box-shadow: 0 0 0 3px rgba(229,57,53,0.12) !important;
        background: #fff8f8 !important;
    }}
    [data-testid="stForm"] input::placeholder {{
        color: rgba(0,31,94,0.35) !important; font-size: 12px !important;
    }}

    /* ── Login button ── */
    [data-testid="stForm"] button[kind="primaryFormSubmit"] {{
        background: linear-gradient(135deg,#001640 0%,#001F5E 40%,#003087 80%,#0044bb 100%) !important;
        color: white !important; border: none !important; border-radius: 9px !important;
        font-weight: 800 !important; font-size: 13px !important;
        letter-spacing: 2px !important; text-transform: uppercase !important;
        padding: 12px 20px !important; width: 100% !important;
        box-shadow: 0 4px 18px rgba(0,31,94,0.40) !important;
    }}

    /* ── Checkbox (Remember Me) — force visible against any theme ── */
    .stCheckbox label, .stCheckbox label p, .stCheckbox span {{
        font-size: 12px !important; color: #001F5E !important;
        font-weight: 600 !important; opacity: 1 !important;
    }}
    .stCheckbox {{ background: transparent !important; }}
    [data-testid="InputInstructions"] {{ display: none !important; }}

    /* ── Expander header styling ── */
    [data-testid="stExpander"] summary {{
        font-size: 13px !important; font-weight: 600 !important;
        color: #001F5E !important;
        display: flex !important; align-items: center !important;
    }}
    [data-testid="stExpander"] details summary > p {{
        font-size: 13px !important; margin: 0 !important;
    }}
    </style>
    """, unsafe_allow_html=True)


def _dashboard_css():
    st.markdown("""
    <style>
    /* ── Fonts ── */
    html, body, [class*="css"], button, input, select {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif !important;
    }

    /* ── App & main background ── */
    .stApp,
    [data-testid="stAppViewContainer"],
    [data-testid="stAppViewContainer"] > .main {
        background: linear-gradient(160deg, #edf1fa 0%, #e4eaf6 50%, #eef2f7 100%) !important;
    }

    /* ── Override Streamlit's header-height CSS variable ── */
    :root { --header-height: 0px !important; }

    /* ── Hide header elements completely, out of flow ── */
    [data-testid="stHeader"],
    [data-testid="stDecoration"],
    [data-testid="stToolbar"],
    #MainMenu, footer {
        display:    none !important;
        height:     0    !important;
        min-height: 0    !important;
        overflow:   hidden !important;
    }

    /* ── Zero padding on every wrapper above the block-container ── */
    .stApp,
    .stApp > section,
    [data-testid="stAppViewContainer"],
    [data-testid="stAppViewContainer"] > section,
    section[data-testid="stMain"],
    section.main {
        padding-top: 0 !important;
        margin-top:  0 !important;
    }

    /* ── Block-container: broadest selector, horizontal padding only ── */
    .block-container,
    div.block-container {
        padding-top:    0 !important;
        padding-bottom: 1rem !important;
        padding-left:   1.4rem !important;
        padding-right:  1.4rem !important;
        max-width:      100% !important;
        margin-top:     0 !important;
    }
    div[data-testid="stVerticalBlock"] {
        gap:         0.6rem !important;
        padding-top: 0 !important;
        margin-top:  0 !important;
    }

    /* ── Sidebar: modern spec redesign ── */

    /* Zero padding on all wrapper layers */
    [data-testid="stSidebar"] > div,
    [data-testid="stSidebar"] > div > div,
    [data-testid="stSidebar"] > div > div > div,
    [data-testid="stSidebarContent"],
    [data-testid="stSidebarUserContent"],
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"],
    [data-testid="stSidebar"] section,
    [data-testid="stSidebar"] .block-container {
        padding-top: 0 !important; margin-top: 0 !important;
    }
    [data-testid="stSidebarUserContent"] > div:first-child,
    [data-testid="stSidebarUserContent"] > div:first-child > div:first-child,
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] > div:first-child {
        margin-top: -2rem !important; padding-top: 0 !important;
    }

    /* Container */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #072A86 0%, #041D63 100%) !important;
        min-width: 240px !important; max-width: 240px !important;
        padding: 0 !important;
    }
    /* Strip white from every inner wrapper layer (light-theme default bleeds through) */
    [data-testid="stSidebar"] > div,
    [data-testid="stSidebar"] > div > div,
    [data-testid="stSidebar"] > div > div > div,
    [data-testid="stSidebarContent"],
    [data-testid="stSidebarUserContent"],
    [data-testid="stSidebar"] section,
    [data-testid="stSidebar"] .block-container {
        background: transparent !important;
        background-color: transparent !important;
        padding-top: 0 !important;
        margin-top: 0 !important;
    }
    [data-testid="stSidebar"] > div:first-child { padding: 0 !important; }
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] { gap: 0 !important; padding: 0 !important; }

    /* All sidebar text → white */
    /* All sidebar text → white, Segoe UI, compact */
    [data-testid="stSidebar"] * {
        font-family: 'Segoe UI', system-ui, -apple-system, sans-serif !important;
        font-size: 11.5px !important; color: #ffffff !important;
    }
    [data-testid="stSidebar"] .stButton button {
        font-size: 11.5px !important; font-weight: 500 !important; color: #ffffff !important;
    }

    /* Hide the keyboard_double collapse-arrow artifact */
    [data-testid="collapsedControl"] span[data-testid="stExpanderToggleIcon"],
    [data-testid="collapsedControl"] .material-symbols-rounded {
        display: none !important;
    }

    /* Markdown containers flush + transparent so blue gradient shows through */
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
    [data-testid="stSidebar"] .stMarkdownContainer {
        margin: 0 !important; padding: 0 !important; width: 100% !important;
        background: transparent !important; background-color: transparent !important;
    }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] > div,
    [data-testid="stSidebar"] .stMarkdownContainer > div {
        margin: 0 !important; padding: 0 !important; width: 100% !important;
        background: transparent !important; background-color: transparent !important;
    }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
    [data-testid="stSidebar"] .stMarkdownContainer p { margin: 0 !important; }

    /* Crush vertical gaps inside sidebar */
    [data-testid="stSidebar"] div[data-testid="stVerticalBlock"] {
        gap: 2px !important; padding: 0 !important;
    }
    [data-testid="stSidebar"] .element-container {
        margin-bottom: 0 !important;
        background: transparent !important; background-color: transparent !important;
    }

    /* Nav buttons — glass-card container, rounded, highlight on hover */
    [data-testid="stSidebar"] [data-testid="stButton"] > button {
        width: calc(100% - 16px) !important;
        display: flex !important; flex-direction: row !important;
        align-items: center !important; justify-content: flex-start !important;
        text-align: left !important;
        background: rgba(255,255,255,0.06) !important; color: #ffffff !important;
        border: 1px solid rgba(255,255,255,0.08) !important;
        border-radius: 8px !important;
        margin: 2px 8px !important; padding: 8px 13px !important;
        font-size: 11.5px !important; font-weight: 500 !important;
        white-space: nowrap !important; overflow: hidden !important;
        text-overflow: ellipsis !important;
        box-shadow: none !important; outline: none !important;
        transition: background 0.15s ease, transform 0.12s ease, border-color 0.15s !important;
    }
    [data-testid="stSidebar"] [data-testid="stButton"] > button > div {
        width: 100% !important; display: flex !important;
        justify-content: flex-start !important; text-align: left !important;
    }
    [data-testid="stSidebar"] [data-testid="stButton"] > button p,
    [data-testid="stSidebar"] [data-testid="stButton"] > button span,
    [data-testid="stSidebar"] [data-testid="stButton"] > button div {
        text-align: left !important; width: 100% !important;
        display: block !important; color: #ffffff !important; font-size: 11.5px !important;
    }
    [data-testid="stSidebar"] [data-testid="stButton"] > button:hover {
        background: rgba(255,255,255,0.14) !important; color: white !important;
        border-color: rgba(255,255,255,0.18) !important;
        transform: translateX(2px) !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.18) !important;
    }
    [data-testid="stSidebar"] [data-testid="stButton"] > button:focus,
    [data-testid="stSidebar"] [data-testid="stButton"] > button:active,
    [data-testid="stSidebar"] [data-testid="stButton"] > button:focus-visible {
        background: rgba(255,255,255,0.10) !important;
        box-shadow: none !important; outline: none !important;
    }

    /* Active / disabled → blue gradient pill */
    [data-testid="stSidebar"] [data-testid="stButton"] > button:disabled {
        background: linear-gradient(135deg, #1E63FF 0%, #2D8CFF 100%) !important;
        color: white !important; border: 1px solid rgba(45,140,255,0.40) !important;
        border-radius: 8px !important;
        box-shadow: 0 4px 14px rgba(30,99,255,0.40) !important;
        opacity: 1 !important; font-weight: 700 !important; cursor: default !important;
        transform: none !important;
    }

    /* Force all button kinds in sidebar to the same nav style */
    [data-testid="stSidebar"] .stButton button[kind="secondary"],
    [data-testid="stSidebar"] button[data-testid="baseButton-secondary"],
    [data-testid="stSidebar"] .stButton button {
        background: rgba(255,255,255,0.06) !important; color: #ffffff !important;
        border: 1px solid rgba(255,255,255,0.08) !important; border-radius: 8px !important;
        box-shadow: none !important; font-size: 11px !important;
        margin: 2px 8px !important; width: calc(100% - 16px) !important;
    }
    [data-testid="stSidebar"] .stButton button:hover,
    [data-testid="stSidebar"] .stButton button[kind="secondary"]:hover,
    [data-testid="stSidebar"] button[data-testid="baseButton-secondary"]:hover {
        background: rgba(255,255,255,0.08) !important; color: white !important;
    }

    /* Download button in sidebar — match the nav button style */
    [data-testid="stSidebar"] [data-testid="stDownloadButton"] > button {
        width: calc(100% - 16px) !important;
        display: flex !important; flex-direction: row !important;
        align-items: center !important; justify-content: flex-start !important;
        text-align: left !important;
        background: rgba(255,255,255,0.06) !important; color: #ffffff !important;
        border: 1px solid rgba(255,255,255,0.08) !important;
        border-radius: 8px !important;
        margin: 2px 8px !important; padding: 8px 13px !important;
        font-size: 11.5px !important; font-weight: 500 !important;
        box-shadow: none !important; outline: none !important;
        transition: background 0.15s ease, border-color 0.15s !important;
    }
    [data-testid="stSidebar"] [data-testid="stDownloadButton"] > button * {
        color: #ffffff !important; font-size: 11.5px !important;
        text-align: left !important;
    }
    [data-testid="stSidebar"] [data-testid="stDownloadButton"] > button:hover {
        background: rgba(255,255,255,0.14) !important; color: white !important;
        border-color: rgba(255,255,255,0.18) !important;
        transform: translateX(2px) !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.18) !important;
    }

    /* Sidebar collapse control */
    [data-testid="collapsedControl"] button {
        background: #041D63 !important; color: white !important; border: none !important;
    }

    /* Thin custom scrollbar */
    [data-testid="stSidebar"]::-webkit-scrollbar { width: 3px !important; }
    [data-testid="stSidebar"]::-webkit-scrollbar-track { background: transparent !important; }
    [data-testid="stSidebar"]::-webkit-scrollbar-thumb {
        background: rgba(255,255,255,0.22) !important; border-radius: 2px !important;
    }

    /* ── Primary action buttons: HPCL red gradient ── */
    .stButton button[kind="primary"],
    button[data-testid="baseButton-primary"],
    .stFormSubmitButton > button {
        background: linear-gradient(135deg, #E53935 0%, #C62828 100%) !important;
        color: #ffffff !important; border: none !important; border-radius: 10px !important;
        font-weight: 700 !important; letter-spacing: 0.3px !important;
        box-shadow: 0 4px 14px rgba(198,40,40,0.28) !important;
        transition: all 0.2s ease !important;
    }
    .stButton button[kind="primary"] *,
    button[data-testid="baseButton-primary"] *,
    .stFormSubmitButton > button * { color: #ffffff !important; }
    .stButton button[kind="primary"]:hover,
    button[data-testid="baseButton-primary"]:hover,
    .stFormSubmitButton > button:hover {
        background: linear-gradient(135deg, #EF5350 0%, #D32F2F 100%) !important;
        box-shadow: 0 6px 20px rgba(198,40,40,0.38) !important;
        transform: translateY(-1px) !important;
    }

    /* ── Sidebar active-section div — strip the markdown container's own margins
       so the highlighted div is flush with the surrounding buttons */
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
    [data-testid="stSidebar"] .stMarkdownContainer {
        margin:  0 !important;
        padding: 0 !important;
        width:   100% !important;
    }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] > div,
    [data-testid="stSidebar"] .stMarkdownContainer > div {
        margin:  0 !important;
        padding: 0 !important;
        width:   100% !important;
    }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
    [data-testid="stSidebar"] .stMarkdownContainer p {
        margin: 0 !important;
    }

    /* Tooltip icon — make the ? more prominent */
    [data-testid="stTooltipIcon"] svg {
        width:  16px !important;
        height: 16px !important;
        color:  #4a80e8 !important;
    }

    /* ── General action buttons (non-primary, non-sidebar) ── */
    .stButton > button {
        border-radius: 10px !important; font-weight: 600 !important; font-size: 14px !important;
        transition: all 0.15s ease !important;
    }


    /* ── Selectbox labels ── */
    .stSelectbox label {
        font-weight: 600 !important; font-size: 13px !important; color: #333 !important;
    }

    /* ── Section/form action buttons by keyword (JS-applied via class) ── */
    /* Save Draft → navy blue */
    button[data-action="save"] {
        background: linear-gradient(135deg,#001F5E,#003087) !important;
        color:white !important; border:none !important;
        box-shadow:0 3px 10px rgba(0,31,94,0.25) !important;
    }
    /* Approve / Submit → green */
    button[data-action="approve"] {
        background: linear-gradient(135deg,#15803d,#166534) !important;
        color:white !important; border:none !important;
        box-shadow:0 3px 10px rgba(21,128,61,0.28) !important;
    }

    /* ── Card containers in dashboard ── */
    .dash-card {
        background: white; border-radius: 14px; padding: 18px 22px;
        box-shadow: 0 2px 12px rgba(0,31,94,0.09);
        margin-bottom: 14px; border-top: 3px solid #001F5E;
    }

    /* ── Horizontal rule separator ── */
    hr { border-color: #e0e6f0 !important; }

    /* ── Improve form section spacing ── */
    [data-testid="stForm"] {
        background: white !important; border-radius: 14px !important;
        padding: 18px !important; box-shadow: 0 2px 12px rgba(0,31,94,0.07) !important;
        border: 1px solid #e4eaf6 !important;
    }

    /* ── Tabs in dashboard ── */
    .stTabs [data-baseweb="tab-list"] {
        background: white !important; border-radius: 10px !important;
        padding: 4px !important; gap: 2px !important;
        box-shadow: 0 1px 6px rgba(0,31,94,0.08) !important;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px !important; font-weight:600 !important; font-size:14px !important;
        color: #001F5E !important;
    }
    .stTabs [aria-selected="true"],
    .stTabs [aria-selected="true"] *,
    .stTabs [aria-selected="true"] p,
    .stTabs [aria-selected="true"] span,
    .stTabs [aria-selected="true"] div {
        background: #001F5E !important; color: #ffffff !important;
        border-bottom: none !important;
    }

    /* ── DataFrame / st.table professional styling ── */
    [data-testid="stDataFrame"] div[data-testid="stDataFrameResizable"],
    [data-testid="stTable"] {
        border-radius: 12px !important;
        overflow: hidden !important;
        box-shadow: 0 2px 12px rgba(0,31,94,0.10) !important;
        border: 1px solid #dce6f7 !important;
    }
    /* Header cells */
    [data-testid="stDataFrame"] th,
    [data-testid="stTable"] th {
        background: linear-gradient(135deg,#001060 0%,#002b8f 100%) !important;
        color: white !important; font-weight:700 !important;
        font-size:12px !important; letter-spacing:0.5px !important;
        text-transform:uppercase !important; text-align:center !important;
        padding:10px 14px !important; border:none !important;
    }
    /* Data cells */
    [data-testid="stDataFrame"] td,
    [data-testid="stTable"] td {
        font-size:13px !important; padding:8px 14px !important;
        text-align:center !important; vertical-align:middle !important;
        border-bottom: 1px solid #edf1fb !important; color:#1a2a4a !important;
    }
    /* Alternating row color */
    [data-testid="stDataFrame"] tbody tr:nth-child(even) td,
    [data-testid="stTable"] tbody tr:nth-child(even) td {
        background:#f0f4ff !important;
    }
    [data-testid="stDataFrame"] tbody tr:hover td,
    [data-testid="stTable"] tbody tr:hover td {
        background:#dce6f7 !important;
    }

    /* ── Professional heading hierarchy ── */
    h1 { font-size:24px !important; font-weight:800 !important; color:#001060 !important;
         letter-spacing:-0.3px !important; }
    h2 { font-size:20px !important; font-weight:700 !important; color:#002b8f !important; }
    h3 { font-size:16px !important; font-weight:700 !important; color:#002b8f !important; }
    p  { font-size:14px !important; color:#2d3a52 !important; line-height:1.65 !important; }

    /* ── .mis-tbl — HTML table used for S5A summaries and detail-table read views ── */
    table.mis-tbl {
        width:100%; border-collapse:collapse; font-size:13px;
        margin-bottom:10px; border-radius:10px; overflow:hidden;
        box-shadow:0 2px 10px rgba(0,26,110,0.09);
    }
    table.mis-tbl th {
        background:linear-gradient(135deg,#001060 0%,#002b8f 100%);
        color:#fff; font-weight:700; padding:9px 13px;
        text-align:center; font-size:11.5px;
        letter-spacing:0.4px; text-transform:uppercase; border:none;
    }
    table.mis-tbl td {
        padding:7px 12px; text-align:center;
        border-bottom:1px solid #edf1fb; color:#1a2a4a;
        vertical-align:middle;
    }
    table.mis-tbl tbody tr:nth-child(even) td { background:#f0f4ff; }
    table.mis-tbl tbody tr:hover td { background:#dce6f7; }
    </style>
    """, unsafe_allow_html=True)


# ── Login page ────────────────────────────────────────────────────────────────

def show_login():
    # Guard: if user already authenticated (e.g. mid-rerun after successful login),
    # skip rendering the login UI entirely to prevent a 1-frame flash.
    if st.session_state.get("user") and st.session_state.get("page", "login") != "login":
        return

    # Show session-end banners (cleared after display)
    if st.session_state.pop("_timeout_msg", False):
        st.warning(
            "⏱️  You were automatically logged out after 30 minutes of inactivity.  "
            "Please log in again."
        )
    if st.session_state.pop("_displaced_msg", False):
        st.warning(
            "🔒  Your session was ended because the same account was opened on another "
            "device or browser.  If this was not you, contact your Zone officer."
        )

    assets = _assets()
    _login_css(assets.get("login_bg") or "")

    # ── block-container is position:fixed over the card area — just render content ──

    # Card header: lock icon + Sign In + red underline + subtitle
    st.markdown("""
    <div style="text-align:center;padding:12px 8px 6px;">
      <div style="width:54px;height:54px;
                  background:linear-gradient(145deg,#001F5E 0%,#003087 100%);
                  border-radius:50%;margin:0 auto 12px;
                  display:flex;align-items:center;justify-content:center;
                  box-shadow:0 0 0 4px rgba(204,0,0,0.20),0 4px 18px rgba(0,31,94,0.42);">
        <svg viewBox="0 0 24 24" width="25" height="25" fill="white">
          <path d="M18 8h-1V6A5 5 0 0 0 7 6v2H6a2 2 0 0 0-2 2v10
                   a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V10a2 2 0 0 0-2-2z
                   M12 17a2 2 0 1 1 0-4 2 2 0 0 1 0 4z
                   M15.1 8H8.9V6a3.1 3.1 0 0 1 6.2 0v2z"/>
        </svg>
      </div>
      <div style="font-size:22px;font-weight:800;color:#001F5E;
                  margin-bottom:5px;letter-spacing:0.3px;">Sign In</div>
      <div style="width:42px;height:3px;background:linear-gradient(90deg,#E53935,#EF5350);
                  border-radius:2px;margin:0 auto 9px;"></div>
      <div style="font-size:11px;color:#003087;opacity:0.75;
                  font-weight:500;letter-spacing:0.4px;margin-bottom:2px;">
        HPCL SOD e-MIS &nbsp;&bull;&nbsp; Authorised Users Only
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── CSS: style Forgot Password submit btn as red italic text ─────────────
    st.markdown("""
    <style>
    div[data-testid="stForm"]
        div[data-testid="column"]:nth-child(2)
        div[data-testid="stFormSubmitButton"] button {
        all: unset !important;
        color: #E53935 !important;
        font-style: italic !important;
        font-size: 13px !important;
        font-weight: 700 !important;
        cursor: pointer !important;
        display: block !important;
        width: 100% !important;
        text-align: right !important;
        padding: 6px 2px !important;
        line-height: 1.4 !important;
    }
    div[data-testid="stForm"]
        div[data-testid="column"]:nth-child(2)
        div[data-testid="stFormSubmitButton"] button:hover {
        text-decoration: underline !important;
        color: #b71c1c !important;
    }
    </style>
    """, unsafe_allow_html=True)

    # ── Login form — original layout with Forgot Password as text link ────────
    with st.form("login_form", clear_on_submit=False):
        loc_code = st.text_input(
            "User ID",
            placeholder="Enter your User ID",
            max_chars=20,
        )
        password = st.text_input(
            "Password",
            type="password",
            placeholder="Enter your password",
        )
        c_rem, c_fgt = st.columns([1, 1])
        with c_rem:
            st.checkbox("Remember me")
        with c_fgt:
            fgt_btn = st.form_submit_button(
                "Forgot Password?", use_container_width=True
            )
        st.markdown("<div style='height:4px;'></div>", unsafe_allow_html=True)
        login_btn = st.form_submit_button("🔑  Login", use_container_width=True, type="primary")

    # ── Forgot Password inline form (no card/tile, plain and minimal) ─────────
    if st.session_state.get("_fp_open"):
        with st.form("forgot_pw_form", clear_on_submit=True):
            fp_code = st.text_input(
                "User ID", placeholder="Your User ID (e.g. 1424)", max_chars=20,
            )
            fp_issue = st.text_area(
                "Describe your issue",
                placeholder="e.g. I forgot my password and need it reset.",
                max_chars=400,
                height=80,
            )
            fp_submit = st.form_submit_button("📨  Submit", use_container_width=True)
        if fp_submit:
            if not fp_code.strip():
                st.error("Please enter your User ID.")
            elif not fp_issue.strip():
                st.error("Please describe your issue.")
            else:
                with st.spinner("Sending…"):
                    import emails as _em
                    _fp_res = _em.send_forgot_password_email(fp_code.strip(), fp_issue.strip())
                if _fp_res.get("ok"):
                    st.success("✅ Your request has been sent to the Admin. You will be contacted shortly.")
                    st.session_state["_fp_open"] = False
                else:
                    st.error(f"Could not send email: {_fp_res.get('msg')}")

    if fgt_btn:
        st.session_state["_fp_open"] = not st.session_state.get("_fp_open", False)
        st.rerun()

    if login_btn:
        if not loc_code or not password:
            st.error("Please enter your User ID and Password.")
        else:
            with st.spinner("Verifying credentials…"):
                result = sheets.check_login(loc_code.strip(), password)
            if result["ok"]:
                import uuid as _uuid, time as _time
                _token = str(_uuid.uuid4())
                sheets.register_session(result["userId"], _token)
                result["_session_token"] = _token
                st.session_state.user = result
                st.session_state["last_activity"]      = _time.time()
                st.session_state["_last_session_check"] = _time.time()
                st.session_state.page = (
                    "change_password" if result["isFirstLogin"] else "dashboard"
                )
                st.rerun()
            else:
                st.error(result["msg"])

    # ── Copyright overlay — bottom-right, inside the image's dark bar ──────────
    st.markdown("""
    <div style="position:fixed;bottom:2.2vh;right:2.2vw;z-index:9999;
                font-size:10.5px;color:rgba(210,225,245,0.88);
                text-align:right;pointer-events:none;
                font-family:'Segoe UI',Arial,sans-serif;letter-spacing:0.2px;">
      &copy; 2026 Hindustan Petroleum Corporation Limited.&nbsp;&nbsp;All rights reserved.
    </div>
    """, unsafe_allow_html=True)

    # ── System Administration — hidden once HQO account exists ───
    _setup_done = sheets.hqo_account_exists()
    if not _setup_done:
        st.markdown("<div style='height:6px;'></div>", unsafe_allow_html=True)
        _show_sa = st.session_state.get("show_sysadmin", False)
        if st.button("⚙️ System Administration", key="btn_sysadmin_toggle",
                     use_container_width=True):
            st.session_state["show_sysadmin"] = not _show_sa
        if st.session_state.get("show_sysadmin"):
            st.caption(
                "One-time tool to seed Zone and HQO login accounts "
                "into UserAccess. Protected by an admin setup key."
            )

            # ── Step 1: Create base accounts ──────────────────────
            st.markdown("**Step 1 — Create base accounts**")
            with st.form("admin_setup_form"):
                setup_key = st.text_input(
                    "Admin Setup Key", type="password",
                    placeholder="Enter the setup key to proceed"
                )
                setup_btn = st.form_submit_button(
                    "Create Zone & HQO Accounts", use_container_width=True
                )
            if setup_btn:
                expected = st.secrets.get("ADMIN_SETUP_KEY", "HPCLADMIN2025")
                if setup_key.strip() != expected:
                    st.error("Invalid setup key.")
                else:
                    with st.spinner("Creating accounts…"):
                        res = sheets.setup_zone_accounts()
                    if res["ok"]:
                        added = res.get("added", [])
                        if added:
                            st.success(f"Created {len(added)} account(s): " + ", ".join(added))
                        else:
                            st.info("All zone and HQO accounts already exist.")
                    else:
                        st.error(f"Setup failed: {res['msg']}")

            st.markdown("---")

            # ── Step 2: Fix conflicting zone IDs ──────────────────
            _ZONE_CORRECTIONS = [
                ("North Zone",         "NORZONE", "NORMIS"),
                ("North West Zone",    "NWZZONE", "NWZMIS"),
                ("North Central Zone", "NCZZONE", "NCZMIS"),
                ("South Zone",         "SOUZONE", "SOUMIS"),
                ("South Central Zone", "SCZZONE", "SCZMIS"),
            ]
            st.markdown("**Step 2 — Fix conflicting zone IDs**")
            st.caption(
                "Applies corrected User IDs and passwords for zones "
                "whose names share the same 3-letter prefix."
            )
            st.dataframe(
                pd.DataFrame(_ZONE_CORRECTIONS,
                             columns=["Zone", "New User ID", "New Password"]),
                use_container_width=True, hide_index=True,
            )
            if st.button("Apply Zone ID Corrections", use_container_width=True,
                         key="btn_fix_zones"):
                errs, ok_list = [], []
                for zone_name, uid, pw in _ZONE_CORRECTIONS:
                    with st.spinner(f"Updating {zone_name}…"):
                        r = sheets.upsert_zone_account(zone_name, uid, pw)
                    if r["ok"]:
                        ok_list.append(f"{uid} ({r['action']})")
                    else:
                        errs.append(f"{uid}: {r.get('msg','')}")
                if ok_list:
                    st.success("Updated: " + ", ".join(ok_list))
                if errs:
                    st.error("Errors: " + "; ".join(errs))

            st.markdown("---")

            # ── Diagnostic: view current accounts ─────────────────
            if st.button("🔍 View Current Zone & HQO Accounts",
                         use_container_width=True, key="btn_view_accts"):
                with st.spinner("Reading UserAccess…"):
                    accts = sheets.get_zone_admin_accounts()
                if accts:
                    df = pd.DataFrame(accts)[["user_id", "password", "role", "zone"]]
                    df.columns = ["User ID", "Password", "Role", "Zone"]
                    st.dataframe(df, use_container_width=True, hide_index=True)
                else:
                    st.warning("No Zone or Admin accounts found in UserAccess.")



# ── Change Password page ──────────────────────────────────────────────────────

def show_change_password():
    _login_css()
    user = st.session_state.user

    _, mid, _ = st.columns([1, 10, 1])
    with mid:
        st.markdown("<div style='height:48px;'></div>", unsafe_allow_html=True)

        dh = _assets().get("dh_logo")
        if dh:
            st.markdown(
                f'<div style="text-align:center;margin-bottom:18px;">'
                f'<img src="{dh}" style="height:65px;object-fit:contain;"></div>',
                unsafe_allow_html=True,
            )

        st.markdown(f"""
        <h1 style="font-size:20px;font-weight:800;color:{HPCL_BLUE};margin:0 0 6px;">
          🔐 Set Your New Password
        </h1>
        <p style="font-size:13px;color:#888;margin:0 0 6px;">
          Location: <strong style="color:#1a1a2e;">{user['locName']}</strong>
          ({user['userId']}) &nbsp;·&nbsp; Zone: {user['zone']}
        </p>
        <div style="width:44px;height:3px;margin-bottom:16px;
                    background:linear-gradient(to right,#e53935,#0033A0);border-radius:2px;"></div>
        <hr style="border:none;border-top:1px solid #f0f0f0;margin-bottom:16px;">
        """, unsafe_allow_html=True)

        if user.get("isFirstLogin"):
            _is_default = user.get("_password", "") == user.get("userId", "")
            if _is_default:
                st.warning(
                    "🔑  Your current password is the **default password (same as your User ID)**.  "
                    "You must set a new password before you can access the portal."
                )
            else:
                st.info("**First Login detected.** You must set a new password to continue.")

        with st.form("chgpass_form"):
            curr = st.text_input("Current Password",     type="password")
            new1 = st.text_input("New Password",         type="password",
                                 help="Minimum 6 characters")
            new2 = st.text_input("Confirm New Password", type="password")
            chg  = st.form_submit_button("Change Password", use_container_width=True)

        if chg:
            with st.spinner("Updating…"):
                res = sheets.change_password(user["userId"], curr, new1, new2)
            if res["ok"]:
                st.session_state.user["isFirstLogin"] = False
                st.session_state.user["_password"]    = new1
                st.session_state.page  = "dashboard"
                st.session_state.flash = "Password changed successfully! Welcome to the SOD MIS Portal."
                st.rerun()
            else:
                st.error(res["msg"])

        if st.button("← Back to Login", key="back_btn"):
            _lo_uid = (st.session_state.get("user") or {}).get("userId", "")
            if _lo_uid:
                sheets.clear_session(_lo_uid)
            st.session_state.user = None
            st.session_state.page = "login"
            st.rerun()

    st.markdown("""
    <div class="login-footer">
      <span>© 2026 Hindustan Petroleum Corporation Limited. All rights reserved.</span>
      <span>HPCL SOD &nbsp;·&nbsp; Supply Operations & Distribution &nbsp;·&nbsp; Authorised Users Only</span>
    </div>""", unsafe_allow_html=True)


# ── Detail-table helpers (Railway Claims / IRR Details / Legal Cases) ────────

_DETAIL_UI = {
    "RAILWAY_CLAIMS": {
        "title": "Railway Claims Details",
        "sub":   "Railway Claims",
        "cols": {
            "claim_no":          {"label": "Claim No.",         "type": "text"},
            "year":              {"label": "Year",              "type": "int",  "min": 1990, "max": 2100},
            "amount":            {"label": "Amount (Rs)",       "type": "float"},
            "rr_nos":            {"label": "RR Nos.",           "type": "text"},
            "ex_station":        {"label": "Ex",                "type": "text"},
            "to_station":        {"label": "To",                "type": "text"},
            "wagon_nos":         {"label": "T/Wagon Nos.",      "type": "text"},
            "product":           {"label": "Product",           "type": "text"},
            "qty":               {"label": "Qty.",              "type": "float"},
            "rly":               {"label": "Rly.",              "type": "text"},
            "pending_stage":     {"label": "Pending Stage",     "type": "text"},
            "status_claim":      {"label": "Status of Claim",   "type": "text"},
            "last_hearing":      {"label": "Last Hearing Date", "type": "text"},
            "next_hearing":      {"label": "Next Hearing Date", "type": "text"},
            "rct_status":        {"label": "RCT Status",        "type": "text"},
            "case_facts":        {"label": "Case Facts",        "type": "text"},
            "rejection_reasons": {"label": "Rejection Reasons", "type": "text"},
            "shortcomings":      {"label": "ShortComings",      "type": "text"},
            "strength":          {"label": "Strength of Case",  "type": "text"},
            "recommendation":    {"label": "Recommendation",    "type": "text"},
        },
    },
    "IRR_DETAILS": {
        "title": "IRR Details",
        "sub":   "IRR Data",
        "cols": {
            "irr_no":       {"label": "IRR #",          "type": "text"},
            "irr_date":     {"label": "IRR Date",       "type": "text"},
            "description":  {"label": "Description",    "type": "text"},
            "amount":       {"label": "Amount (Rs)",    "type": "float"},
            "status":       {"label": "Status",         "type": "select", "opts": ["OPEN", "CLOSED"]},
            "closure_date": {"label": "Closure Date",   "type": "text"},
        },
    },
    "LEGAL_CASES": {
        "title": "Legal Cases Details",
        "sub":   "Legal Data",
        "cols": {
            "court_name":  {"label": "Court Name",              "type": "text"},
            "case_number": {"label": "Case Number",             "type": "text"},
            "cause_title": {"label": "Cause Title",             "type": "text"},
            "advocate":    {"label": "Advocate",                "type": "text"},
            "nature":      {"label": "Nature of Case",          "type": "text"},
            "dealership":  {"label": "Dealership / Location",   "type": "text"},
            "background":  {"label": "Background",              "type": "text"},
            "status":      {"label": "Status",                  "type": "text"},
            "last_hearing":{"label": "Last Hearing Date",       "type": "text"},
            "next_hearing":{"label": "Next Hearing Date",       "type": "text"},
        },
    },
}


def _load_detail_to_session(user_id: str, month_year: str, tab_key: str, mc: str):
    """Load detail table rows from Google Sheets to session state (once per section/month)."""
    flag = f"det_loaded_{tab_key}_{mc}"
    if st.session_state.get(flag):
        return
    rows = sheets.load_detail_table(user_id, month_year, tab_key)
    st.session_state.pop(f"de_{tab_key}_{mc}", None)     # clear editor so it re-inits
    st.session_state[f"det_init_{tab_key}_{mc}"] = rows  # store initial data
    st.session_state[flag] = True


def _render_detail_table(tab_key: str, month_year: str, is_locked: bool) -> pd.DataFrame:
    """Render a data_editor for a detail table and return the edited DataFrame."""
    mc   = month_year.replace("-", "_")
    ui   = _DETAIL_UI[tab_key]
    keys = list(ui["cols"].keys())

    init_rows = st.session_state.get(f"det_init_{tab_key}_{mc}", [])
    if init_rows:
        df_init = pd.DataFrame(init_rows, columns=keys)
        for k, cfg in ui["cols"].items():
            if cfg["type"] == "float":
                df_init[k] = pd.to_numeric(df_init[k], errors="coerce")
            elif cfg["type"] == "int":
                df_init[k] = pd.to_numeric(df_init[k], errors="coerce")
    else:
        df_init = pd.DataFrame(columns=keys)

    # Build Streamlit column_config
    col_cfg = {}
    for k, cfg in ui["cols"].items():
        t = cfg["type"]
        lbl = cfg["label"]
        if t == "text":
            col_cfg[k] = st.column_config.TextColumn(lbl)
        elif t == "float":
            col_cfg[k] = st.column_config.NumberColumn(lbl, format="%.2f")
        elif t == "int":
            col_cfg[k] = st.column_config.NumberColumn(
                lbl, min_value=cfg.get("min"), max_value=cfg.get("max"),
                step=1, format="%d")
        elif t == "select":
            col_cfg[k] = st.column_config.SelectboxColumn(lbl, options=cfg["opts"])

    if is_locked:
        # Read-only view: render as styled HTML table with HPCL blue header
        st.markdown(
            f'<div class="sub-header" style="margin-top:22px;">&#128203; &nbsp; {ui["title"]}</div>',
            unsafe_allow_html=True,
        )
        if df_init.empty:
            st.info("No records entered.")
        else:
            display_df = df_init.rename(
                columns={k: v["label"] for k, v in ui["cols"].items()}
            )
            # Build styled HTML table with coloured headers
            cols_html = "".join(
                f'<th style="background:#0033A0;color:#ffffff;font-weight:700;'
                f'font-size:11px;padding:7px 12px;text-align:left;'
                f'border-right:1px solid #1a5fcc;white-space:nowrap;">{c}</th>'
                for c in display_df.columns
            )
            rows_html = ""
            for ri, row in display_df.iterrows():
                bg = "#f8faff" if ri % 2 == 0 else "#ffffff"
                cells = "".join(
                    f'<td style="padding:5px 12px;font-size:11px;color:#1a1a2e;'
                    f'border-right:1px solid #eef0f5;">{v if v not in (None,"") else "—"}</td>'
                    for v in row
                )
                rows_html += f'<tr style="background:{bg};">{cells}</tr>'
            html = (
                f'<div style="overflow-x:auto;border-radius:8px;'
                f'border:1px solid #d0d7e8;margin-top:8px;">'
                f'<table style="border-collapse:collapse;width:100%;min-width:600px;">'
                f'<thead><tr>{cols_html}</tr></thead>'
                f'<tbody>{rows_html}</tbody>'
                f'</table></div>'
            )
            st.markdown(html, unsafe_allow_html=True)
        return df_init

    st.markdown(
        f'<div class="sub-header" style="margin-top:22px;">'
        f'&#128203; &nbsp; {ui["title"]}'
        f'<span style="font-size:11px;font-weight:400;margin-left:10px;opacity:0.8;">'
        f'— Use + button to add rows &nbsp;·&nbsp; Click a cell to edit</span>'
        f'</div>',
        unsafe_allow_html=True,
    )

    edited = st.data_editor(
        df_init,
        key=f"de_{tab_key}_{mc}",
        num_rows="dynamic",
        use_container_width=True,
        disabled=False,
        column_config=col_cfg,
        hide_index=True,
    )
    return edited


def _save_detail_tables(detail_dfs: dict, user: dict, month_year: str, is_locked: bool):
    """Save each detail table DataFrame to its Google Sheet tab."""
    if is_locked or not detail_dfs:
        return
    for tab_key, df in detail_dfs.items():
        if df is None or df.empty:
            continue
        rows = df.dropna(how="all").to_dict("records")
        rows = [r for r in rows if any(str(v).strip() for v in r.values())]
        sheets.save_detail_table(user["userId"], month_year, tab_key, rows, user)


# ── Section data-entry form (Phase 3) ────────────────────────────────────────

def _inject_field_enhancements(fields: list):
    """
    Inject JS (via components iframe→parent) that shows a floating
    validation-rule tooltip on hover AND focus for number, textarea,
    and selectbox widgets.
    """
    import json
    import streamlit.components.v1 as components

    tt = {}
    for f in fields:
        if f.get("auto"):
            continue
        lbl = f["label"] + (" *" if f.get("req") else "")
        tt[lbl] = _field_help(f)

    tj = json.dumps(tt, ensure_ascii=False)

    components.html(f"""
<script>
(function(){{
  var T={tj};
  var pw=window.parent, pd=pw.document;

  /* ── Create tooltip element once in parent doc ── */
  if(!pd.getElementById('_mis_tip')){{
    var cs=pd.createElement('style');
    cs.textContent=
      '#_mis_tip{{position:fixed;background:#001a6e;color:#fff;'+
      'padding:10px 16px;border-radius:10px;font-size:12px;'+
      'font-family:Segoe UI,Tahoma,sans-serif;max-width:400px;'+
      'z-index:99999;pointer-events:none;display:none;'+
      'box-shadow:0 6px 24px rgba(0,26,110,.45);'+
      'line-height:1.75;border:1px solid rgba(255,255,255,.14)}}'+
      '#_mis_tip::after{{content:"";position:absolute;bottom:-8px;left:22px;'+
      'border:8px solid transparent;border-top-color:#001a6e;border-bottom:none}}';
    pd.head.appendChild(cs);
    var tb=pd.createElement('div'); tb.id='_mis_tip';
    pd.body.appendChild(tb);
  }}
  var TIP=pd.getElementById('_mis_tip');

  var _ALL_TID=
    '[data-testid="stNumberInput"],[data-testid="stTextInput"],'+
    '[data-testid="stTextArea"],[data-testid="stSelectbox"],'+
    '[data-testid="stDateInput"]';

  /* ── Get label text by traversing UP from the inner element ── */
  function getLabel(el){{
    var c=el.closest(_ALL_TID);
    return c?(c.querySelector('label')||{{}}).innerText||'':'';
  }}

  /* ── Show tooltip above whichever container holds el ── */
  /* position:fixed → use viewport rect directly, NO scrollY offset  */
  function showTip(el){{
    var lb=getLabel(el).trim(), txt=T[lb];
    if(!txt)return;
    var pts=txt.split('  ·  ');
    TIP.innerHTML=
      '<strong style="color:#90caf9;font-size:12.5px">'+pts[0]+'</strong>'+
      pts.slice(1).map(function(x){{
        return '<br><span style="opacity:.85">&#183; '+x+'</span>';
      }}).join('');
    TIP.style.display='block';
    var posEl=el.closest(_ALL_TID)||el;
    var r=posEl.getBoundingClientRect(), h=TIP.offsetHeight||56;
    TIP.style.left=Math.max(8,r.left)+'px';
    TIP.style.top=(r.top-h-12)+'px';
  }}
  function hideTip(){{ TIP.style.display='none'; }}

  /* ── Attach handlers; mark INNER elements to survive Streamlit re-renders ── */
  function setup(){{

    /* Number / int inputs */
    pd.querySelectorAll('[data-testid="stNumberInput"] input').forEach(function(el){{
      if(el.dataset.mis)return; el.dataset.mis='1';
      el.addEventListener('focus',function(){{ showTip(el); }});
      el.addEventListener('blur', function(){{ hideTip(); }});
      var ni=el.closest('[data-testid="stNumberInput"]');
      if(ni&&!ni.dataset.misHov){{ ni.dataset.misHov='1';
        ni.addEventListener('mouseenter',function(){{ showTip(el); }});
        ni.addEventListener('mouseleave',function(){{ hideTip(); }});
      }}
    }});

    /* Text inputs */
    pd.querySelectorAll('[data-testid="stTextInput"] input').forEach(function(el){{
      if(el.dataset.mis)return; el.dataset.mis='1';
      el.addEventListener('focus',function(){{ showTip(el); }});
      el.addEventListener('blur', function(){{ hideTip(); }});
      var wi=el.closest('[data-testid="stTextInput"]');
      if(wi&&!wi.dataset.misHov){{ wi.dataset.misHov='1';
        wi.addEventListener('mouseenter',function(){{ showTip(el); }});
        wi.addEventListener('mouseleave',function(){{ hideTip(); }});
      }}
    }});

    /* Date inputs */
    pd.querySelectorAll('[data-testid="stDateInput"] input').forEach(function(el){{
      if(el.dataset.mis)return; el.dataset.mis='1';
      el.addEventListener('focus',function(){{ showTip(el); }});
      el.addEventListener('blur', function(){{ hideTip(); }});
      var wd=el.closest('[data-testid="stDateInput"]');
      if(wd&&!wd.dataset.misHov){{ wd.dataset.misHov='1';
        wd.addEventListener('mouseenter',function(){{ showTip(el); }});
        wd.addEventListener('mouseleave',function(){{ hideTip(); }});
      }}
    }});

    /* Text areas */
    pd.querySelectorAll('[data-testid="stTextArea"] textarea').forEach(function(ta){{
      if(ta.dataset.mis)return; ta.dataset.mis='1';
      ta.addEventListener('focus',function(){{ showTip(ta); }});
      ta.addEventListener('blur', function(){{ hideTip(); }});
      var wa=ta.closest('[data-testid="stTextArea"]');
      if(wa&&!wa.dataset.misHov){{ wa.dataset.misHov='1';
        wa.addEventListener('mouseenter',function(){{ showTip(ta); }});
        wa.addEventListener('mouseleave',function(){{ hideTip(); }});
      }}
    }});

    /* Selectboxes */
    pd.querySelectorAll('[data-testid="stSelectbox"]').forEach(function(sb){{
      if(sb.dataset.mis)return; sb.dataset.mis='1';
      function tipEl(){{
        return sb.querySelector('[role="combobox"]')||sb.querySelector('input')||sb;
      }}
      sb.addEventListener('mouseenter',function(){{ showTip(tipEl()); }});
      sb.addEventListener('mouseleave',function(){{ hideTip(); }});
      var inp=tipEl();
      if(inp&&inp!==sb){{
        inp.addEventListener('focus',function(){{ showTip(inp); }});
        inp.addEventListener('blur', function(){{ hideTip(); }});
      }}
    }});

  }}

  /* Run immediately, then watch for Streamlit re-renders (debounced) */
  setup();
  if(pw._misObs)pw._misObs.disconnect();
  pw._misObs=new MutationObserver(function(){{
    if(!pw._misPending){{
      pw._misPending=true;
      requestAnimationFrame(function(){{ pw._misPending=false; setup(); }});
    }}
  }});
  pw._misObs.observe(pd.body,{{childList:true,subtree:true}});
}})();
</script>
""", height=0)


def show_section_form(section_num: int, user: dict, month_year: str, month_label: str):
    _dashboard_css()

    st.markdown("""
    <style>
    /* ── Sub-section headers ── */
    .sub-header {
        background: linear-gradient(90deg,#001a6e 0%,#0033A0 60%,#0050d0 100%);
        color: white; font-size: 13px; font-weight: 700;
        padding: 9px 18px; border-radius: 10px; margin: 20px 0 12px;
        letter-spacing: 0.4px; box-shadow: 0 2px 8px rgba(0,51,160,0.18);
    }

    /* ── Auto-calc display boxes ── */
    .auto-box {
        background: linear-gradient(135deg,#e8f0ff 0%,#dce8ff 100%);
        border: 1.5px solid #b0c8f8; border-radius: 10px;
        padding: 12px 16px; margin-bottom: 10px;
        box-shadow: 0 1px 4px rgba(0,51,160,0.08);
    }
    .auto-label { font-size: 10.5px; font-weight: 700; color: #0033A0;
        text-transform: uppercase; letter-spacing: 0.6px; margin-bottom: 6px; }
    .auto-val   { font-size: 22px; font-weight: 800; color: #001a6e; line-height: 1; }
    .auto-hint  { font-size: 10.5px; color: #5577bb; margin-top: 4px; }

    /* ── HPCL blue bold labels on every field type ── */
    [data-testid="stNumberInput"] label,
    [data-testid="stTextInput"]   label,
    [data-testid="stSelectbox"]   label,
    [data-testid="stTextArea"]    label {
        color: #001a6e !important;
        font-weight: 800 !important;
        font-size: 13px !important;
        letter-spacing: 0.1px !important;
        text-transform: none !important;
    }

    /* ── Number inputs & textareas: red by default, green when filled, blue on focus ── */
    [data-testid="stNumberInput"] input,
    [data-testid="stTextArea"] textarea {
        border: 1.5px solid #c62828 !important;
        border-radius: 8px !important;
        background: #fffefe !important;
        padding: 10px 12px !important;
        font-size: 14px !important;
        transition: border-color 0.2s ease, box-shadow 0.2s ease !important;
        outline: none !important;
    }
    /* Green when value present (placeholder hidden = value entered) */
    [data-testid="stNumberInput"] input:not(:placeholder-shown),
    [data-testid="stTextArea"] textarea:not(:placeholder-shown) {
        border-color: #2e7d32 !important;
        box-shadow: 0 0 0 2px rgba(46,125,50,0.1) !important;
    }
    /* HPCL blue on focus */
    [data-testid="stNumberInput"] input:focus,
    [data-testid="stTextArea"] textarea:focus {
        border-color: #0033A0 !important;
        box-shadow: 0 0 0 3px rgba(0,51,160,0.15) !important;
    }

    /* ── Hide +/− spinner buttons ── */
    [data-testid="stNumberInput"] button {
        display: none !important;
    }

    /* ── Selectbox: green border (form selects always default to index 0, always have a value) ── */
    [data-testid="stSelectbox"] [data-baseweb="select"] > div:first-child {
        border: 1.5px solid #2e7d32 !important;
        box-shadow: 0 0 0 2px rgba(46,125,50,0.1) !important;
        border-radius: 8px !important;
        transition: border-color 0.25s ease, box-shadow 0.25s ease !important;
    }

    /* ── Selectbox value text: dark bold — target the SingleValue/Placeholder div via DOM path ── */
    [data-testid="stSelectbox"] [data-baseweb="select"] > div:first-child > div:first-child > div:first-child {
        color: #0d0d1a !important;
        font-weight: 700 !important;
        font-size: 14px !important;
    }

    /* ── Hide Streamlit's "Press Enter to apply" hint — Tab/blur already commits value ── */
    [data-testid="InputInstructions"] {
        display: none !important;
    }

    /* ── Disabled inputs (locked/checker view): force dark bold text ── */
    [data-testid="stNumberInput"] input:disabled,
    [data-testid="stTextInput"]   input:disabled,
    [data-testid="stDateInput"]   input:disabled {
        color: #0d0d1a !important;
        -webkit-text-fill-color: #0d0d1a !important;
        font-weight: 700 !important;
        opacity: 1 !important;
    }
    [data-testid="stTextArea"] textarea:disabled {
        color: #0d0d1a !important;
        -webkit-text-fill-color: #0d0d1a !important;
        font-weight: 600 !important;
        opacity: 1 !important;
    }
    /* ── Tooltip icon slightly larger ── */
    [data-testid="stTooltipIcon"] svg {
        width: 15px !important; height: 15px !important;
        color: #4a80e8 !important; opacity: 0.9 !important;
    }
    </style>
    """, unsafe_allow_html=True)

    # ── Resolve lock status BEFORE loading draft so we have a current view ──────
    # This must run first: if status changed from PENDING_REVIEW→REJECTED we clear
    # the draft_loaded flags and rerun so _load_draft_to_session reloads from Sheets.
    _lock_key = f"lock_{month_year}"
    if _lock_key not in st.session_state:
        st.session_state[_lock_key] = sheets.get_month_status(user["userId"], month_year)
    elif st.session_state[_lock_key].get("status") == "PENDING_REVIEW":
        _fresh = sheets.get_month_status(user["userId"], month_year)
        if _fresh["status"] != "PENDING_REVIEW":
            # Checker acted — clear all draft_loaded flags so every section
            # re-reads the saved draft from Google Sheets on the next render.
            _mc0 = month_year.replace("-", "_")
            for _s0 in range(1, 11):
                st.session_state.pop(f"draft_loaded_{_mc0}_s{_s0}", None)
            st.session_state[_lock_key] = _fresh
            st.rerun()                  # restart render with fresh status + no stale flags
        st.session_state[_lock_key] = _fresh

    _load_draft_to_session(user["userId"], month_year, section_num)

    # Load detail table data (once per session) for sections that have them
    mc = month_year.replace("-", "_")
    if section_num == 3:
        _load_detail_to_session(user["userId"], month_year, "RAILWAY_CLAIMS", mc)
    elif section_num == 10:
        _load_detail_to_session(user["userId"], month_year, "IRR_DETAILS",    mc)
        _load_detail_to_session(user["userId"], month_year, "LEGAL_CASES",    mc)

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        spb = _assets().get("side_panel_banner")
        if spb:
            st.markdown(
                f'<div style="margin:0;padding:0;line-height:0;width:100%;overflow:hidden;">'
                f'<img src="{spb}" style="width:100%;height:auto;display:block;'
                f'border-radius:0;margin:0;padding:0;"></div>',
                unsafe_allow_html=True,
            )

        if st.button("🏠  Back to Dashboard", key="sid_dash", use_container_width=True):
            st.session_state.selected_section = None
            st.rerun()

        st.markdown("""
        <div style="padding:10px 16px 4px;">
          <div style="color:#C8D7FF;font-size:9px;font-weight:700;
                      letter-spacing:2px;text-transform:uppercase;">MIS SECTIONS</div>
        </div>
        """, unsafe_allow_html=True)

        for num, name in SECTIONS:
            lbl = f"▶  S{num} - {name}"
            if num == section_num:
                st.button(lbl, key=f"sid_{num}", disabled=True, use_container_width=True)
            else:
                if st.button(lbl, key=f"sid_{num}", use_container_width=True):
                    st.session_state.selected_section = num
                    st.rerun()
            if num == 5 and user.get("locType", "HPCL") == "HPCL":
                if st.button("↳ S5A - M&I MIS", key="btn_mi_mis_sec", use_container_width=True,
                             help="Maintenance & Inspection detailed MIS entry"):
                    st.session_state.selected_section = "mi_mis"
                    st.rerun()

    # ── Page header ───────────────────────────────────────────────────────────
    _dash_header(user)

    # ── Section title bar ────────────────────────────────────────────────────
    # Lock status was already resolved at the top of this function into _lock_key.
    lock_cache_key = _lock_key
    status_data    = st.session_state[lock_cache_key]
    is_locked      = status_data["is_locked"]
    sec_name    = SECTION_NAMES.get(section_num, f"S{section_num}")
    lock_note   = "&nbsp;·&nbsp; 🔒 <em>Read-only (locked)</em>" if is_locked else ""
    lock_badge  = (
        '<span style="background:#e8f5e9;color:#2e7d32;padding:4px 14px;'
        'border-radius:8px;font-size:12px;font-weight:600;">&#10003; Locked</span>'
        if is_locked else ""
    )
    st.markdown(f"""
    <div style="background:white;border-radius:14px;padding:14px 22px;
                box-shadow:0 2px 10px rgba(0,43,143,0.09);margin-bottom:14px;
                display:flex;align-items:center;justify-content:space-between;">
      <div>
        <div style="font-size:16px;font-weight:700;color:#002b8f;">
          &#128203; {sec_name}
        </div>
        <div style="font-size:12px;color:#667085;margin-top:3px;">
          Period: <strong>{month_label}</strong> &nbsp;·&nbsp;
          Location: <strong>{user['locName']}</strong>{lock_note}
        </div>
      </div>
      {lock_badge}
    </div>
    """, unsafe_allow_html=True)

    # ── Location-type field filtering ─────────────────────────────────────────
    from form_defs import get_excluded_fields, get_skip_sections
    loc_type     = user.get("locType", "HPCL")
    excl_fields  = get_excluded_fields(loc_type)
    skip_secs    = get_skip_sections(loc_type)

    if section_num in skip_secs:
        st.info(
            f"This section is **Not Applicable** for {loc_type} locations "
            f"and is automatically marked complete.",
            icon="ℹ️",
        )
        return

    # ── Compute auto-calc field values ────────────────────────────────────────
    fields = [f for f in SECTION_FIELDS.get(section_num, []) if f["key"] not in excl_fields]
    # Include all sections' field values so cross-section auto-calc refs (e.g. f61 = f60/f3) work
    _all_sec_fields = [f for sec in SECTION_FIELDS.values() for f in sec]
    all_vals = {f["key"]: st.session_state.get(_sk(month_year, f["key"])) for f in _all_sec_fields}
    for f in fields:
        if f.get("auto"):
            val = _compute_auto(f["auto"], all_vals)
            sk  = _sk(month_year, f["key"])
            st.session_state[sk] = val
            all_vals[f["key"]]   = val

    # ── Render fields grouped by sub-section ──────────────────────────────────
    seen_subs: list = []
    sub_groups: dict = {}
    for f in fields:
        sub = f["sub"]
        if sub not in sub_groups:
            sub_groups[sub] = []
            seen_subs.append(sub)
        sub_groups[sub].append(f)

    def _visible(fld: dict) -> bool:
        """Return False if the field has a show_when condition that isn't met."""
        sw = fld.get("show_when")
        if not sw:
            return True
        return all(str(all_vals.get(k) or "") == str(v) for k, v in sw.items())

    for sub in seen_subs:
        sfl = sub_groups[sub]
        # Filter to only visible fields (respecting show_when conditions)
        vis_sfl = [f for f in sfl if _visible(f)]
        if not vis_sfl:
            continue
        st.markdown(f'<div class="sub-header">&#128204; &nbsp; {sub}</div>',
                    unsafe_allow_html=True)
        i = 0
        while i < len(vis_sfl):
            f         = vis_sfl[i]
            full_w    = f["type"] == "textarea" or bool(f.get("auto"))
            next_full = (i + 1 < len(vis_sfl) and
                         (vis_sfl[i + 1]["type"] == "textarea" or bool(vis_sfl[i + 1].get("auto"))))
            if full_w or next_full:
                _render_field(f, _sk(month_year, f["key"]), is_locked, all_vals)
                i += 1
            else:
                col1, col2 = st.columns(2)
                with col1:
                    _render_field(f, _sk(month_year, f["key"]), is_locked, all_vals)
                if i + 1 < len(vis_sfl):
                    with col2:
                        f2 = vis_sfl[i + 1]
                        _render_field(f2, _sk(month_year, f2["key"]), is_locked, all_vals)
                    i += 2
                else:
                    i += 1

    # Inject focus-tooltip + red/green border JS (runs in hidden iframe → parent)
    _inject_field_enhancements(fields)

    # ── Cross-field validation warnings ───────────────────────────────────────
    if section_num == 1 and not is_locked:
        try:
            _total    = all_vals.get("f160")
            _complied = all_vals.get("f161")
            if _total is not None and _complied is not None:
                if int(_complied) > int(_total):
                    st.error(
                        "⛔  **Complied Recommendations cannot exceed Total Recommendations.**  "
                        "Correct the value of 'No. of Complied Recommendations' before saving — "
                        f"currently **{_complied}** complied out of **{_total}** total."
                    )
        except (ValueError, TypeError):
            pass

    # ── Detail tables for S3 and S10 ─────────────────────────────────────────
    detail_dfs: dict = {}
    if section_num == 3:
        detail_dfs["RAILWAY_CLAIMS"] = _render_detail_table(
            "RAILWAY_CLAIMS", month_year, is_locked)
    elif section_num == 10:
        detail_dfs["IRR_DETAILS"] = _render_detail_table(
            "IRR_DETAILS", month_year, is_locked)
        detail_dfs["LEGAL_CASES"] = _render_detail_table(
            "LEGAL_CASES", month_year, is_locked)

    # ── Bottom navigation & save ──────────────────────────────────────────────
    st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)
    navigate_to = None
    col_prev, col_save, col_next = st.columns([1.5, 3, 1.5])

    with col_prev:
        if section_num > 1:
            prev_label = SECTION_NAMES.get(section_num - 1, "").split("—")[-1].strip()[:14]
            if st.button(f"⬅  S{section_num - 1}: {prev_label}",
                         key="btn_prev", use_container_width=True):
                navigate_to = section_num - 1

    with col_save:
        save_clicked = False
        if not is_locked:
            save_clicked = st.button("💾  Save Draft", key="btn_save",
                                     use_container_width=True, type="primary")
        else:
            st.markdown(
                '<p style="text-align:center;color:#8c9db5;font-style:italic;padding-top:8px;">'
                '🔒 Read-only — submission is locked</p>',
                unsafe_allow_html=True,
            )

    with col_next:
        if section_num == 5 and user.get("locType", "HPCL") == "HPCL":
            # After S5 M&I (HPCL only), go to S5A (M&I MIS detail)
            if st.button("S5A: M&I MIS  ➡", key="btn_next", use_container_width=True):
                navigate_to = "mi_mis"
        elif section_num == 10:
            # After S10 (last section), offer Back to Dashboard
            if st.button("🏠 Back to Dashboard", key="btn_next", use_container_width=True):
                navigate_to = "dashboard"
        elif section_num < 10:
            next_label = SECTION_NAMES.get(section_num + 1, "").split("—")[-1].strip()[:14]
            if st.button(f"S{section_num + 1}: {next_label}  ➡",
                         key="btn_next", use_container_width=True):
                navigate_to = section_num + 1

    # ── Process navigation & save OUTSIDE columns so st.rerun() is safe ──────
    if navigate_to is not None:
        with st.spinner("Saving…"):
            _do_save(user["userId"], month_year, section_num, fields, is_locked)
            _save_detail_tables(detail_dfs, user, month_year, is_locked)
        if navigate_to == "dashboard":
            st.session_state.selected_section = None
        elif navigate_to == "mi_mis":
            st.session_state.selected_section = "mi_mis"
        else:
            st.session_state.pop(f"draft_loaded_{mc}_s{navigate_to}", None)
            st.session_state.selected_section = navigate_to
        st.rerun()
    elif save_clicked:
        with st.spinner("Saving…"):
            res = _do_save(user["userId"], month_year, section_num, fields, is_locked)
            _save_detail_tables(detail_dfs, user, month_year, is_locked)
        if res.get("ok"):
            st.session_state.pop(lock_cache_key, None)
            pct    = int(res.get("pct", 0))
            marker = " ✅ Section complete!" if section_num in res.get("secs_done", []) else ""
            st.success(f"Draft saved! Overall completion: {pct}%{marker}")
        else:
            st.error(f"Save failed: {res.get('msg', 'Unknown error')}")

    st.markdown("""
    <div style="margin-top:24px;padding:10px 4px;border-top:1px solid #dde3ed;
                display:flex;justify-content:space-between;font-size:11px;color:#aaa;">
      <span>&#169; 2026 Hindustan Petroleum Corporation Limited. All rights reserved.</span>
      <span>HPCL SOD &nbsp;·&nbsp; Supply Operations &amp; Distribution</span>
    </div>
    """, unsafe_allow_html=True)


# ── Phase-5: read-only review helpers ────────────────────────────────────────

def _render_section_readonly(sec_num: int, draft: dict):
    """Render one section's fields as an HTML read-only table (no interactive widgets)."""
    fields = SECTION_FIELDS.get(sec_num, [])
    if not fields:
        st.caption("No fields defined for this section.")
        return

    # Build all_vals so auto-calc expressions resolve correctly
    _all_sec_fields = [f for sec in SECTION_FIELDS.values() for f in sec]
    all_vals: dict = {}
    for f in _all_sec_fields:
        raw = draft.get(f["key"])
        if raw is not None:
            try:
                all_vals[f["key"]] = (
                    float(raw) if f["type"] in ("int", "number") else raw
                )
            except (TypeError, ValueError):
                all_vals[f["key"]] = raw
        else:
            all_vals[f["key"]] = None
    for f in fields:
        if f.get("auto"):
            all_vals[f["key"]] = _compute_auto(f["auto"], all_vals)

    # Group by sub-section
    seen_subs: list  = []
    sub_groups: dict = {}
    for f in fields:
        s = f["sub"]
        if s not in sub_groups:
            sub_groups[s] = []
            seen_subs.append(s)
        sub_groups[s].append(f)

    def _ro_visible(fld: dict) -> bool:
        sw = fld.get("show_when")
        if not sw:
            return True
        return all(str(all_vals.get(k) or "") == str(v) for k, v in sw.items())

    for sub in seen_subs:
        sfl = [f for f in sub_groups[sub] if _ro_visible(f)]
        if not sfl:
            continue
        st.markdown(
            f'<div class="sub-header" style="font-size:12px;padding:7px 16px;">'
            f'&#128204; &nbsp; {sub}</div>',
            unsafe_allow_html=True,
        )

        rows_html = ""
        i = 0
        while i < len(sfl):
            f      = sfl[i]
            is_auto = bool(f.get("auto"))
            is_wide = f["type"] == "textarea" or is_auto

            if is_auto:
                val  = all_vals.get(f["key"])
                dec  = f.get("dec") or 0
                disp = f"{val:.{dec}f}" if val is not None else "—"
                rows_html += (
                    f'<tr style="background:#eef4ff;">'
                    f'<td colspan="4" style="padding:8px 14px;">'
                    f'<span style="font-size:11px;font-weight:700;color:#0033A0;'
                    f'text-transform:uppercase;letter-spacing:0.5px;">{f["label"]}</span>'
                    f'<span style="font-size:10px;color:#5577bb;margin-left:6px;">[Auto]</span><br>'
                    f'<span style="font-size:20px;font-weight:800;color:#001a6e;">{disp}</span>'
                    f'</td></tr>'
                )
                i += 1
            elif f["type"] == "textarea":
                raw  = draft.get(f["key"], "")
                disp = str(raw).replace("\n", "<br>") if raw not in (None, "") else "<em style='color:#aaa;'>—</em>"
                rows_html += (
                    f'<tr><td colspan="4" style="padding:8px 14px;">'
                    f'<span style="font-size:11px;font-weight:600;color:#334;">{f["label"]}</span><br>'
                    f'<span style="font-size:13px;font-weight:600;color:#0d0d1a;">{disp}</span>'
                    f'</td></tr>'
                )
                i += 1
            else:
                def _cell(fld: dict) -> str:
                    raw  = draft.get(fld["key"], "")
                    disp = str(raw) if raw not in (None, "") else "<em style='color:#aaa;'>—</em>"
                    return (
                        f'<td style="padding:8px 14px;width:25%;font-size:11px;font-weight:600;'
                        f'color:#334;border-right:1px solid #eef0f5;vertical-align:top;">{fld["label"]}</td>'
                        f'<td style="padding:8px 14px;width:25%;font-size:13px;font-weight:700;'
                        f'color:#0d0d1a;">{disp}</td>'
                    )

                row_h = "<tr style='border-bottom:1px solid #f5f5f5;'>" + _cell(f)
                if i + 1 < len(sfl) and not sfl[i + 1].get("auto") and sfl[i + 1]["type"] != "textarea":
                    row_h += _cell(sfl[i + 1]) + "</tr>"
                    i += 2
                else:
                    row_h += "<td colspan='2'></td></tr>"
                    i += 1
                rows_html += row_h

        st.markdown(
            f'<table style="width:100%;border-collapse:collapse;background:white;'
            f'border-radius:8px;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,0.06);'
            f'margin-bottom:8px;">'
            f'<tbody>{rows_html}</tbody></table>',
            unsafe_allow_html=True,
        )


def show_review(user: dict, month_year: str, month_label: str):
    """Checker's read-only view of a maker's submission with approve / reject controls."""
    _dashboard_css()

    # Inject sub-header CSS (normally injected by show_section_form)
    st.markdown("""
    <style>
    .sub-header {
        background: linear-gradient(90deg,#001a6e 0%,#0033A0 60%,#0050d0 100%);
        color: white; font-size: 13px; font-weight: 700;
        padding: 9px 18px; border-radius: 10px; margin: 20px 0 12px;
        letter-spacing: 0.4px; box-shadow: 0 2px 8px rgba(0,51,160,0.18);
    }
    .auto-box {
        background: linear-gradient(135deg,#e8f0ff 0%,#dce8ff 100%);
        border: 1.5px solid #b0c8f8; border-radius: 10px;
        padding: 12px 16px; margin-bottom: 10px;
    }
    </style>
    """, unsafe_allow_html=True)

    maker_id = st.session_state.get("review_maker_id", "").strip()

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        sl = _assets().get("side_logo")
        if sl:
            st.markdown(
                f'<div style="margin:0;padding:0;line-height:0;width:100%;overflow:hidden;">'
                f'<img src="{sl}" style="width:100%;height:auto;display:block;'
                f'margin:0;padding:0;"></div>',
                unsafe_allow_html=True,
            )
        st.markdown(f"""
        <div style="padding:10px 18px 12px;border-bottom:2px solid #c62828;">
          <div style="color:#ff4d4d;font-size:11px;font-weight:700;
                      letter-spacing:1.5px;text-transform:uppercase;">REVIEW MODE</div>
          <div style="color:#ff9999;font-size:10px;margin-top:3px;">
            Reviewing: {maker_id or "—"}
          </div>
        </div>
        <div style="height:4px;"></div>
        """, unsafe_allow_html=True)
        back_lbl = "← Back to Zone View" if user.get("role") == "Zone" else \
                   "← Back to HQO View"  if user.get("role") == "Admin" else \
                   "← Back to Dashboard"
        if st.button(back_lbl, key="rev_back", use_container_width=True):
            st.session_state.selected_section = None
            st.rerun()

    _dash_header(user)

    if not maker_id:
        st.error("No submission selected. Go back to Dashboard and enter a Location Code to review.")
        return

    # ── Submission status ─────────────────────────────────────────────────────
    sd            = sheets.get_month_status(maker_id, month_year)
    status        = sd["status"]
    checker_notes = sd.get("checker_notes", "")
    icon, color, label = STATUS_META.get(status, ("⚪", "#8c9db5", status))

    st.markdown(f"""
    <div style="background:white;border-radius:14px;padding:14px 22px;
                box-shadow:0 2px 10px rgba(0,43,143,0.09);margin-bottom:14px;
                display:flex;align-items:center;justify-content:space-between;">
      <div>
        <div style="font-size:16px;font-weight:700;color:#002b8f;">
          &#128065; Reviewing Submission — Location <strong>{maker_id}</strong>
        </div>
        <div style="font-size:12px;color:#667085;margin-top:3px;">
          Period: <strong>{month_label}</strong>
        </div>
      </div>
      <span style="background:{color};color:white;padding:5px 18px;border-radius:20px;
                    font-size:12px;font-weight:600;white-space:nowrap;">{icon} {label}</span>
    </div>
    """, unsafe_allow_html=True)

    if checker_notes:
        st.info(f"**Previous Checker Note:** {checker_notes}")

    draft = sheets.load_draft(maker_id, month_year)

    # ── All 10 sections (inline — no expander to avoid Streamlit label artifacts) ─
    for sec_num, sec_name in SECTIONS:
        st.markdown(
            f'<div style="background:linear-gradient(90deg,#001a6e 0%,#0033A0 60%,'
            f'#0050d0 100%);color:white;font-size:14px;font-weight:700;'
            f'padding:11px 20px;border-radius:10px;margin:18px 0 8px;'
            f'letter-spacing:0.3px;box-shadow:0 2px 8px rgba(0,51,160,0.20);">'
            f'Section {sec_num} &nbsp;&mdash;&nbsp; {sec_name}'
            f'</div>',
            unsafe_allow_html=True,
        )
        _render_section_readonly(sec_num, draft)

        if sec_num == 3:
            rc_rows = sheets.load_detail_table(maker_id, month_year, "RAILWAY_CLAIMS")
            if rc_rows:
                ui   = _DETAIL_UI["RAILWAY_CLAIMS"]
                keys = list(ui["cols"].keys())
                lbl_map = {k: v["label"] for k, v in ui["cols"].items()}
                df_rc = pd.DataFrame(rc_rows, columns=keys).rename(columns=lbl_map)
                st.markdown(
                    '<div class="sub-header" style="font-size:12px;padding:7px 16px;">'
                    '&#128203; &nbsp; Railway Claims</div>'
                    + df_rc.to_html(index=False, escape=True, border=0, classes="mis-tbl"),
                    unsafe_allow_html=True,
                )
        elif sec_num == 5:
            # ── S5A M&I MIS data (read-only summary for checker) ─────────────
            st.markdown(
                '<div class="sub-header" style="font-size:12px;padding:7px 16px;">'
                '&#128220; &nbsp; S5A — Maintenance &amp; Inspection (M&amp;I) MIS</div>',
                unsafe_allow_html=True,
            )
            _MI_REVIEW_TABS = [
                ("MI_TANK_OUTAGE",    "S5A-1 Tank Outage",
                 ["tank_no","outage_for","planned_start","planned_end","actual_start","actual_end","current_status"]),
                ("MI_MAJOR_REPAIR",   "S5A-2 Major Repair",
                 ["tank_no","nature_of_repair","revenue_capex","current_status","etc_date"]),
                ("MI_VRU",            "S5A-3 VRU",
                 ["vru_operational","date_not_operating","ms_vol_recovered_kl","vru_uptime_pct"]),
                ("MI_AUDIT_2526",     "S5A-4 M&I Audit 25-26",
                 ["audit_date","no_recommendations","no_pending","external_score"]),
                ("MI_AUDIT_2627",     "S5A-5 M&I Audit 26-27",
                 ["audit_carried_out","audit_date","no_recommendations","no_pending","external_score"]),
                ("MI_TECH_AUDIT",     "S5A-6 Tech. Audit",
                 ["audit_date","no_recommendations","no_pending","ref_no"]),
                ("MI_EQUIP_BREAKDOWN","S5A-7 Equip. Breakdown",
                 ["equipment_name","equipment_details","start_date","current_status"]),
                ("MI_INT_PIPELINE",   "S5A-8 Int. Pipeline",
                 ["last_ut_date","last_hydrotest_date","last_dcvg_date","last_lrut_date"]),
                ("MI_EXT_PIPELINE",   "S5A-9 Ext. Pipeline",
                 ["pipeline_type","pipeline_details","product","length_metres","last_ut_date"]),
                ("MI_TANK_STATUS",    "S5A-10 Tank Status",
                 ["tank_no","tank_status","cleaning_completed_date","inspection_date","painting_date"]),
            ]
            _any_mi = False
            for _mi_key, _mi_label, _mi_cols in _MI_REVIEW_TABS:
                _mi_rows = sheets.load_mi_data(_mi_key, maker_id, month_year)
                if not _mi_rows:
                    continue
                _any_mi = True
                if _mi_rows[0].get("na_flag") == "Y":
                    st.markdown(
                        f'<div class="sub-header" style="font-size:12px;padding:7px 16px;">'
                        f'&#128203; &nbsp; {_mi_label} — Not Applicable</div>',
                        unsafe_allow_html=True,
                    )
                    continue
                # Filter columns that exist in the data
                _avail_cols = [c for c in _mi_cols if any(c in r for r in _mi_rows)]
                if not _avail_cols:
                    _avail_cols = _mi_cols
                _df_mi = pd.DataFrame(
                    [{c: r.get(c, "") for c in _avail_cols} for r in _mi_rows]
                )
                # Prettify column names
                _col_rename = {c: c.replace("_", " ").title() for c in _avail_cols}
                _df_mi.rename(columns=_col_rename, inplace=True)
                st.markdown(
                    f'<div class="sub-header" style="font-size:12px;padding:7px 16px;">'
                    f'&#128203; &nbsp; {_mi_label}</div>'
                    + _df_mi.to_html(index=False, escape=True, border=0, classes="mis-tbl"),
                    unsafe_allow_html=True,
                )
            if not _any_mi:
                st.info(
                    "S5A — M&I MIS data has not been submitted for this period. "
                    "The Maker should complete the M&I MIS form (accessible via the "
                    "'M&I MIS' button on the dashboard) before final submission."
                )
        elif sec_num == 10:
            for tk in ("IRR_DETAILS", "LEGAL_CASES"):
                tk_rows = sheets.load_detail_table(maker_id, month_year, tk)
                if tk_rows:
                    ui   = _DETAIL_UI[tk]
                    keys = list(ui["cols"].keys())
                    lbl_map = {k: v["label"] for k, v in ui["cols"].items()}
                    df_tk = pd.DataFrame(tk_rows, columns=keys).rename(columns=lbl_map)
                    st.markdown(
                        f'<div class="sub-header" style="font-size:12px;padding:7px 16px;">'
                        f'&#128203; &nbsp; {ui["title"]}</div>'
                        + df_tk.to_html(index=False, escape=True, border=0, classes="mis-tbl"),
                        unsafe_allow_html=True,
                    )

    # ── Checker decision panel ────────────────────────────────────────────────
    st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)
    st.markdown("""
    <div style="background:#fffde7;border:1.5px solid #f59e0b;border-radius:14px;
                padding:16px 22px;box-shadow:0 2px 8px rgba(245,158,11,0.12);margin-bottom:12px;">
      <div style="font-size:15px;font-weight:700;color:#92400e;">&#9878; Checker Decision</div>
    </div>
    """, unsafe_allow_html=True)

    # Checker decision controls only shown when Checker is reviewing their own location.
    # Zone/HQO users arrive with review_show_controls = False (read-only view).
    show_controls = st.session_state.pop("review_show_controls", True)

    if not show_controls:
        st.info("This is a read-only view. Use the Revision Request workflow to request corrections.")
        if status == "SUBMITTED":
            st.markdown('<div style="border-top:1px solid #eef0f6;margin:14px 0 10px;"></div>',
                        unsafe_allow_html=True)
            _rpt_key_ro = f"_mis_rpt_ro_{maker_id}_{month_year}"
            if st.button("📊  Generate MIS Report", key="btn_gen_mis_rpt_ro",
                         use_container_width=True):
                with st.spinner("Generating MIS Report…"):
                    try:
                        loc_info = sheets.get_maker_info(maker_id)
                        rpt = sheets.generate_filled_mis_report(
                            maker_id, month_year, loc_info, draft)
                        st.session_state[_rpt_key_ro] = rpt
                    except Exception as _ex:
                        st.error(f"Report error: {_ex}")
            if st.session_state.get(_rpt_key_ro):
                st.download_button(
                    label="⬇️  Download Filled MIS Report",
                    data=st.session_state[_rpt_key_ro],
                    file_name=f"MIS_Report_{maker_id}_{month_year.replace('-','_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key=f"dl_rpt_ro_{maker_id}_{month_year}",
                )
    else:
        # ── Approve + Reject — only when PENDING_REVIEW ──────────────────────
        if status == "SUBMITTED":
            st.info("This submission is already approved and locked.")
        elif status == "PENDING_REVIEW":
            # ── Pre-approval completeness check ──────────────────────────────
            from form_defs import get_skip_sections
            _maker_loc_type = sheets.get_loc_type(maker_id)
            _skip_secs      = get_skip_sections(_maker_loc_type)
            _required_secs  = {n for n, _ in SECTIONS if n not in _skip_secs}
            _dash            = sheets.get_dashboard_data(maker_id, month_year, _maker_loc_type)
            _secs_done       = set(_dash.get("secs_done", []))
            _incomplete      = sorted(_required_secs - _secs_done)
            _all_complete    = len(_incomplete) == 0

            if not _all_complete:
                st.error(
                    f"⛔  **Cannot Approve — {len(_incomplete)} section(s) incomplete:**  "
                    + ", ".join(f"S{n}" for n in _incomplete)
                    + "  |  The Maker must fill and save all sections before approval."
                )

            _rej_key = f"_chk_rej_{maker_id}_{month_year}"
            col_approve, col_reject = st.columns(2)

            with col_approve:
                if st.button("✅  Approve & Lock Submission",
                             use_container_width=True, type="primary", key="rev_approve",
                             disabled=not _all_complete):
                    with st.spinner("Approving…"):
                        maker_info = {"locName": maker_id, "zone": user.get("zone", "")}
                        res = sheets.approve_submission(
                            maker_id, month_year, user["userId"], draft, maker_info
                        )
                    if res["ok"]:
                        st.session_state.pop(f"target_status_{maker_id}_{month_year}", None)
                        st.session_state.pop(f"lock_{month_year}", None)
                        st.success(
                            f"✅ Submission from location **{maker_id}** approved and locked!"
                        )
                        st.rerun()
                    else:
                        st.error(f"Approval failed: {res['msg']}")

            with col_reject:
                _rej_open = st.session_state.get(_rej_key, False)
                _rej_label = "▲  Reject — hide form" if _rej_open else "❌  Reject — enter reason"
                if st.button(_rej_label, key="rev_reject_toggle",
                             use_container_width=True):
                    st.session_state[_rej_key] = not _rej_open
                    st.rerun()

            if st.session_state.get(_rej_key, False):
                st.markdown(
                    '<div style="background:#fff5f5;border:1.5px solid #f87171;'
                    'border-radius:10px;padding:14px 18px;margin-top:8px;">',
                    unsafe_allow_html=True,
                )
                reject_note = st.text_area(
                    "Rejection Note for Maker (required)",
                    key="rev_reject_note",
                    placeholder="Explain specifically what needs to be corrected…",
                    height=110,
                )
                if st.button("Confirm Rejection", key="rev_reject_btn",
                             use_container_width=True):
                    note = (reject_note or "").strip()
                    if not note:
                        st.error("Please enter a rejection note before confirming.")
                    else:
                        with st.spinner("Rejecting…"):
                            res = sheets.reject_submission(
                                maker_id, month_year, user["userId"], note
                            )
                        if res["ok"]:
                            st.session_state.pop(f"target_status_{maker_id}_{month_year}", None)
                            st.session_state.pop(f"lock_{month_year}", None)
                            st.session_state.pop(_rej_key, None)
                            st.success(
                                f"❌ Submission rejected. Maker **{maker_id}** "
                                "can now revise and resubmit."
                            )
                            st.rerun()
                        else:
                            st.error(f"Rejection failed: {res['msg']}")
                st.markdown("</div>", unsafe_allow_html=True)

        else:
            # IN_PROGRESS, REJECTED, NOT_STARTED — inform but allow Reset below
            st.info(
                f"Status is **{status.replace('_', ' ')}** — "
                "Approve/Reject require the maker to submit first. "
                "You can still Reset the draft below."
            )

        # ── Reset — available for any non-SUBMITTED status ───────────────────
        if status != "SUBMITTED":
            _rst_key = f"_chk_rst_{maker_id}_{month_year}"
            st.markdown("<div style='height:12px;'></div>", unsafe_allow_html=True)
            st.markdown(
                '<div style="background:#fff8e1;border:1.5px solid #f59e0b;'
                'border-radius:10px;padding:10px 18px;margin-bottom:6px;">'
                '<span style="font-size:13px;font-weight:700;color:#92400e;">'
                '&#9888;&nbsp; Maker Data Reset</span>'
                '<span style="font-size:11px;color:#78350f;margin-left:10px;">'
                'Use only when the submission needs to be cleared entirely for fresh entry'
                '</span></div>',
                unsafe_allow_html=True,
            )
            _rst_open = st.session_state.get(_rst_key, False)
            _rst_label = "▲  Reset — hide form" if _rst_open else \
                         "🔄  Reset Draft — clear all data for maker"
            if st.button(_rst_label, key="rev_reset_toggle", use_container_width=True):
                st.session_state[_rst_key] = not _rst_open
                st.rerun()

            if st.session_state.get(_rst_key, False):
                st.warning(
                    "This will permanently delete ALL entered data for this maker's "
                    f"**{month_year}** submission. The maker will need to re-enter "
                    "everything from scratch."
                )
                reset_reason = st.text_area(
                    "Reason for Reset (required)",
                    key="rev_reset_reason",
                    placeholder="Explain why all data needs to be cleared…",
                    height=100,
                )
                if st.button("Confirm Reset", key="rev_reset_btn",
                             use_container_width=True):
                    note = (reset_reason or "").strip()
                    if not note:
                        st.error("Please enter a reason for the reset.")
                    else:
                        with st.spinner("Resetting draft…"):
                            res = sheets.reset_draft(
                                maker_id, month_year, user["userId"], note
                            )
                        if res["ok"]:
                            st.session_state.pop(f"target_status_{maker_id}_{month_year}", None)
                            st.session_state.pop(f"lock_{month_year}", None)
                            st.session_state.pop(_rst_key, None)
                            st.success(
                                f"🔄 Draft for location **{maker_id}** has been reset. "
                                "Maker can now re-enter all data from scratch."
                            )
                            st.rerun()
                        else:
                            st.error(f"Reset failed: {res['msg']}")

    st.markdown("""
    <div style="margin-top:24px;padding:10px 4px;border-top:1px solid #dde3ed;
                display:flex;justify-content:space-between;font-size:11px;color:#aaa;">
      <span>&#169; 2026 Hindustan Petroleum Corporation Limited. All rights reserved.</span>
      <span>HPCL SOD &nbsp;·&nbsp; Supply Operations &amp; Distribution</span>
    </div>
    """, unsafe_allow_html=True)


# ── Dashboard ─────────────────────────────────────────────────────────────────

def _dash_header(user: dict):
    tb       = _assets().get("title_banner")
    tb_html  = (f'<img src="{tb}" style="height:60px;width:auto;'
                f'object-fit:contain;display:block;">' if tb else "")

    # ── Single unified banner: title LEFT · logo CENTRE · badge RIGHT ────────
    st.markdown(f"""
    <div style="background:linear-gradient(135deg,#001060 0%,#002b8f 55%,#003eb5 100%);
                border-radius:14px;padding:12px 22px;
                box-shadow:0 3px 14px rgba(0,43,143,0.28);margin-bottom:10px;
                display:flex;align-items:center;justify-content:space-between;
                min-height:80px;gap:12px;">

      <!-- Left: portal title -->
      <div style="flex:1 1 0;min-width:0;">
        <div style="font-size:15px;font-weight:800;color:#ffffff;
                    line-height:1.25;letter-spacing:-0.1px;white-space:nowrap;">
          HPCL SOD &mdash; MIS Entry Portal
        </div>
        <div style="font-size:10px;color:rgba(255,255,255,0.68);margin-top:3px;
                    letter-spacing:0.2px;">
          Supply, Operations &amp; Distribution
        </div>
      </div>

      <!-- Centre: HPCL logo / banner image -->
      <div style="flex:0 0 auto;display:flex;align-items:center;justify-content:center;
                  padding:0 8px;">
        {tb_html}
      </div>

      <!-- Right: location + role badge -->
      <div style="flex:1 1 0;display:flex;justify-content:flex-end;align-items:center;">
        <span style="display:inline-flex;align-items:center;gap:5px;
                     background:rgba(255,255,255,0.12);color:white;
                     border:1.5px solid rgba(255,255,255,0.30);
                     padding:6px 16px;border-radius:20px;
                     font-size:11px;font-weight:600;white-space:nowrap;
                     backdrop-filter:blur(4px);">
          &#128205; {user['locName']} &nbsp;|&nbsp; {user['role']}
        </span>
      </div>
    </div>""", unsafe_allow_html=True)


def _deadline_banner(dl: dict, month_label: str):
    u, days, dt = dl["urgency"], dl["days_left"], dl["date"]
    if u == "overdue":
        grad  = "linear-gradient(135deg,#b71c1c,#c62828)"
        txt   = "white"
        ic    = "⚠️"
        title = "OVERDUE"
        msg   = f"MIS for <strong>{month_label}</strong> was due {dt} — {abs(days)} day(s) past deadline. Submit immediately!"
        pill_bg = "rgba(255,255,255,0.22)"
        pill_txt = "white"
        pill_lbl = f"{abs(days)}d overdue"
    elif u == "urgent":
        grad  = "linear-gradient(135deg,#ea580c,#c2410c)"
        txt   = "white"
        ic    = "🔴"
        title = "URGENT"
        msg   = f"MIS for <strong>{month_label}</strong> due on {dt} — only <strong>{days} day(s)</strong> left!"
        pill_bg = "rgba(255,255,255,0.22)"
        pill_txt = "white"
        pill_lbl = f"{days}d left"
    elif u == "warning":
        grad  = "linear-gradient(135deg,#1565c0,#1976d2)"
        txt   = "white"
        ic    = "⚠️"
        title = "DUE SOON"
        msg   = f"MIS for <strong>{month_label}</strong> due on {dt}. {days} days remaining."
        pill_bg = "rgba(255,255,255,0.22)"
        pill_txt = "white"
        pill_lbl = f"{days}d left"
    else:
        grad  = "linear-gradient(135deg,#15803d,#166534)"
        txt   = "white"
        ic    = "✅"
        title = "ON TRACK"
        msg   = f"MIS for <strong>{month_label}</strong> due on {dt}. {days} days remaining."
        pill_bg = "rgba(255,255,255,0.22)"
        pill_txt = "white"
        pill_lbl = f"{days}d left"

    st.markdown(f"""
    <div style="background:{grad};border-radius:10px;padding:9px 16px;
                margin-bottom:10px;display:flex;align-items:center;
                justify-content:space-between;box-shadow:0 2px 8px rgba(0,0,0,0.15);">
      <div style="display:flex;align-items:center;gap:8px;">
        <span style="font-size:15px;">{ic}</span>
        <div>
          <div style="font-size:9px;font-weight:700;color:rgba(255,255,255,0.8);
                      text-transform:uppercase;letter-spacing:0.7px;">{title}</div>
          <div style="font-size:11px;font-weight:600;color:{txt};">{msg}</div>
        </div>
      </div>
      <span style="background:{pill_bg};color:{pill_txt};padding:3px 10px;
                   border-radius:20px;font-size:10px;font-weight:700;
                   white-space:nowrap;border:1px solid rgba(255,255,255,0.30);">
        {pill_lbl}
      </span>
    </div>""", unsafe_allow_html=True)


def _status_card(user: dict, month_label: str, data: dict):
    from form_defs import get_skip_sections
    status    = data["status"]
    pct       = min(max(float(data.get("completion_pct", 0)), 0), 100)
    locked    = data.get("is_locked", False)
    secs_done = set(data.get("secs_done", []))
    skip_secs = get_skip_sections(user.get("locType", "HPCL"))
    icon, color, label = STATUS_META.get(status, ("⚪", "#8c9db5", status))

    # Section pill builder — circular icon + gradient background
    def _sec_cell(num: int, name: str) -> str:
        if num in skip_secs:
            grad  = "linear-gradient(135deg,#f1f5f9,#e2e8f0)"
            bd    = "#cbd5e1"
            fg    = "#64748b"
            ic_bg = "#94a3b8"
            ic    = "N/A"
            ic_fs = "7px"
        elif num in secs_done:
            grad  = "linear-gradient(135deg,#dcfce7,#f0fdf4)"
            bd    = "#86efac"
            fg    = "#166534"
            ic_bg = "#22c55e"
            ic    = "✔"
            ic_fs = "9px"
        else:
            grad  = "linear-gradient(135deg,#fff1f2,#fff5f5)"
            bd    = "#fecaca"
            fg    = "#991b1b"
            ic_bg = "#ef4444"
            ic    = "✕"
            ic_fs = "9px"
        return (
            f'<td style="padding:3px 4px;width:50%;">'
            f'<div style="display:flex;align-items:center;gap:6px;padding:6px 9px;'
            f'background:{grad};border:1px solid {bd};border-radius:8px;">'
            f'<span style="display:inline-flex;align-items:center;justify-content:center;'
            f'width:17px;height:17px;border-radius:50%;background:{ic_bg};'
            f'color:white;font-size:{ic_fs};font-weight:800;flex-shrink:0;">{ic}</span>'
            f'<span style="font-size:11.5px;font-weight:700;color:#001060;'
            f'white-space:nowrap;">S{num}</span>'
            f'<span style="font-size:10.5px;color:{fg};font-weight:500;'
            f'white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">{name}</span>'
            f'</div></td>'
        )

    rows_html = "".join(
        f"<tr>{''.join(_sec_cell(SECTIONS[r*2][0], SECTIONS[r*2][1]) + _sec_cell(SECTIONS[r*2+1][0], SECTIONS[r*2+1][1]))}</tr>"
        for r in range(5)
    )

    incomplete = [f"S{n}" for n, _ in SECTIONS if n not in secs_done and n not in skip_secs]
    hint = ""
    if incomplete:
        hint = (
            f'<div style="margin-top:8px;padding:7px 12px;background:#fff5f5;'
            f'border-left:3px solid #ef4444;border-radius:7px;'
            f'font-size:11px;color:#991b1b;font-weight:600;">'
            f'&#9888;&nbsp; Pending: {", ".join(incomplete)}</div>'
        )

    lock_badge = (
        '<span style="display:inline-block;background:#e0e7ff;color:#3730a3;'
        'padding:2px 9px;border-radius:12px;font-size:10.5px;font-weight:700;'
        'margin-left:8px;">🔒 Locked</span>'
    ) if locked else ""

    pct_int = int(pct)
    bar_color = "#22c55e" if pct_int == 100 else ("#f59e0b" if pct_int >= 50 else "#ef4444")

    sc_html = (
        f'<div style="background:white;border-radius:12px;overflow:hidden;'
        f'box-shadow:0 2px 10px rgba(0,43,143,0.09);margin-bottom:10px;">'
        f'<div style="background:linear-gradient(135deg,#001060 0%,#002b8f 60%,#003eb5 100%);'
        f'padding:10px 18px 8px;">'
        f'<div style="display:flex;justify-content:space-between;align-items:center;">'
        f'<div>'
        f'<div style="font-size:14px;font-weight:700;color:white;line-height:1.25;">'
        f'{user["locName"]}</div>'
        f'<div style="font-size:10px;color:rgba(255,255,255,0.72);margin-top:2px;">'
        f'Zone: {user["zone"]} &nbsp;&middot;&nbsp; {month_label}</div>'
        f'</div>'
        f'<span style="background:rgba(255,255,255,0.18);color:white;'
        f'border:1px solid rgba(255,255,255,0.35);padding:3px 12px;border-radius:20px;'
        f'font-size:10px;font-weight:600;white-space:nowrap;">{icon} {label}</span>'
        f'</div></div>'
        f'<div style="padding:12px 18px 14px;">'
        f'<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:5px;">'
        f'<div style="font-size:9px;color:#667085;font-weight:600;'
        f'text-transform:uppercase;letter-spacing:0.4px;">Completion Progress{lock_badge}</div>'
        f'<div style="font-size:20px;font-weight:800;color:#002b8f;line-height:1;">'
        f'{pct_int}<span style="font-size:11px;font-weight:600;color:#667085;">%</span></div>'
        f'</div>'
        f'<div style="background:#e8edf8;border-radius:5px;height:7px;margin-bottom:12px;overflow:hidden;">'
        f'<div style="background:linear-gradient(90deg,{bar_color},{bar_color}cc);'
        f'width:{pct_int}%;height:7px;border-radius:5px;"></div>'
        f'</div>'
        f'<div style="font-size:9px;color:#667085;font-weight:700;'
        f'text-transform:uppercase;letter-spacing:0.4px;margin-bottom:5px;">Section Completion</div>'
        f'<table style="width:100%;border-collapse:separate;border-spacing:0 3px;">'
        f'{rows_html}</table>'
        f'{hint}'
        f'</div></div>'
    )
    st.markdown(sc_html, unsafe_allow_html=True)


def _action_area(user: dict, data: dict, month_year: str):
    role          = user["role"]
    status        = data["status"]
    locked        = data.get("is_locked", False)
    pct           = float(data.get("completion_pct", 0))
    checker_notes = data.get("checker_notes", "")

    st.markdown(
        '<div style="font-size:14px;font-weight:700;color:#333;margin:4px 0 10px;">Actions</div>',
        unsafe_allow_html=True,
    )

    if role == "Maker":
        # Show rejection note when data was sent back
        if status == "REJECTED" and checker_notes:
            st.warning(
                f"**Rejected by Checker:** {checker_notes}\n\n"
                "Please correct the data and resubmit."
            )

        c1, c2 = st.columns(2)
        with c1:
            if st.button("✏️  Enter / Edit Data",
                         disabled=(locked or status == "SUBMITTED"),
                         use_container_width=True, key="btn_edit"):
                st.session_state.selected_section = 1
                st.rerun()
        with c2:
            can_submit = (
                pct >= 100
                and status not in ("PENDING_REVIEW", "SUBMITTED", "LOCKED")
            )
            if st.button("📤  Submit for Review",
                         disabled=not can_submit,
                         use_container_width=True, key="btn_submit"):
                with st.spinner("Submitting for review…"):
                    res = sheets.submit_for_review(user["userId"], month_year)
                if res["ok"]:
                    st.session_state.pop(f"lock_{month_year}", None)
                    st.success("Submitted for Checker review! ✅")
                    st.rerun()
                else:
                    st.error(res["msg"])

        if not can_submit and pct < 100:
            st.caption(
                f"Complete all 10 sections (current: {int(pct)}%) to enable submission."
            )

        # ── Generate MIS Report — active when approved (SUBMITTED) ───────────
        if status == "SUBMITTED":
            st.markdown(
                '<div style="border-top:1px solid #eef0f6;margin:14px 0 10px;"></div>',
                unsafe_allow_html=True,
            )
            _rpt_key = f"_mis_rpt_{user['userId']}_{month_year}"
            if st.button("📊  Generate MIS Report", key="btn_gen_mis_rpt",
                         use_container_width=True):
                with st.spinner("Generating MIS Report…"):
                    try:
                        draft = sheets.load_draft(user["userId"], month_year)
                        rpt   = sheets.generate_filled_mis_report(
                            user["userId"], month_year, user, draft
                        )
                        st.session_state[_rpt_key] = rpt
                    except Exception as ex:
                        st.error(f"Report error: {ex}")
                        st.session_state[_rpt_key] = None
            if st.session_state.get(_rpt_key):
                st.download_button(
                    label="⬇️  Download Filled MIS Report",
                    data=st.session_state[_rpt_key],
                    file_name=f"MIS_Report_{user['userId']}_{month_year.replace('-','_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key=f"dl_rpt_{user['userId']}_{month_year}",
                )
            st.caption("✅ MIS approved and locked. Download your filled MIS Report above.")

        # ── Reset Draft — visible until submitted/pending review ──────────────
        can_reset = status not in ("SUBMITTED", "PENDING_REVIEW")
        if can_reset:
            mc = month_year.replace("-", "_")
            _rst_key = f"_mkr_rst_{mc}"
            st.markdown(
                '<div style="border-top:1px solid #eef0f6;margin:14px 0 10px;"></div>',
                unsafe_allow_html=True,
            )
            _rst_open = st.session_state.get(_rst_key, False)
            _rst_label = "▲  Hide Reset" if _rst_open else "🔄  Reset Draft — clear all data and start over"
            if st.button(_rst_label, key="btn_maker_reset_toggle",
                         use_container_width=True):
                st.session_state[_rst_key] = not _rst_open
                st.rerun()

            if st.session_state.get(_rst_key, False):
                st.warning(
                    "This will permanently delete **all entered data** for "
                    f"**{month_year}**. You will need to re-enter or re-upload "
                    "everything from scratch. This cannot be undone."
                )
                if st.button("Confirm — Reset All My Data",
                             key="btn_maker_reset_confirm",
                             use_container_width=True,
                             type="primary"):
                    with st.spinner("Resetting draft…"):
                        res = sheets.reset_draft(
                            user["userId"], month_year,
                            user["userId"], "Self-reset by maker"
                        )
                    if res["ok"]:
                        # Clear all draft session state for this month
                        for k in list(st.session_state.keys()):
                            if k.startswith(f"draft_{mc}_") or \
                               k.startswith(f"draft_loaded_{mc}"):
                                st.session_state.pop(k, None)
                        st.session_state.pop(f"lock_{month_year}", None)
                        st.session_state.pop(f"dash_{user['userId']}_{month_year}", None)
                        st.session_state.pop(_rst_key, None)
                        st.success("Draft reset. You can now re-enter your data.")
                        st.rerun()
                    else:
                        st.error(f"Reset failed: {res['msg']}")

    elif role == "Checker":
        # Checker is tied to the same location code as the Maker —
        # no manual input needed; use their own userId directly.
        target_id     = user["userId"]
        target_status = status   # already loaded via get_dashboard_data

        can_act = (target_status == "PENDING_REVIEW")
        c1, c2, c3 = st.columns(3)
        with c1:
            if st.button("👁️  Review Data",
                         use_container_width=True, key="btn_review"):
                st.session_state["review_maker_id"] = target_id
                st.session_state.selected_section   = "review"
                st.rerun()
        with c2:
            if st.button("✅  Approve & Lock",
                         disabled=not can_act,
                         use_container_width=True, key="btn_approve"):
                st.session_state["review_maker_id"] = target_id
                st.session_state.selected_section   = "review"
                st.rerun()
        with c3:
            if st.button("❌  Reject to Maker",
                         disabled=not can_act,
                         use_container_width=True, key="btn_reject"):
                st.session_state["review_maker_id"] = target_id
                st.session_state.selected_section   = "review"
                st.rerun()

        if target_status == "SUBMITTED":
            st.markdown(
                '<div style="border-top:1px solid #eef0f6;margin:14px 0 10px;"></div>',
                unsafe_allow_html=True,
            )
            _rpt_key = f"_mis_rpt_chk_{target_id}_{month_year}"
            if st.button("📊  Generate MIS Report", key="btn_gen_mis_rpt_chk",
                         use_container_width=True):
                with st.spinner("Generating MIS Report…"):
                    try:
                        draft = sheets.load_draft(target_id, month_year)
                        rpt   = sheets.generate_filled_mis_report(
                            target_id, month_year, user, draft
                        )
                        st.session_state[_rpt_key] = rpt
                    except Exception as ex:
                        st.error(f"Report error: {ex}")
                        st.session_state[_rpt_key] = None
            if st.session_state.get(_rpt_key):
                st.download_button(
                    label="⬇️  Download Filled MIS Report",
                    data=st.session_state[_rpt_key],
                    file_name=f"MIS_Report_{target_id}_{month_year.replace('-','_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key=f"dl_rpt_chk_{target_id}_{month_year}",
                )
            st.caption("✅ MIS approved. Generate and download the filled MIS Report above.")


@st.cache_data(ttl=None)
def _mis_guidelines_bytes(mtime: float) -> bytes | None:
    path = os.path.join(os.path.dirname(__file__), 'assets', 'MIS_Guidelines.pdf')
    try:
        with open(path, 'rb') as f:
            return f.read()
    except FileNotFoundError:
        return None


def _quick_links(user: dict, month_year: str, data: dict):
    # ── MIS Guidelines card with PDF download link ────────────────────────
    _gl = (
        '<div style="background:white;border-radius:16px;overflow:hidden;'
        'box-shadow:0 2px 14px rgba(0,43,143,0.08);margin-bottom:6px;">'
        '<div style="background:linear-gradient(135deg,#001060 0%,#002b8f 70%,#003eb5 100%);'
        'padding:12px 18px 10px;">'
        '<div style="font-size:14px;font-weight:700;color:white;letter-spacing:0.2px;">&#128203; MIS Guidelines</div>'
        '<div style="font-size:11px;color:rgba(255,255,255,0.72);margin-top:2px;">'
        'Complete instructions for filling all 10 sections &amp; submission rules</div>'
        '</div>'
        '<div style="padding:10px 16px 12px;">'
        '<div style="background:linear-gradient(135deg,#fff8e1,#fffde7);'
        'border-left:4px solid #f59e0b;border-radius:9px;padding:8px 13px;margin-bottom:8px;">'
        '<div style="font-size:10px;font-weight:700;color:#92400e;'
        'text-transform:uppercase;letter-spacing:0.6px;margin-bottom:2px;">&#128197; Submission Deadline</div>'
        '<div style="font-size:12px;color:#78350f;font-weight:600;">'
        'Submit by <strong>5th of every month</strong> for the preceding month</div>'
        '</div>'
        '<div style="font-size:11px;color:#64748b;">'
        'Click the button below to open the full MIS Guidelines PDF with section-by-section instructions.'
        '</div>'
        '</div>'
        '</div>'
    )
    st.markdown(_gl, unsafe_allow_html=True)
    _pdf_path = os.path.join(os.path.dirname(__file__), 'assets', 'MIS_Guidelines.pdf')
    _mtime = os.path.getmtime(_pdf_path) if os.path.exists(_pdf_path) else 0
    pdf_bytes = _mis_guidelines_bytes(_mtime)
    if pdf_bytes:
        st.download_button(
            label="📄  Open MIS Guidelines PDF",
            data=pdf_bytes,
            file_name="MIS_Guidelines.pdf",
            mime="application/pdf",
            use_container_width=True,
            key="dl_mis_guidelines",
        )
    else:
        st.info("MIS Guidelines PDF not found. Run gen_mis_guidelines_pdf.py to generate it.")

    # ── Excel Download ────────────────────────────────────────────────────
    st.markdown(
        '<div style="background:white;border-radius:16px;overflow:hidden;'
        'box-shadow:0 2px 14px rgba(0,43,143,0.08);border-top:3px solid #002b8f;">'
        '<div style="padding:14px 18px 8px;">'
        '<div style="display:flex;align-items:center;gap:9px;margin-bottom:4px;">'
        '<span style="background:linear-gradient(135deg,#002b8f,#003eb5);color:white;'
        'padding:4px 11px;border-radius:7px;font-size:12px;font-weight:700;">'
        '&#128229; Excel Template</span>'
        '</div>'
        '<div style="font-size:11.5px;color:#667085;margin-bottom:10px;">'
        'Download, fill offline, then upload to auto-populate all 135 MIS fields.'
        '</div>'
        '</div>'
        '</div>',
        unsafe_allow_html=True,
    )

    role   = user.get("role", "Maker")
    locked = data.get("is_locked", False)

    if role == "Maker":
        # Build template bytes (cached in session to avoid regenerating on every rerun)
        cache_key = f"_xlsx_v9_{user['userId']}_{month_year}"
        if cache_key not in st.session_state:
            with st.spinner("Building template…"):
                try:
                    draft_flat = sheets.load_draft(user["userId"], month_year)
                    st.session_state[cache_key] = sheets.generate_mis_template(
                        user["userId"], month_year, user, draft_flat,
                        loc_type=user.get("locType", "HPCL"),
                    )
                except Exception as ex:
                    st.session_state[cache_key] = None
                    st.error(f"Template error: {ex}")

        xlsx_bytes = st.session_state.get(cache_key)
        if xlsx_bytes:
            fname = f"MIS_{user['userId']}_{month_year.replace('-','_')}.xlsx"
            st.download_button(
                label="⬇️  Download MIS Template",
                data=xlsx_bytes,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"dl_{cache_key}",
            )

        # ── Upload (only when not locked) ───────────────────────────────
        if not locked:
            st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)

            # Track last processed file to prevent the re-upload loop:
            # st.rerun() keeps the file in the uploader widget, so without
            # this guard, upload processing fires on every render cycle.
            _up_done_key = f"_upload_done_{month_year}"

            up_file = st.file_uploader(
                "📤 Upload Filled Template",
                type=["xlsx"],
                key=f"upload_{month_year}",
                help="Upload the MIS Excel template you downloaded and filled offline.",
            )

            if up_file is not None:
                # Identify this exact file by name + size
                _file_id = f"{up_file.name}_{up_file.size}"

                if st.session_state.get(_up_done_key) == _file_id:
                    # Already processed — show result stored from last run
                    _prev = st.session_state.get(f"_upload_result_{month_year}", {})

                    # ── Section completion mini-grid ──────────────────────
                    _prev_secs = set(_prev.get("secs_done") or [])
                    if "secs_done" in _prev:
                        st.markdown(
                            '<div style="font-size:11px;font-weight:700;color:#555;'
                            'margin:6px 0 4px;text-transform:uppercase;letter-spacing:0.5px;">'
                            'Sections uploaded</div>',
                            unsafe_allow_html=True,
                        )
                        _ul_cells = []
                        for _sn, _sname in SECTIONS:
                            _done = _sn in _prev_secs
                            _ic   = "✅" if _done else "❌"
                            _fg   = "#1b5e20" if _done else "#b71c1c"
                            _bg   = "#f0faf0" if _done else "#fff5f5"
                            _ul_cells.append(
                                f'<span style="display:inline-block;margin:2px;'
                                f'padding:2px 8px;background:{_bg};color:{_fg};'
                                f'border-radius:6px;font-size:11px;font-weight:600;">'
                                f'{_ic} S{_sn}</span>'
                            )
                        st.markdown(
                            '<div style="line-height:2;">' + "".join(_ul_cells) + "</div>",
                            unsafe_allow_html=True,
                        )

                    if _prev.get("missing"):
                        with st.expander(
                            f"⚠️ {len(_prev['missing'])} required field(s) still empty",
                            expanded=False,
                        ):
                            for sec_name, flabel in _prev["missing"]:
                                st.markdown(
                                    f'<span style="color:#888;">{sec_name}</span>'
                                    f' → <strong style="color:#dc3545;">{flabel}</strong>',
                                    unsafe_allow_html=True,
                                )
                        st.info("Fill the missing fields in the app or re-upload a corrected template.")
                    if _prev.get("msg"):
                        st.success(_prev["msg"])
                else:
                    # ── New file: process it ─────────────────────────────
                    _prog = st.progress(0, text="📂 Reading uploaded file…")
                    try:
                        _file_bytes = up_file.read()
                        _prog.progress(20, text="🔍 Parsing MIS fields…")
                        parsed = sheets.parse_mis_upload(_file_bytes)
                        _prog.progress(60, text="💾 Saving to Google Sheets…")
                    except Exception as _up_ex:
                        _prog.empty()
                        st.error(f"Upload failed: {_up_ex}")
                        parsed = {"fields": {}, "errors": [str(_up_ex)],
                                  "railway_claims": [], "irr_details": [],
                                  "legal_cases": [], "mi_tabs": {}}

                    if parsed["errors"]:
                        for err in parsed["errors"]:
                            st.warning(err)

                    fields_from_xl = parsed["fields"]
                    if not fields_from_xl:
                        st.warning("No MIS field data found. "
                                   "Please use the downloaded template and fill row 4.")
                    else:
                        mc = month_year.replace("-", "_")

                        # ── Merge: template wins where filled;
                        #    existing app value kept where template is blank ──
                        overwritten, kept_app = [], []
                        for _sf in SECTION_FIELDS.values():
                            for f in _sf:
                                if f.get("auto"):
                                    continue
                                sk      = f"draft_{mc}_{f['key']}"
                                xl_val  = fields_from_xl.get(f["key"], "")
                                app_val = st.session_state.get(sk)
                                has_xl  = xl_val not in (None, "")
                                has_app = app_val not in (None, "")
                                if has_xl:
                                    if has_app and str(app_val) != str(xl_val):
                                        overwritten.append(f["label"])
                                    st.session_state[sk] = xl_val
                                elif has_app:
                                    kept_app.append(f["label"])

                        # ── Completeness check ───────────────────────────
                        missing_fields = []
                        for sn in sorted(SECTION_FIELDS):
                            for f in SECTION_FIELDS[sn]:
                                if f.get("auto") or not f.get("req"):
                                    continue
                                if st.session_state.get(f"draft_{mc}_{f['key']}") in (None, ""):
                                    missing_fields.append((SECTION_NAMES[sn], f["label"]))

                        # ── Section completion ───────────────────────────
                        # A section is complete when:
                        #   • it has required fields → ALL required fields are filled
                        #   • it has NO required fields → at least one editable field filled
                        secs_complete = []
                        for sn in sorted(SECTION_FIELDS):
                            sec_flds  = SECTION_FIELDS[sn]
                            req_flds  = [f for f in sec_flds
                                         if f.get("req") and not f.get("auto")]
                            edit_flds = [f for f in sec_flds if not f.get("auto")]
                            if req_flds:
                                if all(
                                    st.session_state.get(f"draft_{mc}_{f['key']}") not in (None, "")
                                    for f in req_flds
                                ):
                                    secs_complete.append(sn)
                            elif edit_flds:
                                if any(
                                    st.session_state.get(f"draft_{mc}_{f['key']}") not in (None, "")
                                    for f in edit_flds
                                ):
                                    secs_complete.append(sn)

                        # ── Persist to Google Sheets ─────────────────────
                        # Only send non-empty values so blank session-state
                        # keys don't overwrite previously saved GSheets data.
                        all_vals = {
                            f["key"]: st.session_state.get(f"draft_{mc}_{f['key']}")
                            for _sf in SECTION_FIELDS.values() for f in _sf
                            if not f.get("auto")
                            and st.session_state.get(f"draft_{mc}_{f['key']}") not in (None, "")
                        }
                        save_res = sheets.save_draft(
                            user["userId"], month_year,
                            field_data=all_vals,
                            sections_complete=secs_complete,
                        )
                        _prog.progress(90, text="✅ Finalising…")

                        if not save_res.get("ok"):
                            _prog.empty()
                            st.error(
                                f"Save failed: {save_res.get('msg', 'Unknown error')}. "
                                "Please try again."
                            )
                        else:
                            for tab_key, out_key in [
                                ("RAILWAY_CLAIMS", "railway_claims"),
                                ("IRR_DETAILS",    "irr_details"),
                                ("LEGAL_CASES",    "legal_cases"),
                            ]:
                                rows_dt = parsed.get(out_key, [])
                                if rows_dt:
                                    sheets.save_detail_table(
                                        user["userId"], month_year,
                                        tab_key, rows_dt, user
                                    )

                            # ── Save S5A M&I tab data from template ──────
                            mi_tabs_parsed = parsed.get("mi_tabs", {})
                            mi_tabs_saved  = 0
                            for mi_tab_key, mi_rows in mi_tabs_parsed.items():
                                if mi_rows == "NA":
                                    # Mark tab as NA
                                    sheets.save_mi_data(
                                        mi_tab_key, user["userId"], month_year,
                                        [{"na_flag": "Y"}])
                                    mi_tabs_saved += 1
                                elif isinstance(mi_rows, list) and mi_rows:
                                    sheets.save_mi_data(
                                        mi_tab_key, user["userId"], month_year,
                                        mi_rows)
                                    mi_tabs_saved += 1
                            # Invalidate M&I completion cache and force-reload all tab states
                            if mi_tabs_saved:
                                sheets.check_mi_complete.clear()
                                st.session_state.pop(
                                    f"_mi_comp_{user['userId']}_{month_year}", None)
                                # Clear per-tab loaded/rows/ctr/na flags so each tab
                                # reloads fresh from Google Sheets on next visit
                                for _tc in ("to","mr","vr","a25","a27","ta","eb","ip","ep","ts"):
                                    for _sk in ("_loaded","_rows","_ctr","_na"):
                                        st.session_state.pop(
                                            f"mi_{_tc}_{user['userId']}_{month_year}{_sk}", None)

                            # ── Clear stale caches ───────────────────────
                            for s in range(1, 11):
                                st.session_state.pop(f"draft_loaded_{mc}_s{s}", None)
                            st.session_state.pop(cache_key, None)
                            st.session_state.pop(f"lock_{month_year}", None)

                            # ── Build result summary ──────────────────────
                            n_xl   = len(fields_from_xl)
                            n_done = len(secs_complete)
                            rc = len(parsed.get("railway_claims", []))
                            ir = len(parsed.get("irr_details",    []))
                            lc = len(parsed.get("legal_cases",    []))
                            parts = [f"{n_xl} MIS fields loaded",
                                     f"{n_done}/10 sections complete"]
                            if rc or ir or lc:
                                parts.append(f"Railway Claims: {rc}  IRRs: {ir}  Legal Cases: {lc}")
                            if mi_tabs_saved:
                                parts.append(f"S5A M&I: {mi_tabs_saved}/10 tabs loaded")
                            if kept_app:
                                parts.append(f"{len(kept_app)} existing app value(s) retained")
                            if overwritten:
                                parts.append(f"{len(overwritten)} field(s) updated from template")
                            msg = "✅ Upload complete — " + "  ·  ".join(parts) + "."
                            _prog.progress(100, text="✅ Upload complete!")

                            # Store result then rerun once
                            st.session_state[_up_done_key] = _file_id
                            st.session_state[f"_upload_result_{month_year}"] = {
                                "missing":  missing_fields,
                                "secs_done": secs_complete,
                                "msg":      msg,
                            }
                            st.rerun()   # single clean rerun to refresh dashboard


def _section_grid():
    st.markdown(f'<p style="font-size:15px;font-weight:700;color:{HPCL_BLUE};'
                f'margin-top:18px;margin-bottom:10px;">MIS Sections</p>',
                unsafe_allow_html=True)
    for row_i in range(5):
        c_l, c_r = st.columns(2)
        for wc, si in ((c_l, row_i * 2), (c_r, row_i * 2 + 1)):
            num, name = SECTIONS[si]
            icon, color, label = STATUS_META["NOT_STARTED"]
            with wc:
                st.markdown(f"""
                <div class="sec-card" style="border-left-color:{color};">
                  <span style="font-weight:600;font-size:14px;color:#1a1a2e;">
                    <span style="color:{HPCL_BLUE};font-weight:800;">S{num}</span>&nbsp; {name}
                  </span>
                  <span class="sec-badge" style="background:{color};">{icon} {label}</span>
                </div>""", unsafe_allow_html=True)



def show_dashboard():
    _dashboard_css()
    user = st.session_state.user
    st.session_state.setdefault("selected_section", None)

    if flash := st.session_state.pop("flash", None):
        st.success(flash)

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        spb = _assets().get("side_panel_banner")
        # Full-width banner flush to top, no card/padding — sits on blue sidebar background
        if spb:
            st.markdown(
                f'<div style="margin:0;padding:0;line-height:0;width:100%;overflow:hidden;">'
                f'<img src="{spb}" style="width:100%;height:auto;display:block;'
                f'border-radius:0;margin:0;padding:0;"></div>',
                unsafe_allow_html=True,
            )

        # ── Dashboard home button (blue active when on main dashboard view) ──────
        _on_dash = st.session_state.get("selected_section") is None
        if st.button("🏠  Dashboard", key="btn_dash_home",
                     use_container_width=True, disabled=_on_dash):
            st.session_state.selected_section = None
            st.rerun()

        # ── MIS SECTIONS heading ────────────────────────────────────────────────
        st.markdown("""
        <div style="padding:10px 16px 4px;">
          <div style="color:#C8D7FF;font-size:9px;font-weight:700;
                      letter-spacing:2px;text-transform:uppercase;">MIS SECTIONS</div>
        </div>
        """, unsafe_allow_html=True)

        for num, name in SECTIONS:
            if st.button(f"S{num} - {name}", key=f"sid_{num}", use_container_width=True):
                st.session_state.selected_section = num
                st.rerun()
            if num == 5 and user.get("locType", "HPCL") == "HPCL":
                if st.button("↳ S5A - M&I MIS", key="btn_mi_mis_dash", use_container_width=True,
                             help="Maintenance & Inspection detailed MIS entry"):
                    st.session_state.selected_section = "mi_mis"
                    st.rerun()

        st.markdown('<div style="height:6px;"></div>', unsafe_allow_html=True)
        if st.button("📈  Analytics", key="btn_analytics_dash", use_container_width=True):
            st.session_state.selected_section = "analytics"
            st.rerun()

        if sheets.get_setting("chatbot_enabled", "FALSE") == "TRUE":
            st.markdown('<div style="height:4px;"></div>', unsafe_allow_html=True)
            if st.button("🤖  AI Assistant", key="btn_chatbot_dash", use_container_width=True):
                st.session_state.selected_section = "chatbot"
                st.rerun()

        st.markdown(
            '<div style="border-top:1px solid rgba(255,255,255,0.10);margin:8px 8px 0;'
            'padding:10px 8px 6px;">'
            '<div style="color:#C8D7FF;font-size:9px;font-weight:700;'
            'letter-spacing:1.5px;text-transform:uppercase;margin-bottom:5px;">'
            '&#128222; Need Help?</div>'
            '<div style="color:rgba(200,215,255,0.70);font-size:9px;margin-bottom:3px;">'
            'Email Support:</div>'
            '<a href="mailto:shoaibrehman@hpcl.in" '
            'style="color:#8AABFF;font-size:9.5px;font-weight:600;text-decoration:none;">'
            '&#128231; shoaibrehman@hpcl.in</a>'
            '</div>',
            unsafe_allow_html=True,
        )
        _show_ticket = st.session_state.get("show_ticket_form", False)
        if st.button(
            "🎫  Raise a Support Ticket",
            key="btn_ticket_toggle",
            use_container_width=True,
        ):
            st.session_state["show_ticket_form"] = not _show_ticket

        if st.session_state.get("show_ticket_form"):
            _tk_type = st.selectbox(
                "Issue Type",
                ["Login / Password Issue",
                 "Data Entry Problem",
                 "Submit / Approval Issue",
                 "Excel Upload Error",
                 "MIS Unlock Request",
                 "Portal Slow / Error",
                 "Others/Suggestions"],
                key="tk_type",
                label_visibility="collapsed",
            )
            _tk_desc = st.text_area(
                "Describe your issue",
                placeholder="Briefly describe the problem you are facing…",
                max_chars=750,
                key="tk_desc",
            )
            if st.button("📨  Submit Ticket", key="btn_submit_ticket",
                         use_container_width=True):
                if not _tk_desc or len(_tk_desc.strip()) < 10:
                    st.warning("Please describe the issue (at least 10 characters).")
                else:
                    _usr = st.session_state.get("user", {})
                    _res = sheets.log_help_request(
                        _usr.get("userId", "Unknown"),
                        _tk_desc.strip(),
                        _tk_type,
                    )
                    if _res["ok"]:
                        st.success(f"✅ {_res['msg']}")
                        st.caption(f"Ref: {_res.get('ref','—')}  |  We will reply to your Zone Officer.")
                        st.session_state["show_ticket_form"] = False
                    else:
                        st.error(_res["msg"])

    # ── Header (white card) ───────────────────────────────────────────────────
    _dash_header(user)

    # ── FY + Month + Logout row ───────────────────────────────────────────────
    fy_map = {"FY 2026-27": 2026}

    col_fy, col_mon, col_logout = st.columns([2, 3, 1])
    with col_logout:
        st.markdown("<div style='padding-top:27px;'></div>", unsafe_allow_html=True)
        if st.button("🚪 Logout", use_container_width=True):
            _lo_uid = (st.session_state.get("user") or {}).get("userId", "")
            if _lo_uid:
                sheets.clear_session(_lo_uid)
            st.session_state.clear(); st.rerun()
    with col_fy:
        sel_fy = st.selectbox("Financial Year", list(fy_map.keys()), key="sel_fy")

    fy_year = fy_map[sel_fy]
    with st.spinner(""):
        fy_res = sheets.get_fy_months(user["userId"], fy_year)
    months = fy_res.get("months", [])
    if not months:
        st.error(fy_res.get("msg", "Could not load month list. Please refresh."))
        return

    icon_map = {"NOT_STARTED":"⚪","IN_PROGRESS":"🔵","PENDING_REVIEW":"🟡","SUBMITTED":"✅","LOCKED":"🔒","REJECTED":"❌"}

    # Smart default: prefer current calendar month if it has any activity,
    # otherwise fall back to the latest month with activity (most recently touched).
    # This prevents jumping to a prior incomplete month when current month is started.
    if "sel_month" not in st.session_state:
        _cur_key = sheets.month_key()   # today's month key e.g. "Jun-2026"
        _cur_idx  = 0
        _best_idx = 0     # last non-NOT_STARTED month (fallback)
        _any_active = False
        for _mi, _mo in enumerate(months):
            if _mo["value"] == _cur_key:
                _cur_idx = _mi
            if _mo["status"] != "NOT_STARTED":
                _best_idx = _mi
                _any_active = True
        # Priority: current month if it has activity, else latest active, else current month
        _cur_status = months[_cur_idx]["status"] if months else "NOT_STARTED"
        if _cur_status != "NOT_STARTED":
            st.session_state["sel_month"] = _cur_idx          # current month is active → show it
        elif _any_active:
            st.session_state["sel_month"] = _best_idx         # latest prior active month
        else:
            st.session_state["sel_month"] = _cur_idx          # nothing active → default to today

    with col_mon:
        sel_idx = st.selectbox(
            "Month", range(len(months)),
            format_func=lambda i: (
                f"{icon_map.get(months[i]['status'], '⚪')}  {months[i]['label']}"
                f"{'  🔒' if months[i]['is_locked'] else ''}"
            ),
            key="sel_month",
        )

    selected = months[sel_idx]
    st.session_state["current_month"]       = selected["value"]
    st.session_state["current_month_label"] = selected["label"]

    with st.spinner(""):
        data = sheets.get_dashboard_data(user["userId"], selected["value"],
                                          user.get("locType", "HPCL"))
    if not data.get("ok"):
        st.error(data.get("msg", "Failed to load dashboard data."))
        return

    # ── Deadline alert ────────────────────────────────────────────────────────
    _deadline_banner(data["deadline"], selected["label"])

    # ── KPI strip ─────────────────────────────────────────────────────────────
    _kpi = [
        ("📅 Period",       selected["label"]),
        ("📊 Status",       STATUS_META.get(data["status"], ("","",""))[2]),
        ("✅ Completion",   f"{int(min(max(float(data.get('completion_pct',0)),0),100))}%"),
        ("📋 Sections Done",
         f"{len(data.get('secs_done',[]))}/10"),
    ]
    _kpi_cells = "".join(
        f'<div style="background:white;border-radius:10px;padding:7px 12px;'
        f'box-shadow:0 1px 5px rgba(0,31,94,0.08);text-align:center;flex:1;min-width:0;">'
        f'<div style="font-size:8px;color:#8899aa;font-weight:700;text-transform:uppercase;'
        f'letter-spacing:0.4px;margin-bottom:2px;">{lbl}</div>'
        f'<div style="font-size:12px;font-weight:700;color:#001F5E;white-space:nowrap;'
        f'overflow:hidden;text-overflow:ellipsis;">{val}</div>'
        f'</div>'
        for lbl, val in _kpi
    )
    st.markdown(
        f'<div style="display:flex;gap:8px;margin-bottom:10px;">{_kpi_cells}</div>',
        unsafe_allow_html=True,
    )

    # ── Main content: status+actions (left) + quick links (right) ─────────────
    col_main, col_links = st.columns([2.4, 1])
    with col_main:
        _status_card(user, selected["label"], data)
        _action_area(user, data, selected["value"])
    with col_links:
        _quick_links(user, selected["value"], data)

    # ── Footer ────────────────────────────────────────────────────────────────
    st.markdown("""
    <div style="margin-top:24px;padding:10px 4px;border-top:1px solid #dde3ed;
                display:flex;justify-content:space-between;
                font-size:11px;color:#aaa;">
      <span>© 2026 Hindustan Petroleum Corporation Limited. All rights reserved.</span>
      <span>HPCL SOD &nbsp;·&nbsp; Supply Operations & Distribution &nbsp;·&nbsp; Authorised Users Only</span>
    </div>
    """, unsafe_allow_html=True)


# ── Phase-6: Zone dashboard ──────────────────────────────────────────────────

def _month_selector_bar(user: dict, role_color: str):
    """Shared FY + Month + Logout bar. Returns (selected_month_value, selected_month_label)."""
    fy_map = {"FY 2026-27": 2026}

    col_fy, col_mon, col_logout = st.columns([2, 3, 1])
    with col_logout:
        st.markdown("<div style='padding-top:27px;'></div>", unsafe_allow_html=True)
        if st.button("🚪 Logout", use_container_width=True):
            _lo_uid = (st.session_state.get("user") or {}).get("userId", "")
            if _lo_uid:
                sheets.clear_session(_lo_uid)
            st.session_state.clear()
            st.rerun()
    with col_fy:
        sel_fy  = st.selectbox("Financial Year", list(fy_map.keys()), key="z_sel_fy")
    fy_year = fy_map[sel_fy]
    with st.spinner(""):
        fy_res = sheets.get_fy_months(user["userId"], fy_year)
    months = fy_res.get("months", [])
    if not months:
        st.error(fy_res.get("msg", "Could not load month list."))
        return None, None
    icon_map = {"NOT_STARTED":"⚪","IN_PROGRESS":"🔵","PENDING_REVIEW":"🟡",
                "SUBMITTED":"✅","LOCKED":"🔒","REJECTED":"❌"}
    with col_mon:
        sel_idx = st.selectbox(
            "Month", range(len(months)),
            format_func=lambda i: f"{icon_map.get(months[i]['status'],'⚪')}  {months[i]['label']}",
            key="z_sel_month",
        )
    return months[sel_idx]["value"], months[sel_idx]["label"]


def _zone_sidebar(user: dict, title: str, subtitle: str):
    with st.sidebar:
        spb = _assets().get("side_panel_banner")
        if spb:
            st.markdown(
                f'<div style="margin:0;padding:0;line-height:0;width:100%;overflow:hidden;">'
                f'<img src="{spb}" style="width:100%;height:auto;display:block;'
                f'border-radius:0;margin:0;padding:0;"></div>',
                unsafe_allow_html=True,
            )
        st.markdown(f"""
        <div style="padding:6px 16px 8px;">
          <div style="color:#C8D7FF;font-size:9px;font-weight:700;
                      letter-spacing:2px;text-transform:uppercase;">{title}</div>
          <div style="color:#8AABFF;font-size:9px;margin-top:2px;">{subtitle}</div>
        </div>
        """, unsafe_allow_html=True)

        if user.get("role") == "Admin":
            st.markdown(
                '<div style="padding:8px 4px 4px;font-size:11px;font-weight:700;'
                'color:#555;text-transform:uppercase;letter-spacing:0.8px;">'
                'Admin Tools</div>',
                unsafe_allow_html=True,
            )
            if st.button("⚙️ Setup Zone & HQO Accounts",
                         use_container_width=True,
                         help="Auto-generate Zone login IDs for all 16 zones and the SODSBU HQO account in UserAccess"):
                with st.spinner("Setting up accounts…"):
                    res = sheets.setup_zone_accounts()
                if res["ok"]:
                    added = res.get("added", [])
                    if added:
                        st.success(f"Added {len(added)} account(s): {', '.join(added)}")
                    else:
                        st.info("All zone and HQO accounts already exist — nothing to add.")
                else:
                    st.error(f"Setup failed: {res['msg']}")

            st.markdown('<div style="height:6px;"></div>', unsafe_allow_html=True)
            if st.button("🔧 Fix Conflicting Zone IDs",
                         use_container_width=True,
                         key="btn_fix_zone_ids"):
                _CORRECTIONS = [
                    ("North Zone",         "NORZONE", "NORMIS"),
                    ("North West Zone",    "NWZZONE", "NWZMIS"),
                    ("North Central Zone", "NCZZONE", "NCZMIS"),
                    ("South Zone",         "SOUZONE", "SOUMIS"),
                    ("South Central Zone", "SCZZONE", "SCZMIS"),
                ]
                errs, ok_list = [], []
                for zone_name, uid, pw in _CORRECTIONS:
                    r = sheets.upsert_zone_account(zone_name, uid, pw)
                    if r["ok"]:
                        ok_list.append(f"{uid} ({r['action']})")
                    else:
                        errs.append(f"{uid}: {r.get('msg','')}")
                if ok_list:
                    st.success("Updated: " + ", ".join(ok_list))
                if errs:
                    st.error("Errors: " + "; ".join(errs))

            st.markdown('<div style="height:6px;"></div>', unsafe_allow_html=True)
            if st.button("🔍 Audit Zone & HQO Accounts",
                         use_container_width=True,
                         key="btn_audit_accounts"):
                with st.spinner("Reading UserAccess…"):
                    accts = sheets.get_zone_admin_accounts()
                if accts:
                    import io as _io
                    rows_display = [
                        f"{'✅' if a['user_id'] else '❌'}  "
                        f"**{a['user_id']}** / {a['password']}"
                        f"  —  {a['zone']}  [{a['role']}]"
                        for a in sorted(accts, key=lambda x: x["role"] + x["zone"])
                    ]
                    st.markdown(
                        '<div style="font-size:12px;line-height:1.9;">'
                        + "<br>".join(rows_display)
                        + "</div>",
                        unsafe_allow_html=True,
                    )
                    st.caption(f"Total: {len(accts)} account(s)")
                else:
                    st.warning("No Zone or Admin accounts found in UserAccess.")

            st.markdown('<div style="height:6px;"></div>', unsafe_allow_html=True)
            st.markdown(
                '<div style="font-size:10px;color:rgba(255,255,255,0.55);'
                'padding:0 4px 3px;">Run once after updating M&I Separate Block.xlsx '
                'to populate Tank Master in Google Sheet. '
                'Required before deploying to cloud.</div>',
                unsafe_allow_html=True)
            if st.button("🗄️ Sync Tank Master → Google Sheet",
                         use_container_width=True,
                         key="btn_sync_tank_master"):
                with st.spinner("Reading Excel Tank Master and writing to Google Sheet…"):
                    res = sheets.sync_tank_master_to_sheet()
                if res["ok"]:
                    st.success(res["msg"])
                else:
                    st.error(f"Sync failed: {res['msg']}")

            st.markdown('<div style="height:6px;"></div>', unsafe_allow_html=True)
            st.markdown(
                '<div style="font-size:10px;color:rgba(255,255,255,0.55);'
                'padding:0 4px 3px;">Add Maker accounts for any locations in '
                'LocationMaster that are not yet in UserAccess. '
                'Default password = location code (users must change it).</div>',
                unsafe_allow_html=True)
            if st.button("➕ Sync Missing Location Accounts",
                         use_container_width=True,
                         key="btn_sync_loc_accounts"):
                with st.spinner("Reading LocationMaster and adding missing accounts…"):
                    try:
                        res = sheets.sync_missing_maker_accounts()
                    except Exception as _e:
                        res = {"ok": False, "msg": str(_e)}
                if res["ok"]:
                    added = res.get("added", [])
                    skipped = res.get("skipped", 0)
                    if added:
                        st.success(
                            f"Added {len(added)} new account(s): {', '.join(added)}. "
                            f"{skipped} already existed."
                        )
                    else:
                        st.info(f"All {skipped} locations already have accounts.")
                else:
                    st.error(f"Sync failed: {res['msg']}")

            st.markdown('<div style="height:6px;"></div>', unsafe_allow_html=True)
            st.markdown(
                '<div style="font-size:10px;color:rgba(255,255,255,0.55);'
                'padding:0 4px 3px;">Delete all MIS data for specific location '
                'codes (use before launch to wipe test data).</div>',
                unsafe_allow_html=True)
            _reset_codes = st.text_input(
                "Location codes to reset (comma-separated)",
                key="admin_reset_codes",
                placeholder="e.g. 1424, 1457, 1588",
            )
            if st.button("🗑️  Reset Location Data", key="btn_reset_loc_data",
                         use_container_width=True):
                codes = [c.strip() for c in _reset_codes.split(",") if c.strip()]
                if not codes:
                    st.warning("Enter at least one location code.")
                else:
                    msgs = []
                    for code in codes:
                        with st.spinner(f"Resetting {code}…"):
                            _r = sheets.reset_location_data(code)
                        msgs.append(f"**{code}**: {_r['msg']}" if _r["ok"]
                                    else f"**{code}**: ❌ {_r['msg']}")
                    st.success("\n\n".join(msgs))

            st.markdown('<div style="height:6px;"></div>', unsafe_allow_html=True)
            if st.button("📋  Review Email Recipients", key="btn_email_review",
                         use_container_width=True):
                st.session_state["selected_section"] = "email_review"
                st.rerun()

            st.markdown('<div style="height:6px;"></div>', unsafe_allow_html=True)
            st.markdown(
                '<div style="font-size:10px;color:rgba(255,255,255,0.55);'
                'padding:0 4px 3px;">Email MIS Portal login credentials to all '
                'Location In-charges via Outlook. Sends User ID + Password to each location.</div>',
                unsafe_allow_html=True)
            if st.button("📧  Load Locations for Email", key="btn_cred_load",
                         use_container_width=True):
                st.session_state["_cred_accounts"] = sheets.get_all_maker_credentials()
                st.rerun()
            if st.session_state.get("_cred_accounts"):
                import emails as _emails_mod
                _all_accounts = st.session_state["_cred_accounts"]
                _zone_opts    = ["All Zones"] + sorted({a["zone"] for a in _all_accounts if a.get("zone")})
                _cred_zone    = st.selectbox("Filter by Zone", _zone_opts, key="admin_cred_zone")
                _accounts     = _all_accounts if _cred_zone == "All Zones" else [
                    a for a in _all_accounts if a["zone"] == _cred_zone
                ]
                _email_map = _emails_mod.LOCATION_EMAIL_MAP
                _with_email = [(a, _email_map.get(a["userId"], ""))
                               for a in _accounts]
                _ok_count = sum(1 for _, e in _with_email if e)
                _miss = [a["userId"] for a, e in _with_email if not e]
                st.caption(
                    f"{_ok_count} / {len(_accounts)} locations have email addresses. "
                    + (f"No email: {', '.join(_miss[:5])}{'…' if len(_miss)>5 else ''}" if _miss else "")
                )
                _test_mode = st.toggle("🔒 Test mode — send only to me", value=True,
                                       key="cred_test_mode")
                if _test_mode:
                    st.caption(f"Test emails → shoaibrehman@hpcl.in")
                else:
                    st.markdown(
                        '<div style="background:#fdecea;border-left:3px solid #e53935;'
                        'padding:5px 8px;border-radius:4px;font-size:10px;color:#b71c1c;'
                        'font-weight:700;margin:2px 0;">⚠️ LIVE MODE — will send to real recipients</div>',
                        unsafe_allow_html=True,
                    )
                if st.button("📨  Send Credentials", key="btn_send_creds",
                             use_container_width=True, type="primary"):
                    if not _test_mode:
                        if not st.session_state.get("_loc_cred_confirmed"):
                            st.session_state["_loc_cred_confirmed"] = True
                            st.warning("⚠️ This will send LIVE emails to all real recipients. Click Send again to confirm.")
                            st.stop()
                    st.session_state.pop("_loc_cred_confirmed", None)
                    import emails as _em
                    sent, failed, skipped = 0, 0, 0
                    _errs = []
                    ok_email, _em_reason = _em.email_configured()
                    if not ok_email:
                        st.error(f"Outlook not accessible: {_em_reason}\n\nMake sure Microsoft Outlook is open and signed in, then retry.")
                    else:
                        for acct, to_email in _with_email:
                            if not to_email:
                                skipped += 1
                                continue
                            _res = _em.send_credential_email(
                                to_email=to_email,
                                loc_name=acct["locName"],
                                loc_code=acct["userId"],
                                password=acct["password"],
                                test_mode=_test_mode,
                                test_email=_em.SENDER_EMAIL,
                            )
                            if _res["ok"]:
                                sent += 1
                            else:
                                failed += 1
                                _errs.append(f"{acct['userId']}: {_res['msg']}")
                        st.success(
                            f"Sent: {sent} | Failed: {failed} | Skipped (no email): {skipped}"
                        )
                        if _errs:
                            st.error("\n".join(_errs[:5]))

            # ── Zone credential emails ────────────────────────────────────
            st.markdown('<div style="height:6px;"></div>', unsafe_allow_html=True)
            st.markdown(
                '<div style="font-size:10px;color:rgba(255,255,255,0.55);'
                'padding:0 4px 3px;">Send MIS Portal credentials to Zone ODs '
                '(BLRMIS, BHOMIS etc. → personal emails).</div>',
                unsafe_allow_html=True)
            _z_test_mode = st.toggle("🔒 Test mode — send only to me", value=True,
                                     key="zone_cred_test_mode")
            if _z_test_mode:
                st.caption("Test emails → shoaibrehman@hpcl.in")
            else:
                st.markdown(
                    '<div style="background:#fdecea;border-left:3px solid #e53935;'
                    'padding:5px 8px;border-radius:4px;font-size:10px;color:#b71c1c;'
                    'font-weight:700;margin:2px 0;">⚠️ LIVE MODE — will send to real Zone ODs</div>',
                    unsafe_allow_html=True,
                )
            if st.button("🏢  Send Zone Credentials", key="btn_send_zone_creds",
                         use_container_width=True):
                if not _z_test_mode:
                    if not st.session_state.get("_zone_cred_confirmed"):
                        st.session_state["_zone_cred_confirmed"] = True
                        st.warning("⚠️ This will send LIVE emails to all 16 Zone ODs. Click Send again to confirm.")
                        st.stop()
                st.session_state.pop("_zone_cred_confirmed", None)
                import emails as _em2
                ok_e, _em2_reason = _em2.email_configured()
                if not ok_e:
                    st.error(f"Outlook not accessible: {_em2_reason}\n\nOpen Microsoft Outlook and retry.")
                else:
                    _zone_accts = sheets.get_all_zone_credentials()
                    _z_cmap     = _em2.ZONE_CREDENTIAL_MAP
                    zs, zf, zsk = 0, 0, 0
                    _z_errs = []
                    for acct in _zone_accts:
                        zone_cfg = _z_cmap.get(acct["zone"], {})
                        to_email = zone_cfg.get("to", "")
                        cc_email = zone_cfg.get("cc", "") if not _z_test_mode else ""
                        if not to_email:
                            zsk += 1
                            continue
                        _res = _em2.send_credential_email(
                            to_email=to_email,
                            loc_name=acct["zone"],
                            loc_code=acct["userId"],
                            password=acct["password"],
                            cc_email=cc_email,
                            test_mode=_z_test_mode,
                            test_email=_em2.SENDER_EMAIL,
                        )
                        if _res["ok"]:
                            zs += 1
                        else:
                            zf += 1
                            _z_errs.append(f"{acct['userId']}: {_res['msg']}")
                    st.success(f"Zone emails — Sent: {zs} | Failed: {zf} | Skipped: {zsk}")
                    if _z_errs:
                        st.error("\n".join(_z_errs[:5]))

        # ── Chatbot feature toggle (Admin only) ───────────────────────────
        if user.get("role") == "Admin":
            st.markdown('<div style="height:4px;"></div>', unsafe_allow_html=True)
            chatbot_on = sheets.get_setting("chatbot_enabled", "FALSE") == "TRUE"
            new_state  = st.toggle(
                "🤖  AI Assistant",
                value=chatbot_on,
                key="chatbot_admin_toggle",
                help="Enable or disable the MIS AI Chatbot for all users",
            )
            if new_state != chatbot_on:
                res = sheets.set_setting("chatbot_enabled", "TRUE" if new_state else "FALSE",
                                         user["userId"])
                if res["ok"]:
                    st.success("AI Assistant " + ("enabled ✓" if new_state else "disabled"))
                else:
                    st.error(res.get("msg", "Could not save setting."))

        # ── Chatbot nav button (Zone + Admin when enabled) ────────────────
        if sheets.get_setting("chatbot_enabled", "FALSE") == "TRUE":
            st.markdown('<div style="height:4px;"></div>', unsafe_allow_html=True)
            if st.button("🤖  AI Assistant", key="btn_chatbot_zone", use_container_width=True):
                st.session_state.selected_section = "chatbot"
                st.rerun()

        st.markdown('<div style="height:8px;"></div>', unsafe_allow_html=True)
        if st.button("📈  Analytics", key="btn_analytics_nav", use_container_width=True):
            st.session_state.selected_section = "analytics"
            st.rerun()

        st.markdown('<div style="height:4px;"></div>', unsafe_allow_html=True)
        if st.button("📊  MIS Reports", key="btn_reports_nav", use_container_width=True):
            st.session_state.selected_section = "reports"
            st.rerun()


def _loc_table(rows: list, month_year: str, viewer_role: str,
               show_revision_btn: bool = False):
    """Render a table of location submission statuses with action buttons."""
    if not rows:
        st.info("No locations found.")
        return

    navigate_to_review = None
    for loc in rows:
        uid    = loc["userId"]
        name   = loc["locName"]
        zone   = loc.get("zone", "")
        status = loc.get("status", "NOT_STARTED")
        pct    = float(loc.get("completion_pct", 0))
        icon, color, label = STATUS_META.get(status, ("⚪", "#8c9db5", status))

        col_name, col_pct, col_status, col_act = st.columns([3, 1.2, 1.8, 2.5])
        with col_name:
            st.markdown(
                f'<div style="padding:10px 0 6px;">'
                f'<strong style="color:#002b8f;font-size:13px;">{name}</strong>'
                f'<br><span style="font-size:11px;color:#888;">{uid}'
                + (f' &nbsp;·&nbsp; {zone}' if viewer_role == "Admin" else "")
                + f'</span></div>',
                unsafe_allow_html=True,
            )
        with col_pct:
            st.markdown(
                f'<div style="padding:14px 0 0;">'
                f'<strong style="font-size:14px;color:#002b8f;">{int(pct)}%</strong></div>',
                unsafe_allow_html=True,
            )
        with col_status:
            st.markdown(
                f'<div style="padding:10px 0 0;">'
                f'<span style="background:{color};color:white;padding:3px 12px;'
                f'border-radius:12px;font-size:11px;font-weight:600;">'
                f'{icon} {label}</span></div>',
                unsafe_allow_html=True,
            )
        with col_act:
            is_submitted = (status == "SUBMITTED")
            # Column layout: View | [Revision] | [Generate MIS]
            if show_revision_btn and is_submitted:
                btn_cols = st.columns(3)
            elif is_submitted:
                btn_cols = st.columns(2)
            else:
                btn_cols = st.columns(1)

            with btn_cols[0]:
                if st.button("👁 View", key=f"view_{uid}_{month_year}",
                             use_container_width=True):
                    st.session_state["review_maker_id"]     = uid
                    st.session_state["review_show_controls"] = False
                    navigate_to_review = uid
            if show_revision_btn and is_submitted and len(btn_cols) > 1:
                with btn_cols[1]:
                    if st.button("🔄 Revision", key=f"rev_{uid}_{month_year}",
                                 use_container_width=True):
                        st.session_state[f"revision_open_{uid}_{month_year}"] = True
            if is_submitted:
                _rpt_key = f"_mis_rpt_zone_{uid}_{month_year}"
                mis_col_idx = 2 if (show_revision_btn and len(btn_cols) > 2) else 1
                with btn_cols[mis_col_idx]:
                    if st.button("📊 MIS", key=f"mis_{uid}_{month_year}",
                                 use_container_width=True):
                        with st.spinner(f"Generating MIS for {uid}…"):
                            try:
                                draft = sheets.load_draft(uid, month_year)
                                rpt   = sheets.generate_filled_mis_report(
                                    uid, month_year, loc, draft)
                                st.session_state[_rpt_key] = rpt
                            except Exception as _ex:
                                st.error(f"Report error: {_ex}")
                if st.session_state.get(_rpt_key):
                    st.download_button(
                        label="⬇️ Download MIS",
                        data=st.session_state[_rpt_key],
                        file_name=f"MIS_{uid}_{month_year}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_mis_{uid}_{month_year}",
                        use_container_width=True,
                    )

        # Inline revision request form
        if st.session_state.get(f"revision_open_{uid}_{month_year}"):
            with st.container():
                st.markdown(
                    f'<div style="background:#fff8e1;border:1px solid #f59e0b;'
                    f'border-radius:10px;padding:12px 16px;margin:4px 0 8px;">'
                    f'<strong style="color:#92400e;">Request Revision — {name} ({uid})</strong>'
                    f'</div>', unsafe_allow_html=True,
                )
                reason_key = f"reason_{uid}_{month_year}"
                reason = st.text_area("Reason for revision *",
                                      key=reason_key, height=80,
                                      placeholder="Describe the error that needs correction…")
                rc1, rc2 = st.columns([1, 1])
                with rc1:
                    if st.button("Submit Request", key=f"rr_submit_{uid}_{month_year}",
                                 type="primary", use_container_width=True):
                        _zu = st.session_state.get("_zone_user", {})
                        res = sheets.create_revision_request(
                            _zu.get("userId", ""),
                            uid, month_year, (reason or "").strip()
                        )
                        if res["ok"]:
                            st.success(res["msg"])
                            st.session_state.pop(f"revision_open_{uid}_{month_year}", None)
                            st.rerun()
                        else:
                            st.error(res["msg"])
                with rc2:
                    if st.button("Cancel", key=f"rr_cancel_{uid}_{month_year}",
                                 use_container_width=True):
                        st.session_state.pop(f"revision_open_{uid}_{month_year}", None)
                        st.rerun()

        st.markdown('<hr style="margin:2px 0;border:none;border-top:1px solid #f0f4fa;">', unsafe_allow_html=True)

    if navigate_to_review:
        st.session_state.selected_section = "review"
        st.rerun()


# ── Analytics Dashboard helpers ───────────────────────────────────────────────

def _an_monthly_agg(df, fy_months: list, col: str) -> list:
    """Return list of per-month mean values (float or None) for a DataFrame column."""
    if df is None or df.empty or col not in df.columns:
        return [None] * len(fy_months)
    out = []
    for m in fy_months:
        sub = df.loc[df["month_year"] == m, col].dropna()
        out.append(float(sub.mean()) if not sub.empty else None)
    return out


def _an_chart_layout(title: str, height: int = 300) -> dict:
    return dict(
        title=dict(text=title, font=dict(size=13, color="#002b8f")),
        height=height,
        margin=dict(l=10, r=10, t=45, b=40),
        paper_bgcolor="white",
        plot_bgcolor="#f8fafc",
        font=dict(family="Segoe UI, Arial, sans-serif", size=11),
        legend=dict(orientation="h", y=-0.3, font=dict(size=10)),
    )


def _an_financial_tab(df, fy_months: list, mlabels: list):
    import plotly.graph_objects as go

    if df.empty:
        st.info("No approved MIS data found for this period and scope.")
        return

    c1, c2 = st.columns(2)
    with c1:
        opex  = _an_monthly_agg(df, fy_months, "OPEX (Rs/MT)")
        o_tgt = _an_monthly_agg(df, fy_months, "OPEX Target (Rs/MT)")
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=mlabels, y=opex,  name="Actual",
                                  mode="lines+markers",
                                  line=dict(color="#dc3545", width=2.5),
                                  marker=dict(size=7)))
        fig.add_trace(go.Scatter(x=mlabels, y=o_tgt, name="Target",
                                  mode="lines+markers",
                                  line=dict(color="#198754", width=2, dash="dash"),
                                  marker=dict(size=6)))
        fig.update_layout(**_an_chart_layout("OPEX (Rs/MT) vs Target"))
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        meb_pct = _an_monthly_agg(df, fy_months, "MEB % w.r.t Budget")
        colors  = ["#dc3545" if (v or 0) > 100 else "#198754" for v in meb_pct]
        fig2 = go.Figure(go.Bar(x=mlabels, y=meb_pct, marker_color=colors, name="MEB %"))
        fig2.add_hline(y=100, line_dash="dot", line_color="#f59e0b",
                       annotation_text="Budget", annotation_position="top right",
                       annotation_font_size=10)
        fig2.update_layout(**_an_chart_layout("MEB % vs Budget"))
        st.plotly_chart(fig2, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        elec = _an_monthly_agg(df, fy_months, "Electricity Expenditure (Rs in Lakhs)")
        fig3 = go.Figure(go.Scatter(x=mlabels, y=elec, mode="lines+markers",
                                     fill="tozeroy", fillcolor="rgba(13,110,253,0.08)",
                                     line=dict(color="#0d6efd", width=2.5),
                                     marker=dict(size=7), name="Electricity ₹L"))
        fig3.update_layout(**_an_chart_layout("Electricity Expenditure (Rs in Lakhs)"))
        st.plotly_chart(fig3, use_container_width=True)

    with c4:
        capex = _an_monthly_agg(df, fy_months, "CAPEX (Lakhs)")
        aop   = _an_monthly_agg(df, fy_months, "Capex Target as per AOP (Lakhs)")
        fig4 = go.Figure()
        fig4.add_trace(go.Bar(x=mlabels, y=capex, name="CAPEX Actual",
                               marker_color="#6610f2"))
        fig4.add_trace(go.Bar(x=mlabels, y=aop, name="AOP Target",
                               marker_color="rgba(102,16,242,0.28)"))
        fig4.update_layout(**_an_chart_layout("CAPEX vs AOP (Lakhs)"), barmode="group")
        st.plotly_chart(fig4, use_container_width=True)


def _an_operational_tab(df, fy_months: list, mlabels: list):
    import plotly.graph_objects as go

    if df.empty:
        st.info("No approved MIS data found for this period and scope.")
        return

    c1, c2 = st.columns(2)
    with c1:
        ms_vals  = _an_monthly_agg(df, fy_months, "MS (MT)")
        hsd_vals = _an_monthly_agg(df, fy_months, "HSD (MT)")
        tgt_vals = _an_monthly_agg(df, fy_months, "Thruput Target (MT)")
        fig = go.Figure()
        fig.add_trace(go.Bar(x=mlabels, y=ms_vals,  name="MS",  marker_color="#0d6efd"))
        fig.add_trace(go.Bar(x=mlabels, y=hsd_vals, name="HSD", marker_color="#198754"))
        fig.add_trace(go.Scatter(x=mlabels, y=tgt_vals, name="Target",
                                  mode="lines+markers",
                                  line=dict(color="#dc3545", width=2, dash="dash"),
                                  marker=dict(size=6)))
        fig.update_layout(**_an_chart_layout("Throughput MT — MS / HSD vs Target"),
                           barmode="stack")
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        sec_vals = _an_monthly_agg(df, fy_months, "SEC (KWH/MT)")
        fig2 = go.Figure(go.Scatter(x=mlabels, y=sec_vals, mode="lines+markers",
                                     fill="tozeroy", fillcolor="rgba(255,193,7,0.1)",
                                     line=dict(color="#ffc107", width=2.5),
                                     marker=dict(size=7), name="SEC KWH/MT"))
        fig2.update_layout(**_an_chart_layout("SEC (KWH/MT) — Specific Energy Consumption"))
        st.plotly_chart(fig2, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        mdp_t = _an_monthly_agg(df, fy_months, "MDP Qty Target (KL)")
        mdp_a = _an_monthly_agg(df, fy_months, "MDP Qty Actual (KL)")
        fig3 = go.Figure()
        fig3.add_trace(go.Bar(x=mlabels, y=mdp_t, name="Target",
                               marker_color="rgba(13,110,253,0.3)"))
        fig3.add_trace(go.Bar(x=mlabels, y=mdp_a, name="Actual",
                               marker_color="#0d6efd"))
        fig3.update_layout(**_an_chart_layout("MDP Quantity — Target vs Actual (KL)"),
                            barmode="group")
        st.plotly_chart(fig3, use_container_width=True)

    with c4:
        ebp = _an_monthly_agg(df, fy_months, "EBP – Ethanol Blending Percentage")
        fig4 = go.Figure(go.Scatter(x=mlabels, y=ebp, mode="lines+markers",
                                     fill="tozeroy", fillcolor="rgba(25,135,84,0.1)",
                                     line=dict(color="#198754", width=2.5),
                                     marker=dict(size=7), name="EBP %"))
        fig4.update_layout(**_an_chart_layout("EBP — Ethanol Blending Percentage"))
        st.plotly_chart(fig4, use_container_width=True)


def _an_safety_tab(df, fy_months: list, mlabels: list):
    import plotly.graph_objects as go

    if df.empty:
        st.info("No approved MIS data found for this period and scope.")
        return

    c1, c2 = st.columns(2)
    with c1:
        hse = _an_monthly_agg(df, fy_months, "HSE Index vs Target")
        fig = go.Figure(go.Scatter(x=mlabels, y=hse, mode="lines+markers",
                                    fill="tozeroy", fillcolor="rgba(220,53,69,0.08)",
                                    line=dict(color="#dc3545", width=2.5),
                                    marker=dict(size=7), name="HSE Index"))
        fig.update_layout(**_an_chart_layout("HSE Index vs Target"))
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        pm = _an_monthly_agg(df, fy_months, "PM Percentage")
        fig2 = go.Figure(go.Bar(x=mlabels, y=pm,
                                 marker_color=["#198754" if (v or 0) >= 70 else "#f59e0b"
                                               for v in pm],
                                 name="PM %"))
        fig2.add_hline(y=70, line_dash="dot", line_color="#dc3545",
                       annotation_text="70% target", annotation_font_size=10)
        fig2.update_layout(**_an_chart_layout("PM Percentage"))
        st.plotly_chart(fig2, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        mi = _an_monthly_agg(df, fy_months, "M&I Index")
        fig3 = go.Figure(go.Scatter(x=mlabels, y=mi, mode="lines+markers",
                                     fill="tozeroy", fillcolor="rgba(108,117,125,0.1)",
                                     line=dict(color="#6c757d", width=2.5),
                                     marker=dict(size=7), name="M&I Index"))
        fig3.update_layout(**_an_chart_layout("M&I Index (Monthly)"))
        st.plotly_chart(fig3, use_container_width=True)

    with c4:
        swc = _an_monthly_agg(df, fy_months, "SWC (KL/MT)")
        fig4 = go.Figure(go.Scatter(x=mlabels, y=swc, mode="lines+markers",
                                     fill="tozeroy", fillcolor="rgba(13,202,240,0.1)",
                                     line=dict(color="#0dcaf0", width=2.5),
                                     marker=dict(size=7), name="SWC KL/MT"))
        fig4.update_layout(**_an_chart_layout("SWC — Specific Water Consumption (KL/MT)"))
        st.plotly_chart(fig4, use_container_width=True)


def _an_inventory_tab(df, fy_months: list, mlabels: list):
    import plotly.graph_objects as go

    if df.empty:
        st.info("No approved MIS data found for this period and scope.")
        return

    c1, c2 = st.columns(2)
    with c1:
        aim = _an_monthly_agg(df, fy_months, "AIM Holds (Nos.)")
        fig = go.Figure(go.Bar(x=mlabels, y=aim,
                                marker_color=["#dc3545" if (v or 0) > 0 else "#198754"
                                              for v in aim],
                                name="AIM Holds"))
        fig.update_layout(**_an_chart_layout("AIM Holds (Nos.) — 0 is ideal"))
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        ar = _an_monthly_agg(df, fy_months, "Auto-Reconciliation (% of Tanks on Auto Reco)")
        fig2 = go.Figure(go.Scatter(x=mlabels, y=ar, mode="lines+markers",
                                     fill="tozeroy", fillcolor="rgba(25,135,84,0.1)",
                                     line=dict(color="#198754", width=2.5),
                                     marker=dict(size=7), name="Auto-Reco %"))
        fig2.add_hline(y=100, line_dash="dot", line_color="#0d6efd",
                       annotation_text="100%", annotation_font_size=10)
        fig2.update_layout(**_an_chart_layout("Auto-Reconciliation (% of Tanks)"))
        st.plotly_chart(fig2, use_container_width=True)

    c3, _ = st.columns(2)
    with c3:
        ebp = _an_monthly_agg(df, fy_months, "EBP – Ethanol Blending Percentage")
        fig3 = go.Figure(go.Scatter(x=mlabels, y=ebp, mode="lines+markers",
                                     line=dict(color="#f59e0b", width=2.5),
                                     marker=dict(size=7), name="EBP %"))
        fig3.update_layout(**_an_chart_layout("EBP — Ethanol Blending Percentage"))
        st.plotly_chart(fig3, use_container_width=True)


def _an_compliance_tab(compliance_data: dict, fy_months: list, mlabels: list, user: dict):
    import plotly.graph_objects as go
    import pandas as pd

    role = user.get("role", "")

    if not compliance_data or not any(compliance_data.values()):
        st.info("No compliance data available for this period.")
        return

    # ── Monthly bar: % submitted ──────────────────────────────────────────────
    pct_submitted = []
    for m in fy_months:
        mon_data = compliance_data.get(m, {})
        if not mon_data:
            pct_submitted.append(None)
            continue
        n   = len(mon_data)
        sub = sum(1 for v in mon_data.values()
                  if v["status"] in ("SUBMITTED", "LOCKED", "PENDING_REVIEW"))
        pct_submitted.append(round(100 * sub / n) if n else None)

    fig_bar = go.Figure(go.Bar(
        x=mlabels, y=pct_submitted,
        marker_color=["#198754" if (v or 0) >= 80
                      else "#f59e0b" if (v or 0) >= 50
                      else "#dc3545" for v in pct_submitted],
        text=[f"{v}%" if v is not None else "—" for v in pct_submitted],
        textposition="outside",
        name="Compliance %",
    ))
    fig_bar.add_hline(y=80, line_dash="dot", line_color="#198754",
                      annotation_text="80%", annotation_font_size=10)
    fig_bar.update_layout(
        **_an_chart_layout("Monthly Compliance Rate — % Locations Submitted", height=280),
        yaxis=dict(range=[0, 115]),
    )
    st.plotly_chart(fig_bar, use_container_width=True)

    # ── Heatmap: location × month ─────────────────────────────────────────────
    status_score = {"SUBMITTED": 3, "LOCKED": 3, "PENDING_REVIEW": 2,
                    "IN_PROGRESS": 1, "REJECTED": 0, "NOT_STARTED": 0}

    all_uids = sorted({uid for m in fy_months for uid in compliance_data.get(m, {})})
    if not all_uids:
        return

    loc_labels = []
    matrix     = []
    hover_text = []
    for uid in all_uids:
        loc_info = compliance_data[fy_months[0]].get(uid, {}).get("loc", {})
        lname    = loc_info.get("locName", uid)
        if role == "Admin":
            loc_labels.append(f"{lname} ({loc_info.get('zone','?')})")
        else:
            loc_labels.append(lname)
        row_scores = []
        row_hover  = []
        for m in fy_months:
            entry  = compliance_data.get(m, {}).get(uid, {})
            status = entry.get("status", "NOT_STARTED")
            row_scores.append(status_score.get(status, 0))
            row_hover.append(f"{lname}<br>{m}<br>{status}")
        matrix.append(row_scores)
        hover_text.append(row_hover)

    colorscale = [
        [0.00, "#e9ecef"],
        [0.33, "#dc3545"],
        [0.67, "#f59e0b"],
        [1.00, "#198754"],
    ]
    fig_heat = go.Figure(go.Heatmap(
        z=matrix,
        x=mlabels,
        y=loc_labels,
        colorscale=colorscale,
        zmin=0, zmax=3,
        hovertemplate="%{text}<extra></extra>",
        text=hover_text,
        showscale=False,
    ))
    h_heat = max(300, 28 * len(all_uids) + 60)
    fig_heat.update_layout(
        title=dict(text="Submission Status Heatmap", font=dict(size=13, color="#002b8f")),
        height=h_heat,
        margin=dict(l=10, r=10, t=45, b=30),
        paper_bgcolor="white",
        xaxis=dict(side="top"),
        yaxis=dict(autorange="reversed", tickfont=dict(size=10)),
        font=dict(family="Segoe UI, Arial, sans-serif", size=11),
    )
    st.plotly_chart(fig_heat, use_container_width=True)

    # ── Legend ────────────────────────────────────────────────────────────────
    st.markdown(
        '<div style="display:flex;gap:16px;font-size:11px;margin-top:4px;">'
        '<span><span style="background:#198754;color:white;padding:2px 8px;'
        'border-radius:4px;">Submitted</span></span>'
        '<span><span style="background:#f59e0b;color:white;padding:2px 8px;'
        'border-radius:4px;">Pending Review</span></span>'
        '<span><span style="background:#dc3545;color:white;padding:2px 8px;'
        'border-radius:4px;">Not Filed / In Progress</span></span>'
        '<span><span style="background:#e9ecef;color:#555;padding:2px 8px;'
        'border-radius:4px;">No Data</span></span>'
        '</div>',
        unsafe_allow_html=True,
    )

    # ── Leaderboard table ─────────────────────────────────────────────────────
    st.markdown(
        '<div style="font-size:14px;font-weight:700;color:#002b8f;margin:20px 0 8px;">'
        'Location Compliance Leaderboard</div>',
        unsafe_allow_html=True,
    )
    rows_lb = []
    for uid in all_uids:
        loc_info   = compliance_data[fy_months[0]].get(uid, {}).get("loc", {})
        lname      = loc_info.get("locName", uid)
        zname      = loc_info.get("zone", "")
        submitted  = sum(1 for m in fy_months
                         if compliance_data.get(m, {}).get(uid, {}).get("status", "")
                         in ("SUBMITTED", "LOCKED", "PENDING_REVIEW"))
        total_mon  = len(fy_months)
        compliance = round(100 * submitted / total_mon) if total_mon else 0
        rows_lb.append({
            "Location": lname,
            "Zone":     zname,
            "Submitted": submitted,
            "Total Months": total_mon,
            "Compliance %": compliance,
        })
    rows_lb.sort(key=lambda r: r["Compliance %"], reverse=True)
    df_lb = pd.DataFrame(rows_lb)
    if role not in ("Admin", "Viewer"):
        df_lb = df_lb.drop(columns=["Zone"], errors="ignore")
    st.dataframe(df_lb, use_container_width=True, hide_index=True)


# ── Analytics Dashboard page ──────────────────────────────────────────────────

def show_analytics_page(user: dict):
    """FY-wise MIS performance analytics for Zone, Admin, Maker and Checker roles."""
    _dashboard_css()
    import pandas as pd

    role = user.get("role", "")
    zone = user.get("zone", "")

    if role in ("Zone", "Admin", "Viewer"):
        _zone_sidebar(user, "ANALYTICS", "Performance Dashboard")
    else:
        with st.sidebar:
            sl = _assets().get("side_logo")
            if sl:
                st.markdown(
                    f'<div style="margin:0;padding:0;line-height:0;width:100%;overflow:hidden;">'
                    f'<img src="{sl}" style="width:100%;height:auto;display:block;'
                    f'margin:0;padding:0;"></div>',
                    unsafe_allow_html=True,
                )
            st.markdown("""
            <div style="padding:10px 18px 12px;border-bottom:2px solid #c62828;">
              <div style="color:#ff4d4d;font-size:11px;font-weight:700;
                          letter-spacing:1.5px;text-transform:uppercase;">ANALYTICS</div>
              <div style="color:#ff9999;font-size:10px;margin-top:3px;">Performance Dashboard</div>
            </div>
            <div style="height:8px;"></div>
            """, unsafe_allow_html=True)
            if st.button("⬅ Back to Dashboard", key="an_back_dash", use_container_width=True):
                st.session_state.selected_section = None
                st.rerun()

    _dash_header(user)

    # ── FY selector + zone filter ─────────────────────────────────────────────
    today    = date.today()
    cur_fy   = today.year if today.month >= 4 else today.year - 1
    fy_opts  = [cur_fy, cur_fy - 1, cur_fy - 2]
    fy_labels = [f"FY {y}-{str(y + 1)[2:]}" for y in fy_opts]

    col_fy, col_zone, col_back = st.columns([2, 3, 1])
    with col_fy:
        sel_fy_label = st.selectbox("Financial Year", fy_labels, key="an_fy_sel")
        sel_fy       = fy_opts[fy_labels.index(sel_fy_label)]

    if role in ("Admin", "Viewer"):
        with col_zone:
            all_zones   = ["All Zones"] + sorted({
                l.get("zone", "") for l in sheets.get_all_maker_locations() if l.get("zone")
            })
            sel_zone       = st.selectbox("Zone Filter", all_zones, key="an_zone_sel")
            eff_role       = "Admin" if sel_zone == "All Zones" else "Zone"
            eff_zone       = "" if sel_zone == "All Zones" else sel_zone
    else:
        eff_role = role
        eff_zone = zone
        with col_zone:
            st.markdown(
                f'<div style="padding-top:26px;font-size:13px;color:#555;">'
                f'Zone: <strong>{zone}</strong></div>',
                unsafe_allow_html=True,
            )
    with col_back:
        st.markdown('<div style="height:22px;"></div>', unsafe_allow_html=True)
        if st.button("⬅ Back", key="an_back_top", use_container_width=True):
            st.session_state.selected_section = None
            st.rerun()

    fy_months = sheets._fy_month_years(sel_fy)
    mlabels   = [m.split("-")[0] for m in fy_months]

    # ── Load data ─────────────────────────────────────────────────────────────
    with st.spinner("Loading analytics data…"):
        compliance_data = sheets.get_compliance_analytics(eff_role, eff_zone, sel_fy)
        field_rows      = sheets.get_analytics_field_data(eff_role, eff_zone, sel_fy)

    # For Maker / Checker: keep all zone data for leaderboard but chart own loc only
    own_uid = user.get("userId", "")
    if role in ("Maker", "Checker") and field_rows:
        chart_rows = [r for r in field_rows if r["user_id"] == own_uid]
        if not chart_rows:
            chart_rows = field_rows
    else:
        chart_rows = field_rows

    df = pd.DataFrame(chart_rows) if chart_rows else pd.DataFrame()

    # ── KPI strip ─────────────────────────────────────────────────────────────
    total_slots = 0
    on_time     = 0
    missed      = 0
    perfect_mon = 0
    for m in fy_months:
        mon_data = compliance_data.get(m, {})
        if not mon_data:
            continue
        n   = len(mon_data)
        sub = sum(1 for v in mon_data.values()
                  if v["status"] in ("SUBMITTED", "LOCKED", "PENDING_REVIEW"))
        total_slots += n
        on_time     += sub
        missed      += n - sub
        if sub == n:
            perfect_mon += 1

    comp_rate  = round(100 * on_time / total_slots) if total_slots else 0
    n_locs     = len({uid for m in fy_months for uid in compliance_data.get(m, {})})

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Compliance Rate",     f"{comp_rate}%")
    k2.metric("Submitted Slots",     f"{on_time}",     delta=f"-{missed} missed")
    k3.metric("Perfect Months",      f"{perfect_mon}", help="Months all locations submitted")
    k4.metric("Locations Tracked",   f"{n_locs}")

    # ── Alert pills ───────────────────────────────────────────────────────────
    alerts = []
    if not df.empty:
        oc, ot = "OPEX (Rs/MT)", "OPEX Target (Rs/MT)"
        if oc in df.columns and ot in df.columns:
            mask = df[oc].notna() & df[ot].notna() & (df[ot] > 0)
            if mask.any() and (df.loc[mask, oc] > df.loc[mask, ot]).mean() > 0.3:
                alerts.append(("⚠️ OPEX Overshoot", "#dc3545"))
        mc = "MEB % w.r.t Budget"
        if mc in df.columns:
            vals = df[mc].dropna()
            if not vals.empty and (vals > 100).mean() > 0.3:
                alerts.append(("⚠️ MEB Over Budget", "#f59e0b"))
        tc, tt = "Total (MT) incl. Other Products", "Thruput Target (MT)"
        if tc in df.columns and tt in df.columns:
            mask = df[tc].notna() & df[tt].notna() & (df[tt] > 0)
            if mask.any() and (df.loc[mask, tc] < df.loc[mask, tt]).mean() > 0.3:
                alerts.append(("⚠️ Throughput Below Target", "#0d6efd"))

    if alerts:
        pills = "".join(
            f'<span style="background:{c};color:white;padding:4px 14px;border-radius:20px;'
            f'font-size:12px;font-weight:600;margin-right:8px;">{t}</span>'
            for t, c in alerts
        )
        st.markdown(f'<div style="margin:8px 0 16px;">{pills}</div>',
                    unsafe_allow_html=True)

    # ── 5 Tabs ────────────────────────────────────────────────────────────────
    tabs = st.tabs(["💰 Financial", "⚙️ Operational",
                    "🛡️ Safety & HSE", "📦 Inventory", "📋 Compliance"])

    with tabs[0]:
        _an_financial_tab(df, fy_months, mlabels)
    with tabs[1]:
        _an_operational_tab(df, fy_months, mlabels)
    with tabs[2]:
        _an_safety_tab(df, fy_months, mlabels)
    with tabs[3]:
        _an_inventory_tab(df, fy_months, mlabels)
    with tabs[4]:
        _an_compliance_tab(compliance_data, fy_months, mlabels, user)

    st.markdown("""
    <div style="margin-top:24px;padding:10px 4px;border-top:1px solid #dde3ed;
                display:flex;justify-content:space-between;font-size:11px;color:#aaa;">
      <span>&#169; 2026 Hindustan Petroleum Corporation Limited.</span>
      <span>HPCL SOD &nbsp;·&nbsp; Analytics Dashboard</span>
    </div>""", unsafe_allow_html=True)


def show_email_review(user: dict):
    """Admin-only full-page review of all email recipient mappings."""
    import emails as _em

    st.markdown(
        '<div style="font-size:22px;font-weight:800;color:#001F5E;margin-bottom:4px;">'
        '📧 Email Recipient Review</div>'
        '<div style="font-size:13px;color:#666;margin-bottom:18px;">'
        'Review all recipient mappings and sender details before sending credentials. '
        'Verify test mode is ON before sending.</div>',
        unsafe_allow_html=True,
    )
    if st.button("← Back to Dashboard", key="btn_er_back"):
        st.session_state.pop("selected_section", None)
        st.rerun()

    st.markdown("---")

    # ── Sender info ───────────────────────────────────────────────────────────
    st.markdown(
        f'<div style="background:#e8f5e9;border-left:4px solid #388e3c;padding:10px 16px;'
        f'border-radius:6px;margin-bottom:16px;font-size:13px;">'
        f'<b>Sender (From):</b> {_em.SENDER_EMAIL} &nbsp;·&nbsp; '
        f'<b>Test mode routes to:</b> {_em.SENDER_EMAIL}</div>',
        unsafe_allow_html=True,
    )

    tab_loc, tab_zone = st.tabs(["📍 Location Recipients (111)", "🏢 Zone OD Recipients (16)"])

    # ── Location tab ──────────────────────────────────────────────────────────
    with tab_loc:
        st.caption(
            "These are the email addresses that will receive credential emails "
            "for each plant. Locations without email are skipped automatically."
        )
        try:
            all_accts = sheets.get_all_maker_credentials()
        except Exception as exc:
            st.error(f"Could not load accounts: {exc}")
            all_accts = []

        em_map = _em.LOCATION_EMAIL_MAP
        rows = []
        for a in sorted(all_accts, key=lambda x: x.get("zone", "") + x.get("userId", "")):
            email = em_map.get(a["userId"], "")
            rows.append({
                "Zone":      a.get("zone", ""),
                "Code":      a["userId"],
                "Location":  a.get("locName", ""),
                "Email":     email if email else "— NO EMAIL —",
                "Status":    "✅" if email else "❌",
            })

        import pandas as _pd
        df = _pd.DataFrame(rows)
        no_email = df[df["Status"] == "❌"]
        with_email = df[df["Status"] == "✅"]

        col1, col2, col3 = st.columns(3)
        col1.metric("Total Locations", len(df))
        col2.metric("With Email", len(with_email))
        col3.metric("No Email (skipped)", len(no_email))

        if not no_email.empty:
            st.warning(f"These {len(no_email)} locations have no email and will be skipped: "
                       + ", ".join(no_email["Code"].tolist()))

        st.dataframe(
            df[["Zone", "Code", "Location", "Email", "Status"]],
            use_container_width=True, hide_index=True,
            column_config={
                "Status": st.column_config.TextColumn("", width="small"),
                "Code":   st.column_config.TextColumn("Code", width="small"),
            },
        )

    # ── Zone tab ──────────────────────────────────────────────────────────────
    with tab_zone:
        st.caption(
            "These are the Zone OD email addresses that will receive zone "
            "credential emails (BLRMIS, BHOMIS etc.)."
        )
        zone_rows = []
        for zone_name, cfg in sorted(_em.ZONE_CREDENTIAL_MAP.items()):
            zone_rows.append({
                "Zone":    zone_name,
                "To":      cfg.get("to", "— MISSING —"),
                "CC":      cfg.get("cc", ""),
                "Status":  "✅" if cfg.get("to") else "❌",
            })
        df_z = _pd.DataFrame(zone_rows)
        st.dataframe(
            df_z[["Zone", "To", "CC", "Status"]],
            use_container_width=True, hide_index=True,
            column_config={
                "Status": st.column_config.TextColumn("", width="small"),
            },
        )

    st.markdown("---")
    st.markdown(
        '<div style="background:#fff3e0;border-left:4px solid #f57c00;padding:10px 16px;'
        'border-radius:6px;font-size:12px;color:#7c4c00;">'
        '<b>⚠️ Before sending live credentials:</b><br>'
        '1. Run in <b>Test mode</b> first — one email goes to <b>shoaibrehman@hpcl.in</b>.<br>'
        '2. Verify the email looks correct (content, User ID, Password).<br>'
        '3. Toggle <b>Test mode OFF</b> only when ready. A confirmation click is required.<br>'
        '4. Location credentials and Zone credentials each have their <b>own independent test toggle</b>.</div>',
        unsafe_allow_html=True,
    )


def show_zone_dashboard(user: dict):
    """Dashboard for Zone-role users: see all locations in their zone."""
    _dashboard_css()

    _zone_sidebar(user, "ZONE VIEW", user.get("locName", user["zone"]))
    _dash_header(user)

    month_year, month_label = _month_selector_bar(user, "#7b2d8b")
    if not month_year:
        return

    st.session_state["current_month"]       = month_year
    st.session_state["current_month_label"] = month_label

    # ── Summary stats ─────────────────────────────────────────────────────────
    with st.spinner("Loading zone data…"):
        locs = sheets.get_locations_by_zone(user["zone"])
        rows = sheets.get_submissions_for_locations(locs, month_year)

    counts = {}
    for r in rows:
        counts[r["status"]] = counts.get(r["status"], 0) + 1

    submitted   = counts.get("SUBMITTED", 0)
    pending_rev = counts.get("PENDING_REVIEW", 0)
    in_prog     = counts.get("IN_PROGRESS", 0)
    not_started = counts.get("NOT_STARTED", 0) + counts.get("REJECTED", 0)
    total       = len(rows)

    st.markdown(f"""
    <div style="display:flex;gap:12px;margin-bottom:16px;flex-wrap:wrap;">
      <div style="flex:1;min-width:120px;background:white;border-radius:12px;padding:14px 18px;
                  box-shadow:0 2px 8px rgba(0,43,143,0.08);border-left:4px solid #198754;">
        <div style="font-size:11px;font-weight:700;color:#888;text-transform:uppercase;">Total</div>
        <div style="font-size:26px;font-weight:800;color:#002b8f;">{total}</div>
      </div>
      <div style="flex:1;min-width:120px;background:white;border-radius:12px;padding:14px 18px;
                  box-shadow:0 2px 8px rgba(0,43,143,0.08);border-left:4px solid #198754;">
        <div style="font-size:11px;font-weight:700;color:#888;text-transform:uppercase;">Submitted</div>
        <div style="font-size:26px;font-weight:800;color:#198754;">{submitted}</div>
      </div>
      <div style="flex:1;min-width:120px;background:white;border-radius:12px;padding:14px 18px;
                  box-shadow:0 2px 8px rgba(0,43,143,0.08);border-left:4px solid #f59e0b;">
        <div style="font-size:11px;font-weight:700;color:#888;text-transform:uppercase;">Pending Review</div>
        <div style="font-size:26px;font-weight:800;color:#f59e0b;">{pending_rev}</div>
      </div>
      <div style="flex:1;min-width:120px;background:white;border-radius:12px;padding:14px 18px;
                  box-shadow:0 2px 8px rgba(0,43,143,0.08);border-left:4px solid #0d6efd;">
        <div style="font-size:11px;font-weight:700;color:#888;text-transform:uppercase;">In Progress</div>
        <div style="font-size:26px;font-weight:800;color:#0d6efd;">{in_prog}</div>
      </div>
      <div style="flex:1;min-width:120px;background:white;border-radius:12px;padding:14px 18px;
                  box-shadow:0 2px 8px rgba(0,43,143,0.08);border-left:4px solid #8c9db5;">
        <div style="font-size:11px;font-weight:700;color:#888;text-transform:uppercase;">Pending / Not Started</div>
        <div style="font-size:26px;font-weight:800;color:#8c9db5;">{not_started}</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Location table ────────────────────────────────────────────────────────
    st.markdown(
        f'<div style="background:white;border-radius:12px;padding:18px 22px;'
        f'box-shadow:0 2px 8px rgba(0,43,143,0.08);margin-bottom:14px;">'
        f'<div style="font-size:15px;font-weight:700;color:#002b8f;margin-bottom:14px;">'
        f'Locations — {user.get("locName", user["zone"])} &nbsp;·&nbsp; {month_label}'
        f'</div></div>',
        unsafe_allow_html=True,
    )

    # Sort: Submitted last (already done), pending first
    order = {"PENDING_REVIEW": 0, "REJECTED": 1, "IN_PROGRESS": 2,
             "NOT_STARTED": 3, "SUBMITTED": 4}
    rows_sorted = sorted(rows, key=lambda r: order.get(r["status"], 9))

    # Pass user into revision form via closure (workaround: store in session)
    st.session_state["_zone_user"] = user
    _loc_table(rows_sorted, month_year, viewer_role="Zone", show_revision_btn=True)

    # ── My revision requests ──────────────────────────────────────────────────
    my_rr = [r for r in sheets.get_revision_requests(zone_filter=user["userId"])
             if r["status"] in ("PENDING_HQO", "APPROVED", "REJECTED")]
    if my_rr:
        st.markdown(
            '<div style="font-size:15px;font-weight:700;color:#002b8f;'
            'margin:20px 0 10px;">My Revision Requests</div>',
            unsafe_allow_html=True,
        )
        for rr in my_rr:
            st_col = {"PENDING_HQO": "#f59e0b", "APPROVED": "#198754", "REJECTED": "#dc3545"}
            bc     = st_col.get(rr["status"], "#8c9db5")
            st.markdown(
                f'<div style="background:white;border-radius:10px;padding:12px 16px;'
                f'border-left:4px solid {bc};margin-bottom:8px;'
                f'box-shadow:0 1px 4px rgba(0,0,0,0.06);">'
                f'<strong>#{rr["request_id"]}</strong> — '
                f'Location <strong>{rr["location_id"]}</strong> / {rr["month_year"]}'
                f'<span style="float:right;background:{bc};color:white;padding:2px 10px;'
                f'border-radius:10px;font-size:11px;">{rr["status"]}</span><br>'
                f'<span style="font-size:12px;color:#555;">{rr["reason"]}</span>'
                + (f'<br><em style="font-size:11px;color:#555;">HQO Note: {rr["notes"]}</em>'
                   if rr.get("notes") else "")
                + f'</div>',
                unsafe_allow_html=True,
            )

    # ── Zone-level downloads ──────────────────────────────────────────────────
    st.markdown(
        '<div style="border-top:1px solid #dde3ed;margin-top:20px;padding-top:14px;">'
        '<div style="font-size:14px;font-weight:700;color:#002b8f;margin-bottom:10px;">'
        '&#11015;&#65039; Downloads</div></div>',
        unsafe_allow_html=True,
    )
    _z_tm_key  = f"_z_tm_{user['userId']}_{month_year}"
    _z_mis_key = f"_z_mis_{user['userId']}_{month_year}"
    _zone_slug = user.get("zone", "Zone").replace(" ", "_")
    _zdc1, _zdc2 = st.columns(2)
    with _zdc1:
        if st.button("📊 Zone Tank Master", key=f"btn_z_tm_{month_year}",
                     use_container_width=True):
            with st.spinner("Preparing Zone Tank Master…"):
                try:
                    st.session_state[_z_tm_key] = sheets.get_full_tank_master_excel(
                        zone=user["zone"])
                except Exception as _ex:
                    st.error(f"Error: {_ex}")
                    st.session_state[_z_tm_key] = None
        if st.session_state.get(_z_tm_key):
            st.download_button(
                label="⬇️ Download Zone Tank Master",
                data=st.session_state[_z_tm_key],
                file_name=f"TankMaster_{_zone_slug}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"dl_z_tm_{month_year}",
            )
    with _zdc2:
        if st.button("📄 Zone Approved MIS", key=f"btn_z_mis_{month_year}",
                     use_container_width=True):
            with st.spinner("Preparing Zone Approved MIS…"):
                try:
                    st.session_state[_z_mis_key] = sheets.get_approved_mis_excel(
                        zone=user["zone"], month_year=month_year)
                except Exception as _ex:
                    st.error(f"Error: {_ex}")
                    st.session_state[_z_mis_key] = None
        if st.session_state.get(_z_mis_key):
            st.download_button(
                label="⬇️ Download Zone Approved MIS",
                data=st.session_state[_z_mis_key],
                file_name=f"ApprovedMIS_{_zone_slug}_{month_year.replace('-', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"dl_z_mis_{month_year}",
            )

    st.markdown("""
    <div style="margin-top:24px;padding:10px 4px;border-top:1px solid #dde3ed;
                display:flex;justify-content:space-between;font-size:11px;color:#aaa;">
      <span>&#169; 2026 Hindustan Petroleum Corporation Limited.</span>
      <span>HPCL SOD &nbsp;·&nbsp; Zone View</span>
    </div>""", unsafe_allow_html=True)


# ── Phase-6: HQO Admin dashboard ─────────────────────────────────────────────

def show_hqo_dashboard(user: dict):
    """Dashboard for HQO Admin (SODSBU) and Viewer: all zones, all locations."""
    _dashboard_css()

    is_viewer = user.get("role") == "Viewer"
    if is_viewer:
        _zone_sidebar(user, "HQO VIEW", "All Zones — View Only")
    else:
        _zone_sidebar(user, "HQO ADMIN", "All Zones — Full Access")
    _dash_header(user)

    month_year, month_label = _month_selector_bar(user, "#6d2077")
    if not month_year:
        return

    st.session_state["current_month"]       = month_year
    st.session_state["current_month_label"] = month_label

    # ── All-location summary ──────────────────────────────────────────────────
    with st.spinner("Loading all-India data…"):
        all_locs = sheets.get_all_maker_locations()
        all_rows = sheets.get_submissions_for_locations(all_locs, month_year)

    counts = {}
    for r in all_rows:
        counts[r["status"]] = counts.get(r["status"], 0) + 1

    total       = len(all_rows)
    submitted   = counts.get("SUBMITTED", 0)
    pend_rev    = counts.get("PENDING_REVIEW", 0)
    in_prog     = counts.get("IN_PROGRESS", 0)
    not_started = counts.get("NOT_STARTED", 0) + counts.get("REJECTED", 0)

    st.markdown(f"""
    <div style="display:flex;gap:12px;margin-bottom:16px;flex-wrap:wrap;">
      <div style="flex:1;min-width:100px;background:white;border-radius:12px;padding:14px 18px;
                  box-shadow:0 2px 8px rgba(0,43,143,0.08);border-left:4px solid #002b8f;">
        <div style="font-size:11px;font-weight:700;color:#888;text-transform:uppercase;">Total Locations</div>
        <div style="font-size:26px;font-weight:800;color:#002b8f;">{total}</div>
      </div>
      <div style="flex:1;min-width:100px;background:white;border-radius:12px;padding:14px 18px;
                  box-shadow:0 2px 8px rgba(0,43,143,0.08);border-left:4px solid #198754;">
        <div style="font-size:11px;font-weight:700;color:#888;text-transform:uppercase;">Submitted</div>
        <div style="font-size:26px;font-weight:800;color:#198754;">{submitted}</div>
      </div>
      <div style="flex:1;min-width:100px;background:white;border-radius:12px;padding:14px 18px;
                  box-shadow:0 2px 8px rgba(0,43,143,0.08);border-left:4px solid #f59e0b;">
        <div style="font-size:11px;font-weight:700;color:#888;text-transform:uppercase;">Pending Review</div>
        <div style="font-size:26px;font-weight:800;color:#f59e0b;">{pend_rev}</div>
      </div>
      <div style="flex:1;min-width:100px;background:white;border-radius:12px;padding:14px 18px;
                  box-shadow:0 2px 8px rgba(0,43,143,0.08);border-left:4px solid #0d6efd;">
        <div style="font-size:11px;font-weight:700;color:#888;text-transform:uppercase;">In Progress</div>
        <div style="font-size:26px;font-weight:800;color:#0d6efd;">{in_prog}</div>
      </div>
      <div style="flex:1;min-width:100px;background:white;border-radius:12px;padding:14px 18px;
                  box-shadow:0 2px 8px rgba(0,43,143,0.08);border-left:4px solid #8c9db5;">
        <div style="font-size:11px;font-weight:700;color:#888;text-transform:uppercase;">Not Filed</div>
        <div style="font-size:26px;font-weight:800;color:#8c9db5;">{not_started}</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── All locations with zone filter ────────────────────────────────────────
    all_zones = sorted({r["zone"] for r in all_rows if r.get("zone")})
    zone_opts = ["All Zones"] + all_zones
    sel_zone  = st.selectbox("Filter by Zone", zone_opts, key="hqo_zone_filter")

    filtered = all_rows if sel_zone == "All Zones" else [r for r in all_rows if r["zone"] == sel_zone]
    order    = {"PENDING_REVIEW": 0, "REJECTED": 1, "IN_PROGRESS": 2,
                "NOT_STARTED": 3, "SUBMITTED": 4}
    filtered_sorted = sorted(filtered, key=lambda r: order.get(r["status"], 9))

    st.markdown(
        f'<div style="font-size:15px;font-weight:700;color:#002b8f;'
        f'margin:8px 0 10px;">All Locations — {month_label}</div>',
        unsafe_allow_html=True,
    )
    _loc_table(filtered_sorted, month_year, viewer_role="Admin", show_revision_btn=False)

    # ── Pending revision requests — Admin only, not shown to Viewer ───────────
    if not is_viewer:
        pending_rr_hqo = [r for r in sheets.get_revision_requests() if r["status"] == "PENDING_HQO"]
        if pending_rr_hqo:
            st.markdown(
                f'<div style="background:#fffde7;border:1.5px solid #f59e0b;border-radius:12px;'
                f'padding:16px 22px;margin-top:16px;margin-bottom:16px;">'
                f'<div style="font-size:15px;font-weight:700;color:#92400e;margin-bottom:12px;">'
                f'&#9888; Pending Revision Requests ({len(pending_rr_hqo)})</div></div>',
                unsafe_allow_html=True,
            )
            for rr in pending_rr_hqo:
                rc1, rc2, rc3 = st.columns([5, 1.2, 1.2])
                with rc1:
                    st.markdown(
                        f'<div style="padding:8px 0;">'
                        f'<strong>#{rr["request_id"]}</strong> &nbsp;·&nbsp; '
                        f'Zone: <strong>{rr["zone_id"]}</strong> &nbsp;·&nbsp; '
                        f'Location: <strong>{rr["location_id"]}</strong> &nbsp;·&nbsp; '
                        f'Month: {rr["month_year"]}<br>'
                        f'<span style="font-size:12px;color:#555;">{rr["reason"]}</span>'
                        f'</div>', unsafe_allow_html=True,
                    )
                with rc2:
                    if st.button("✅ Approve", key=f"hqo_app_{rr['request_id']}",
                                 use_container_width=True, type="primary"):
                        res = sheets.approve_revision_request(rr["request_id"], user["userId"])
                        if res["ok"]:
                            st.success(f"Revision #{rr['request_id']} approved — location unlocked.")
                            st.rerun()
                        else:
                            st.error(res["msg"])
                with rc3:
                    if st.button("❌ Reject", key=f"hqo_rej_{rr['request_id']}",
                                 use_container_width=True):
                        st.session_state[f"hqo_reject_open_{rr['request_id']}"] = True

                if st.session_state.get(f"hqo_reject_open_{rr['request_id']}"):
                    rej_note = st.text_input(
                        "Rejection reason *",
                        key=f"hqo_rej_note_{rr['request_id']}",
                    )
                    if st.button("Confirm Reject", key=f"hqo_rej_confirm_{rr['request_id']}",
                                 type="primary", use_container_width=True):
                        res = sheets.reject_revision_request(rr["request_id"], rej_note, user["userId"])
                        if res["ok"]:
                            st.session_state.pop(f"hqo_reject_open_{rr['request_id']}", None)
                            st.success(f"Revision #{rr['request_id']} rejected.")
                            st.rerun()
                        else:
                            st.error(res["msg"])

    # ── HQO-level downloads ───────────────────────────────────────────────────
    st.markdown(
        '<div style="border-top:1px solid #dde3ed;margin-top:20px;padding-top:14px;">'
        '<div style="font-size:14px;font-weight:700;color:#002b8f;margin-bottom:10px;">'
        '&#11015;&#65039; Downloads</div></div>',
        unsafe_allow_html=True,
    )
    _dl_zone_filter = None if sel_zone == "All Zones" else sel_zone
    _hqo_zone_slug  = sel_zone.replace(" ", "_")
    _h_tm_key  = f"_h_tm_{month_year}_{sel_zone}"
    _h_mis_key = f"_h_mis_{month_year}_{sel_zone}"
    _hdc1, _hdc2 = st.columns(2)
    with _hdc1:
        if st.button("📊 Download Tank Master", key=f"btn_h_tm_{month_year}_{sel_zone}",
                     use_container_width=True):
            with st.spinner("Preparing Tank Master…"):
                try:
                    st.session_state[_h_tm_key] = sheets.get_full_tank_master_excel(
                        zone=_dl_zone_filter)
                except Exception as _ex:
                    st.error(f"Error: {_ex}")
                    st.session_state[_h_tm_key] = None
        if st.session_state.get(_h_tm_key):
            st.download_button(
                label=f"⬇️ Download Tank Master ({sel_zone})",
                data=st.session_state[_h_tm_key],
                file_name=f"TankMaster_{_hqo_zone_slug}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"dl_h_tm_{month_year}_{sel_zone}",
            )
    with _hdc2:
        if st.button("📄 Download Approved MIS", key=f"btn_h_mis_{month_year}_{sel_zone}",
                     use_container_width=True):
            with st.spinner("Preparing Approved MIS…"):
                try:
                    st.session_state[_h_mis_key] = sheets.get_approved_mis_excel(
                        zone=_dl_zone_filter, month_year=month_year)
                except Exception as _ex:
                    st.error(f"Error: {_ex}")
                    st.session_state[_h_mis_key] = None
        if st.session_state.get(_h_mis_key):
            st.download_button(
                label=f"⬇️ Download Approved MIS ({sel_zone})",
                data=st.session_state[_h_mis_key],
                file_name=f"ApprovedMIS_{_hqo_zone_slug}_{month_year.replace('-', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"dl_h_mis_{month_year}_{sel_zone}",
            )

    footer_label = "HQO View" if is_viewer else "HQO Admin View"
    st.markdown(f"""
    <div style="margin-top:24px;padding:10px 4px;border-top:1px solid #dde3ed;
                display:flex;justify-content:space-between;font-size:11px;color:#aaa;">
      <span>&#169; 2026 Hindustan Petroleum Corporation Limited.</span>
      <span>HPCL SOD &nbsp;·&nbsp; {footer_label}</span>
    </div>""", unsafe_allow_html=True)


# ── Phase-8/9: Reports page ──────────────────────────────────────────────────

def show_reports_page(user: dict):
    """Full-page MIS Reports & analytics view for Admin and Zone roles."""
    import emails as _emails

    _dashboard_css()

    role = user.get("role", "")
    _zone_sidebar(user, "REPORTS", "MIS Submission Analytics")

    if st.sidebar.button("← Back to Dashboard", key="rpt_back", use_container_width=True):
        st.session_state.selected_section = None
        st.rerun()

    _dash_header(user)

    # ── Month selector ────────────────────────────────────────────────────────
    today         = date.today()
    fy_start_year = today.year if today.month >= 4 else today.year - 1

    with st.spinner("Loading months…"):
        fy_res = sheets.get_fy_months(user["userId"], fy_start_year)
    months = fy_res.get("months", [])

    if not months:
        st.error(fy_res.get("msg", "Could not load month list."))
        return

    icon_map = {"NOT_STARTED": "⚪", "IN_PROGRESS": "🔵", "PENDING_REVIEW": "🟡",
                "SUBMITTED": "✅", "LOCKED": "🔒", "REJECTED": "❌"}

    col_mon, col_logout = st.columns([5, 1])
    with col_logout:
        st.markdown("<div style='padding-top:27px;'></div>", unsafe_allow_html=True)
        if st.button("🚪 Logout", key="rpt_logout", use_container_width=True):
            st.session_state.clear()
            st.rerun()
    with col_mon:
        sel_idx = st.selectbox(
            "Select Month",
            range(len(months)),
            format_func=lambda i: f"{icon_map.get(months[i]['status'], '⚪')}  {months[i]['label']}",
            key="rpt_sel_month",
        )

    month_year  = months[sel_idx]["value"]
    month_label = months[sel_idx]["label"]

    # ── Load data ─────────────────────────────────────────────────────────────
    with st.spinner("Loading submission data…"):
        if role == "Admin":
            all_rows = sheets.get_all_status_for_month(month_year)
        else:
            locs     = sheets.get_locations_by_zone(user["zone"])
            all_rows = sheets.get_submissions_for_locations(locs, month_year)

    # ── Compute due date (5th of next month) ──────────────────────────────────
    MONTH_ABBR = {m[:3]: i + 1 for i, m in enumerate(sheets.MONTHS_LONG)}
    parts      = month_year.split("-")
    month_num  = MONTH_ABBR.get(parts[0], 1)
    year_num   = int(parts[1]) if len(parts) > 1 else today.year
    if month_num == 12:
        due_date = date(year_num + 1, 1, 5)
    else:
        due_date = date(year_num, month_num + 1, 5)

    overdue_flag = today > due_date

    # ── Zone filter (Admin only) ──────────────────────────────────────────────
    if role == "Admin":
        all_zones = sorted({r["zone"] for r in all_rows if r.get("zone")})
        zone_filter = st.selectbox("Filter by Zone", ["All Zones"] + all_zones, key="rpt_zone_filter")
        display_rows = all_rows if zone_filter == "All Zones" else [r for r in all_rows if r["zone"] == zone_filter]
    else:
        display_rows = all_rows

    # ── Summary cards ─────────────────────────────────────────────────────────
    counts      = {}
    for r in all_rows:
        counts[r["status"]] = counts.get(r["status"], 0) + 1

    total       = len(all_rows)
    submitted   = counts.get("SUBMITTED", 0)
    pend_rev    = counts.get("PENDING_REVIEW", 0)
    not_sub     = total - submitted - pend_rev

    st.markdown(
        '<div style="background:linear-gradient(90deg,#001a6e,#0033A0,#0050d0);color:white;'
        'font-size:14px;font-weight:700;padding:11px 20px;border-radius:10px;margin:18px 0 8px;">'
        f'Submission Summary — {month_label}</div>',
        unsafe_allow_html=True,
    )
    st.markdown(f"""
    <div style="display:flex;gap:12px;margin-bottom:16px;flex-wrap:wrap;">
      <div style="flex:1;min-width:120px;background:white;border-radius:12px;padding:14px 18px;
                  box-shadow:0 2px 8px rgba(0,43,143,0.08);border-left:4px solid #002b8f;">
        <div style="font-size:11px;font-weight:700;color:#888;text-transform:uppercase;">Total Locations</div>
        <div style="font-size:30px;font-weight:800;color:#002b8f;">{total}</div>
      </div>
      <div style="flex:1;min-width:120px;background:white;border-radius:12px;padding:14px 18px;
                  box-shadow:0 2px 8px rgba(0,43,143,0.08);border-left:4px solid #198754;">
        <div style="font-size:11px;font-weight:700;color:#888;text-transform:uppercase;">Submitted</div>
        <div style="font-size:30px;font-weight:800;color:#198754;">{submitted}</div>
      </div>
      <div style="flex:1;min-width:120px;background:white;border-radius:12px;padding:14px 18px;
                  box-shadow:0 2px 8px rgba(0,43,143,0.08);border-left:4px solid #f59e0b;">
        <div style="font-size:11px;font-weight:700;color:#888;text-transform:uppercase;">Pending Review</div>
        <div style="font-size:30px;font-weight:800;color:#f59e0b;">{pend_rev}</div>
      </div>
      <div style="flex:1;min-width:120px;background:white;border-radius:12px;padding:14px 18px;
                  box-shadow:0 2px 8px rgba(0,43,143,0.08);border-left:4px solid #b71c1c;">
        <div style="font-size:11px;font-weight:700;color:#888;text-transform:uppercase;">Not Submitted</div>
        <div style="font-size:30px;font-weight:800;color:#b71c1c;">{not_sub}</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Styled table helper (local to this report page) ──────────────────────
    def _rpt_table(rows_data: list, center_cols: set = None,
                   status_col: str = None, num_color: dict = None) -> str:
        """Render a professional HTML table in HPCL theme."""
        if not rows_data:
            return ""
        center_cols = center_cols or set()
        num_color   = num_color   or {}   # col → "green" | "red"
        STATUS_BADGE = {
            "Submitted":      ("#e8f5e9", "#2e7d32"),
            "Pending Review": ("#fff8e1", "#e65100"),
            "In Progress":    ("#e3f2fd", "#1565c0"),
            "Not Filed":      ("#fdecea", "#c62828"),
            "Not Started":    ("#f5f5f5", "#616161"),
            "Rejected":       ("#fce4ec", "#b71c1c"),
        }
        TH = ("background:#002B8F;color:white;padding:10px 14px;font-size:12px;"
              "font-weight:700;letter-spacing:0.3px;white-space:nowrap;"
              "border-right:1px solid rgba(255,255,255,0.15);")
        cols = list(rows_data[0].keys())
        header = "".join(
            f'<th style="{TH}text-align:{"center" if c in center_cols else "left"};">{c}</th>'
            for c in cols
        )
        tbody = ""
        for i, row in enumerate(rows_data):
            bg = "#f7f9ff" if i % 2 == 1 else "#ffffff"
            cells = ""
            for c in cols:
                val = row[c]
                align = "center" if c in center_cols else "left"
                td = f"padding:9px 14px;border-bottom:1px solid #e8ecf4;text-align:{align};"
                if c == status_col:
                    sc = STATUS_BADGE.get(str(val), ("#f5f5f5", "#555"))
                    cells += (f'<td style="{td}">'
                              f'<span style="background:{sc[0]};color:{sc[1]};padding:3px 10px;'
                              f'border-radius:12px;font-size:11.5px;font-weight:700;">{val}</span>'
                              f'</td>')
                elif c in num_color and isinstance(val, (int, float)) and val > 0:
                    clr = "#2e7d32" if num_color[c] == "green" else "#c62828"
                    cells += f'<td style="{td}color:{clr};font-weight:700;">{val}</td>'
                elif c == cols[0]:   # first col = name/zone — bold blue
                    cells += f'<td style="{td}color:#1a237e;font-weight:600;">{val}</td>'
                else:
                    cells += f'<td style="{td}color:#333;">{val}</td>'
            tbody += f'<tr style="background:{bg};">{cells}</tr>'
        return (
            '<div style="overflow-x:auto;border-radius:10px;'
            'box-shadow:0 2px 8px rgba(0,43,143,0.10);margin-bottom:4px;">'
            '<table style="width:100%;border-collapse:collapse;'
            'font-family:Arial,sans-serif;font-size:13px;">'
            f'<thead><tr>{header}</tr></thead>'
            f'<tbody>{tbody}</tbody>'
            '</table></div>'
        )

    # ── Zone-wise summary table ───────────────────────────────────────────────
    if role == "Admin":
        st.markdown(
            '<div style="background:linear-gradient(90deg,#001a6e,#0033A0,#0050d0);color:white;'
            'font-size:14px;font-weight:700;padding:11px 20px;border-radius:10px;margin:18px 0 8px;">'
            'Zone-wise Summary</div>',
            unsafe_allow_html=True,
        )
        import pandas as pd
        zone_summary: dict = {}
        for r in all_rows:
            z = r.get("zone", "Unknown")
            if z not in zone_summary:
                zone_summary[z] = {"Zone": z, "Total": 0, "Submitted": 0,
                                   "Pending Review": 0, "In Progress": 0,
                                   "Not Filed": 0, "Overdue": 0}
            zone_summary[z]["Total"] += 1
            st_val = r.get("status", "NOT_STARTED")
            if st_val == "SUBMITTED":
                zone_summary[z]["Submitted"] += 1
            elif st_val == "PENDING_REVIEW":
                zone_summary[z]["Pending Review"] += 1
            elif st_val == "IN_PROGRESS":
                zone_summary[z]["In Progress"] += 1
            else:
                zone_summary[z]["Not Filed"] += 1
                if overdue_flag:
                    zone_summary[z]["Overdue"] += 1

        zone_rows = sorted(zone_summary.values(), key=lambda x: x["Zone"])
        st.markdown(
            _rpt_table(
                zone_rows,
                center_cols={"Total", "Submitted", "Pending Review",
                             "In Progress", "Not Filed", "Overdue"},
                num_color={"Submitted": "green", "Overdue": "red"},
            ),
            unsafe_allow_html=True,
        )
    else:
        import pandas as pd

    # ── Location-wise table ───────────────────────────────────────────────────
    st.markdown(
        '<div style="background:linear-gradient(90deg,#001a6e,#0033A0,#0050d0);color:white;'
        'font-size:14px;font-weight:700;padding:11px 20px;border-radius:10px;margin:18px 0 8px;">'
        'Location-wise Status</div>',
        unsafe_allow_html=True,
    )

    filter_cols = st.columns([2, 2, 3])
    with filter_cols[0]:
        status_opts = ["All Statuses", "Submitted", "Pending Review", "In Progress", "Not Filed"]
        sel_status  = st.selectbox("Filter by Status", status_opts, key="rpt_status_filter")

    status_key_map = {
        "Submitted":      ["SUBMITTED"],
        "Pending Review": ["PENDING_REVIEW"],
        "In Progress":    ["IN_PROGRESS"],
        "Not Filed":      ["NOT_STARTED", "REJECTED"],
    }

    loc_rows = display_rows
    if sel_status != "All Statuses":
        wanted   = status_key_map.get(sel_status, [])
        loc_rows = [r for r in loc_rows if r.get("status") in wanted]

    loc_records = []
    for r in loc_rows:
        st_val = r.get("status", "NOT_STARTED")
        _, _, st_label = STATUS_META.get(st_val, ("", "#8c9db5", st_val.replace("_", " ").title()))
        loc_records.append({
            "Location Code": r.get("userId", ""),
            "Location Name": r.get("locName", ""),
            "Zone":          r.get("zone", ""),
            "Status":        st_label,
            "Done %":        int(float(r.get("completion_pct", 0))),
        })

    if loc_records:
        st.markdown(
            _rpt_table(
                loc_records,
                center_cols={"Location Code", "Done %"},
                status_col="Status",
            ),
            unsafe_allow_html=True,
        )
        st.caption(f"{len(loc_records)} location(s) shown")
    else:
        st.info("No locations match the selected filters.")

    # ── Downloads ─────────────────────────────────────────────────────────────
    st.markdown(
        '<div style="background:linear-gradient(90deg,#001a6e,#0033A0,#0050d0);color:white;'
        'font-size:14px;font-weight:700;padding:11px 20px;border-radius:10px;margin:18px 0 8px;">'
        'Downloads</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div style="background:white;border-radius:12px;padding:14px 18px;'
        'box-shadow:0 2px 8px rgba(0,43,143,0.08);border-left:4px solid #002b8f;margin-bottom:12px;">',
        unsafe_allow_html=True,
    )

    dl_col1, dl_col2 = st.columns(2)

    with dl_col1:
        if st.button("⬇ Generate Submitted MIS Data (Excel)", key="rpt_gen_submitted",
                     use_container_width=True):
            with st.spinner("Building Excel…"):
                excel_bytes = sheets.download_submitted_data_excel(month_year)
            if excel_bytes:
                st.download_button(
                    label="Download Submitted MIS Data",
                    data=excel_bytes,
                    file_name=f"MIS_Submitted_{month_year}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="rpt_dl_submitted",
                )
            else:
                st.warning("No submitted data found for this month.")

    with dl_col2:
        if st.button("⬇ Generate Pending Locations List (Excel)", key="rpt_gen_pending",
                     use_container_width=True):
            pending_for_dl = [
                {**r, "remark": "Overdue" if (overdue_flag and r.get("status") != "SUBMITTED") else ""}
                for r in all_rows if r.get("status") != "SUBMITTED"
            ]
            with st.spinner("Building Excel…"):
                pend_bytes = sheets.download_pending_list_excel(pending_for_dl, month_year)
            st.download_button(
                label="Download Pending Locations List",
                data=pend_bytes,
                file_name=f"Pending_Locations_{month_year}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="rpt_dl_pending",
            )

    st.markdown("</div>", unsafe_allow_html=True)

    # ── Email reminders (Admin only) ──────────────────────────────────────────
    if role == "Admin":
        pending_zones = sorted({
            r["zone"] for r in all_rows
            if r.get("status") != "SUBMITTED" and r.get("zone") in _emails.ZONE_EMAIL_MAP
        })

        st.markdown(
            '<div style="background:linear-gradient(90deg,#001a6e,#0033A0,#0050d0);color:white;'
            'font-size:14px;font-weight:700;padding:11px 20px;border-radius:10px;margin:18px 0 8px;">'
            'Email Reminders</div>',
            unsafe_allow_html=True,
        )

        _email_ok, _email_err = _emails.email_configured()
        if not _email_ok:
            if _email_err == "local_only":
                # Running on cloud / Linux — email is a local-only feature
                st.markdown(
                    '<div style="background:#e8f4fd;border:1.5px solid #1565C0;border-radius:12px;'
                    'padding:14px 20px;margin-bottom:12px;">'
                    '<div style="font-size:13px;font-weight:700;color:#0d47a1;margin-bottom:6px;">'
                    '&#x2709;&nbsp; Email via Outlook — Local Feature</div>'
                    '<div style="font-size:13px;color:#1a237e;line-height:1.8;">'
                    'Reminder emails are sent through <strong>Microsoft Outlook</strong> running on '
                    'your Windows PC. To send emails, run the app <strong>locally</strong> with '
                    'Outlook open and signed in as <strong>shoaibrehman@hpcl.in</strong>.<br><br>'
                    'No email configuration is needed — emails go automatically from your HPCL '
                    'Outlook account.'
                    '</div></div>',
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    f'<div style="background:#fff8e1;border:1.5px solid #f59e0b;border-radius:12px;'
                    f'padding:14px 20px;margin-bottom:12px;">'
                    f'<div style="font-size:13px;font-weight:700;color:#92400e;margin-bottom:6px;">'
                    f'&#9888;&nbsp; Outlook Not Available</div>'
                    f'<div style="font-size:13px;color:#78350f;line-height:1.7;">'
                    f'{_email_err}<br><br>'
                    f'Make sure <strong>Microsoft Outlook is open</strong> and '
                    f'<strong>pywin32</strong> is installed (<code>pip install pywin32</code>).'
                    f'</div></div>',
                    unsafe_allow_html=True,
                )
        else:
            st.markdown(
                f'<div style="background:#fffde7;border:1.5px solid #f59e0b;border-radius:14px;'
                f'padding:16px 22px;margin-bottom:12px;">'
                f'<div style="font-size:14px;font-weight:700;color:#92400e;margin-bottom:8px;">'
                f'Send Reminder Emails to Non-Submitting Locations</div>'
                f'<div style="font-size:13px;color:#78350f;">'
                f'{len(pending_zones)} zone(s) with pending locations will receive emails via '
                f'Outlook ({_emails.SENDER_EMAIL}).'
                f'</div></div>',
                unsafe_allow_html=True,
            )

            email_key   = f"_rpt_email_open_{month_year}"
            confirm_key = f"_rpt_email_confirm_{month_year}"
            if not st.session_state.get(email_key):
                if st.button("Show Email Options", key="rpt_email_toggle", use_container_width=False):
                    st.session_state[email_key] = True
                    st.session_state[confirm_key] = False
                    st.rerun()
            else:
                # ── Zones to receive emails ───────────────────────────────────
                st.warning(
                    "You are about to send reminder emails to zone teams. "
                    "Verify recipients, preview the email, then confirm before sending."
                )
                if pending_zones:
                    st.markdown(f"**{len(pending_zones)} zone(s) with pending locations:**")
                else:
                    st.info("All zones have submitted. No emails to send.")

                # Recipients table
                with st.expander("Show Email Recipients per Zone", expanded=False):
                    rec_rows = []
                    for z in sorted(pending_zones):
                        rcp = _emails.get_zone_recipients(z)
                        rec_rows.append({
                            "Zone": z,
                            "To (Zone Head)": rcp["to"] or "—",
                            "CC": rcp["cc"] or "—",
                            "BCC (HQO)": rcp["bcc"],
                        })
                    if rec_rows:
                        st.markdown(
                            _rpt_table(rec_rows, center_cols=set()),
                            unsafe_allow_html=True,
                        )

                # ── Test Mode ─────────────────────────────────────────────────
                st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
                test_mode = st.toggle(
                    f"Test Mode — send only to {_emails.SENDER_EMAIL} (no actual zone emails)",
                    key=f"rpt_test_mode_{month_year}",
                    value=False,
                )
                if test_mode:
                    st.info(
                        f"**Test Mode ON** — One sample email (first pending zone) will be sent "
                        f"to **{_emails.SENDER_EMAIL}** only. No zone teams will receive anything."
                    )

                # ── Edit email content ─────────────────────────────────────────
                _intro_key = f"rpt_custom_intro_{month_year}"
                # Read current value from session state (persists whether expander is open or not)
                custom_intro = st.session_state.get(_intro_key, "")

                with st.expander("Edit Email Content (optional)", expanded=False):
                    custom_intro = st.text_area(
                        "Opening message in email body",
                        key=_intro_key,
                        height=90,
                        placeholder=(
                            "Leave blank to use the default message:\n"
                            "\"This is a reminder that the MIS submission for [Month] "
                            "is pending for the following locations in your zone. "
                            "Please ensure submissions are completed at the earliest.\""
                        ),
                    )

                    # Email preview
                    if st.button("Preview Email", key="rpt_preview_btn"):
                        sample_zone = pending_zones[0] if pending_zones else ""
                        if sample_zone:
                            sample_locs = [
                                r for r in all_rows
                                if r.get("zone") == sample_zone
                                and r.get("status") != "SUBMITTED"
                            ]
                            preview_html = _emails.build_preview_html(
                                sample_zone, month_year, sample_locs,
                                due_date, custom_intro,
                            )
                            st.caption(f"Preview: email for **{sample_zone}**")
                            import streamlit.components.v1 as components
                            components.html(preview_html, height=580, scrolling=True)
                        else:
                            st.info("No pending zones — nothing to preview.")

                # ── Confirmation checkbox ─────────────────────────────────────
                label = (
                    f"I confirm — send **TEST** email to {_emails.SENDER_EMAIL} for {month_year}."
                    if test_mode else
                    f"I confirm I have reviewed the list above and want to send reminder emails "
                    f"for **{month_year}** via Microsoft Outlook."
                )
                confirmed = st.checkbox(label, key=confirm_key)

                em_col1, em_col2 = st.columns([2, 1])
                with em_col1:
                    btn_label  = "Send Test Email" if test_mode else "Send Reminder Emails"
                    send_clicked = pending_zones and st.button(
                        btn_label, key="rpt_send_emails",
                        type="primary", use_container_width=True,
                    )
                    if send_clicked:
                        if not confirmed:
                            st.warning("Please tick the confirmation checkbox before sending.")
                        else:
                            spinner_msg = (
                                "Sending test email via Outlook…"
                                if test_mode else
                                "Sending emails via Outlook…"
                            )
                            with st.spinner(spinner_msg):
                                result = _emails.send_all_reminders(
                                    month_year, all_rows, due_date,
                                    custom_intro=custom_intro,
                                    test_mode=test_mode,
                                    test_email=_emails.SENDER_EMAIL,
                                )
                            if result["ok"]:
                                st.success(result["msg"])
                            else:
                                st.error(result["msg"])
                            if not test_mode:
                                st.session_state.pop(email_key, None)
                                st.session_state.pop(confirm_key, None)
                with em_col2:
                    if st.button("Cancel", key="rpt_email_cancel", use_container_width=True):
                        st.session_state.pop(email_key, None)
                        st.session_state.pop(confirm_key, None)
                        st.rerun()

    st.markdown("""
    <div style="margin-top:24px;padding:10px 4px;border-top:1px solid #dde3ed;
                display:flex;justify-content:space-between;font-size:11px;color:#aaa;">
      <span>&#169; 2026 Hindustan Petroleum Corporation Limited.</span>
      <span>HPCL SOD &nbsp;·&nbsp; MIS Reports</span>
    </div>""", unsafe_allow_html=True)


# ── AI Assistant page ─────────────────────────────────────────────────────────

def show_chatbot_page(user: dict):
    """Full-page MIS AI Assistant powered by Gemini Flash."""
    _dashboard_css()

    # Sidebar — reuse zone sidebar for Zone/Admin, minimal for Maker/Checker
    if user.get("role") in ("Zone", "Admin"):
        _zone_sidebar(user, "AI ASSISTANT", "MIS Query Interface")
        if st.sidebar.button("← Back to Dashboard", key="chat_back_zone",
                             use_container_width=True):
            st.session_state.selected_section = None
            st.rerun()
    else:
        with st.sidebar:
            sl = _assets().get("side_logo")
            if sl:
                st.markdown(
                    f'<div style="margin:0;padding:0;line-height:0;width:100%;overflow:hidden;">'
                    f'<img src="{sl}" style="width:100%;height:auto;display:block;'
                    f'margin:0;padding:0;"></div>',
                    unsafe_allow_html=True,
                )
            st.markdown(
                '<div style="padding:10px 18px 12px;border-bottom:2px solid #c62828;">'
                '<div style="color:#ff4d4d;font-size:11px;font-weight:700;letter-spacing:1.5px;">'
                'AI ASSISTANT</div>'
                '<div style="color:#ff9999;font-size:10px;margin-top:3px;">MIS Query Interface</div>'
                '</div><div style="height:8px;"></div>',
                unsafe_allow_html=True,
            )
            if st.button("← Back to Dashboard", key="chat_back_maker",
                         use_container_width=True):
                st.session_state.selected_section = None
                st.rerun()

    _dash_header(user)

    st.markdown(
        f'<div style="padding:12px 0 4px;">'
        f'<span style="font-size:20px;font-weight:800;color:{HPCL_BLUE};">🤖 MIS AI Assistant</span>'
        f'<span style="font-size:12px;color:#888;margin-left:10px;">Powered by Google Gemini Flash</span>'
        f'</div>'
        f'<div style="font-size:13px;color:#555;margin-bottom:8px;">'
        f'Ask about throughput, OPEX, compliance, stock loss, rankings — any of the 135 MIS parameters.'
        f'</div>',
        unsafe_allow_html=True,
    )

    st.session_state.setdefault("chat_history", [])
    history = st.session_state.chat_history

    # Clear chat button
    col_clear, _ = st.columns([1, 5])
    with col_clear:
        if st.button("🗑 Clear Chat", key="clear_chat"):
            st.session_state.chat_history = []
            st.rerun()

    st.markdown("<hr style='margin:6px 0 10px;border-color:#eee;'>", unsafe_allow_html=True)

    # Welcome message when empty
    if not history:
        with st.chat_message("assistant", avatar="🤖"):
            st.markdown(
                f"Hello **{user.get('locName', user.get('zone', 'there'))}**! "
                f"I'm your MIS Assistant. I can help you with:\n\n"
                f"- 📦 **Throughput & targets** — MS, HSD, total volumes\n"
                f"- 💰 **Financial** — OPEX ₹/MT, MEB vs budget, electricity\n"
                f"- 📋 **Compliance** — deadline status, missed months\n"
                f"- 🛢️ **Stock loss** — product-wise, month-on-month\n"
                f"- 🦺 **Safety** — HSE Index, PM %, M&I Index\n"
                f"- 🏆 **Rankings** — top/bottom performers in zone or all-India\n\n"
                f"*Ask me anything about MIS data!*"
            )

    # Display chat history
    for msg in history:
        avatar = "🤖" if msg["role"] == "assistant" else "👤"
        with st.chat_message(msg["role"], avatar=avatar):
            st.markdown(msg["content"])

    # Chat input
    if prompt := st.chat_input("Ask about any location, parameter, or ranking…"):
        history.append({"role": "user", "content": prompt})
        with st.chat_message("user", avatar="👤"):
            st.markdown(prompt)

        with st.chat_message("assistant", avatar="🤖"):
            with st.spinner("Thinking…"):
                answer = _gemini_response(prompt, user, history[:-1])
            st.markdown(answer)

        history.append({"role": "assistant", "content": answer})
        st.session_state.chat_history = history


# ── M&I MIS Page ─────────────────────────────────────────────────────────────

def _mi_parse_date(s: str):
    """'DD/MM/YYYY' string → datetime.date or None."""
    from datetime import datetime as _dtp
    if not s or str(s).upper().strip() in ("NA", "", "N/A", "NONE"):
        return None
    try:
        return _dtp.strptime(str(s).strip(), "%d/%m/%Y").date()
    except Exception:
        return None


def _mi_fmt_date(d) -> str:
    """date/datetime → 'DD/MM/YYYY' string, or '' if None."""
    if d is None:
        return ""
    if hasattr(d, "strftime"):
        return d.strftime("%d/%m/%Y")
    return ""


_MI_NA_STYLE = (
    'background:#fff8e1;border:1px solid #f59e0b;border-radius:8px;'
    'padding:10px 16px;font-size:13px;color:#92400e;margin-bottom:8px;'
)
_MI_ROW_DIV = (
    '<div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;'
    'padding:14px 16px 8px;margin-bottom:10px;">'
)


def _mi_summary_table(rows: list, columns: list):
    """Render a compact blue-header summary table of collected S5A row data."""
    if not rows:
        return
    import pandas as pd
    df = pd.DataFrame(rows, columns=columns)
    html = df.to_html(index=False, escape=True, border=0, classes="mis-tbl")
    st.markdown(
        '<div style="font-size:12px;font-weight:700;color:#001a6e;'
        'margin:10px 0 4px;letter-spacing:0.3px;">&#128202; Summary — rows entered so far</div>'
        + html,
        unsafe_allow_html=True,
    )


def _na_checkbox(label: str, key: str) -> bool:
    """Render an amber-highlighted banner + checkbox for 'Not Applicable' M&I sections."""
    st.markdown(
        '<div style="background:#fffbea;border-left:4px solid #f59e0b;border-radius:6px;'
        'padding:9px 14px;margin:10px 0 2px 0;font-size:12px;color:#78350f;line-height:1.5;">'
        '<b>Not applicable to your location?</b> Check the box below to mark this section '
        'as complete. Without either saved data <em>or</em> this checkbox ticked, the '
        'section stays incomplete (❌) and MIS submission will be blocked.</div>',
        unsafe_allow_html=True,
    )
    return st.checkbox(label, key=key)


def _mi_tab_outage(uid: str, month_year: str, tank_opts: list):
    T         = "to"
    sk_na     = f"mi_{T}_{uid}_{month_year}_na"
    sk_rows   = f"mi_{T}_{uid}_{month_year}_rows"
    sk_ctr    = f"mi_{T}_{uid}_{month_year}_ctr"
    sk_loaded = f"mi_{T}_{uid}_{month_year}_loaded"

    OUTAGE_REASONS = [
        "Tank Cleaning", "Hydrotest", "Cathodic Protection",
        "Inspection / Audit", "Planned Maintenance", "Repairs", "Other",
    ]
    STATUS_OPTS = [
        "Under Outage", "Cleaning in Progress", "Repairs in Progress",
        "Pending Commissioning", "Commissioned", "Delayed", "Extended",
    ]

    if not st.session_state.get(sk_loaded):
        saved = sheets.load_mi_data("MI_TANK_OUTAGE", uid, month_year)
        if saved and saved[0].get("na_flag") == "Y":
            st.session_state[sk_na]   = True
            st.session_state[sk_rows] = []
            st.session_state[sk_ctr]  = 0
        elif saved:
            ids = list(range(len(saved)))
            st.session_state[sk_na]   = False
            st.session_state[sk_rows] = ids
            st.session_state[sk_ctr]  = len(saved)
            for rid, row in zip(ids, saved):
                pfx = f"mi_{T}_{uid}_{month_year}_{rid}"
                tn  = row.get("tank_no", "")
                # Use None when tank not found so selectbox shows no implicit default
                st.session_state[f"{pfx}_tank"]    = tn if tn in tank_opts else ("Other Tanks" if "Other Tanks" in tank_opts else None)
                st.session_state[f"{pfx}_other"]   = row.get("other_tank_desc", "")
                st.session_state[f"{pfx}_p_start"] = _mi_parse_date(row.get("planned_start"))
                st.session_state[f"{pfx}_p_end"]   = _mi_parse_date(row.get("planned_end"))
                st.session_state[f"{pfx}_a_start"] = _mi_parse_date(row.get("actual_start"))
                st.session_state[f"{pfx}_a_end"]   = _mi_parse_date(row.get("actual_end"))
                reason_raw = row.get("outage_for", "")
                st.session_state[f"{pfx}_reason"]  = reason_raw if reason_raw in OUTAGE_REASONS else "Other"
                st.session_state[f"{pfx}_reason_o"] = reason_raw if reason_raw not in OUTAGE_REASONS else ""
                status_raw = row.get("current_status", STATUS_OPTS[0])
                st.session_state[f"{pfx}_status"]  = status_raw if status_raw in STATUS_OPTS else STATUS_OPTS[0]
        else:
            st.session_state[sk_na]   = False
            st.session_state[sk_rows] = [0]  # start with one blank row
            st.session_state[sk_ctr]  = 1
        st.session_state[sk_loaded] = True

    na_val = _na_checkbox("Not Applicable — No tank under outage this month", sk_na)

    if not na_val:
        row_ids   = st.session_state.get(sk_rows, [])
        to_delete = None

        for i, rid in enumerate(row_ids):
            pfx = f"mi_{T}_{uid}_{month_year}_{rid}"
            st.markdown(_MI_ROW_DIV, unsafe_allow_html=True)
            hc, dc = st.columns([11, 1])
            with hc:
                st.markdown(f"**Row {i + 1}**")
            with dc:
                if st.button("🗑", key=f"{pfx}_del", help="Remove row"):
                    to_delete = rid

            c1, c2 = st.columns(2)
            with c1:
                _tn_cur = st.session_state.get(f"{pfx}_tank")
                if _tn_cur is not None and _tn_cur in tank_opts:
                    _tn_idx = tank_opts.index(_tn_cur)
                else:
                    _tn_idx = None  # no implicit default for new/unrecognised rows
                st.selectbox("Tank No. *", tank_opts, index=_tn_idx, key=f"{pfx}_tank",
                             help="Select the SAP tank number from Tank Master list")
            with c2:
                if st.session_state.get(f"{pfx}_tank") == "Other Tanks":
                    st.text_input("Other Tank Description *", key=f"{pfx}_other", max_chars=256,
                                  help="Describe the tank if not in the standard Tank Master list")

            c3, c4, c5, c6 = st.columns(4)
            with c3:
                st.date_input("Planned Start *", key=f"{pfx}_p_start",
                              value=st.session_state.get(f"{pfx}_p_start"), format="DD/MM/YYYY",
                              help="Planned start date of the tank outage")
            with c4:
                st.date_input("Planned End *", key=f"{pfx}_p_end",
                              value=st.session_state.get(f"{pfx}_p_end"), format="DD/MM/YYYY",
                              help="Planned end/completion date of the tank outage")
            with c5:
                st.date_input("Actual Start *", key=f"{pfx}_a_start",
                              value=st.session_state.get(f"{pfx}_a_start"), format="DD/MM/YYYY",
                              help="Actual date the tank was taken under outage")
            with c6:
                st.date_input("Actual End (blank if ongoing)", key=f"{pfx}_a_end",
                              value=st.session_state.get(f"{pfx}_a_end"), format="DD/MM/YYYY",
                              help="Actual end date; leave blank if outage is still in progress")

            c7, c8, c9 = st.columns([2, 2, 2])
            with c7:
                r_def = st.session_state.get(f"{pfx}_reason", OUTAGE_REASONS[0])
                r_idx = OUTAGE_REASONS.index(r_def) if r_def in OUTAGE_REASONS else len(OUTAGE_REASONS) - 1
                st.selectbox("Outage For *", OUTAGE_REASONS, index=r_idx, key=f"{pfx}_reason",
                             help="Primary reason this tank has been taken under outage")
            with c8:
                if st.session_state.get(f"{pfx}_reason") == "Other":
                    st.text_input("Specify *", key=f"{pfx}_reason_o", max_chars=256,
                                  help="Describe the outage reason when 'Other' is selected")
            with c9:
                s_def = st.session_state.get(f"{pfx}_status", STATUS_OPTS[0])
                s_idx = STATUS_OPTS.index(s_def) if s_def in STATUS_OPTS else 0
                st.selectbox("Current Status *", STATUS_OPTS, index=s_idx, key=f"{pfx}_status",
                             help="Current operational status of this tank outage")

            st.markdown("</div>", unsafe_allow_html=True)

        if to_delete is not None:
            st.session_state[sk_rows] = [r for r in st.session_state[sk_rows] if r != to_delete]
            st.rerun()

        # ── Summary preview table ──
        _tbl = []
        for rid in st.session_state.get(sk_rows, []):
            pfx = f"mi_{T}_{uid}_{month_year}_{rid}"
            _tbl.append({
                "Tank": st.session_state.get(f"{pfx}_tank",""),
                "Plan Start": _mi_fmt_date(st.session_state.get(f"{pfx}_p_start")),
                "Plan End":   _mi_fmt_date(st.session_state.get(f"{pfx}_p_end")),
                "Actual Start": _mi_fmt_date(st.session_state.get(f"{pfx}_a_start")),
                "Actual End":   _mi_fmt_date(st.session_state.get(f"{pfx}_a_end")),
                "Outage For": st.session_state.get(f"{pfx}_reason",""),
                "Status":     st.session_state.get(f"{pfx}_status",""),
            })
        _mi_summary_table(_tbl, ["Tank","Plan Start","Plan End","Actual Start","Actual End","Outage For","Status"])

        if st.button("➕ Add Row", key=f"mi_{T}_add_{uid}_{month_year}"):
            new_id = st.session_state.get(sk_ctr, 0)
            st.session_state.setdefault(sk_rows, []).append(new_id)
            st.session_state[sk_ctr] = new_id + 1
            st.rerun()

    if st.button("💾 Save Tank Outage Data", key=f"mi_{T}_save_{uid}_{month_year}", type="primary"):
        if na_val:
            rows_to_save = [{"row_no": "1", "na_flag": "Y", "tank_no": "NA",
                             "other_tank_desc": "NA", "planned_start": "NA",
                             "planned_end": "NA", "actual_start": "NA",
                             "actual_end": "NA", "outage_for": "NA", "current_status": "NA"}]
        else:
            rows_to_save = []
            errors = []
            for i, rid in enumerate(st.session_state.get(sk_rows, [])):
                pfx  = f"mi_{T}_{uid}_{month_year}_{rid}"
                tank = st.session_state.get(f"{pfx}_tank", "")
                other_desc = st.session_state.get(f"{pfx}_other", "")
                if tank == "Other Tanks" and not other_desc.strip():
                    errors.append(f"Row {i+1}: Other Tank Description is required.")
                p_s = _mi_fmt_date(st.session_state.get(f"{pfx}_p_start"))
                p_e = _mi_fmt_date(st.session_state.get(f"{pfx}_p_end"))
                a_s = _mi_fmt_date(st.session_state.get(f"{pfx}_a_start"))
                a_e = _mi_fmt_date(st.session_state.get(f"{pfx}_a_end"))
                reason = st.session_state.get(f"{pfx}_reason", "")
                if reason == "Other":
                    reason = st.session_state.get(f"{pfx}_reason_o", "").strip()
                    if not reason:
                        errors.append(f"Row {i+1}: Specify the outage reason.")
                status = st.session_state.get(f"{pfx}_status", "")
                if not all([tank, p_s, p_e, a_s, reason, status]):
                    errors.append(f"Row {i+1}: All required fields must be filled.")
                else:
                    rows_to_save.append({
                        "row_no": str(i + 1), "na_flag": "N",
                        "tank_no": tank, "other_tank_desc": other_desc,
                        "planned_start": p_s, "planned_end": p_e,
                        "actual_start": a_s, "actual_end": a_e,
                        "outage_for": reason, "current_status": status,
                    })
            if errors:
                for e in errors:
                    st.error(e)
                return
            if not rows_to_save:
                st.warning("Add at least one row, or mark as Not Applicable.")
                return
        res = sheets.save_mi_data("MI_TANK_OUTAGE", uid, month_year, rows_to_save)
        if res.get("ok"):
            st.success(f"✅ Saved {res['rows']} row(s).")
            st.session_state.pop(sk_loaded, None)
        else:
            st.error(f"Save failed: {res.get('msg', '')}")


def _mi_tab_repair(uid: str, month_year: str, tank_opts: list):
    T         = "mr"
    sk_na     = f"mi_{T}_{uid}_{month_year}_na"
    sk_rows   = f"mi_{T}_{uid}_{month_year}_rows"
    sk_ctr    = f"mi_{T}_{uid}_{month_year}_ctr"
    sk_loaded = f"mi_{T}_{uid}_{month_year}_loaded"

    STATUS_OPTS = ["In Progress", "Completed", "Delayed", "On Hold", "Cancelled"]

    if not st.session_state.get(sk_loaded):
        saved = sheets.load_mi_data("MI_MAJOR_REPAIR", uid, month_year)
        if saved and saved[0].get("na_flag") == "Y":
            st.session_state[sk_na]   = True
            st.session_state[sk_rows] = []
            st.session_state[sk_ctr]  = 0
        elif saved:
            ids = list(range(len(saved)))
            st.session_state[sk_na]   = False
            st.session_state[sk_rows] = ids
            st.session_state[sk_ctr]  = len(saved)
            for rid, row in zip(ids, saved):
                pfx = f"mi_{T}_{uid}_{month_year}_{rid}"
                tn  = row.get("tank_no", "")
                st.session_state[f"{pfx}_tank"]   = tn if tn in tank_opts else ("Other Tanks" if "Other Tanks" in tank_opts else None)
                st.session_state[f"{pfx}_other"]  = row.get("other_tank_desc", "")
                st.session_state[f"{pfx}_nature"] = row.get("nature_of_repair", "")
                rc_raw = row.get("revenue_capex", "Revenue")
                st.session_state[f"{pfx}_rc"]     = rc_raw if rc_raw in ("Revenue", "Capex") else "Revenue"
                st.session_state[f"{pfx}_ar"]     = row.get("ar_code", "")
                st.session_state[f"{pfx}_status"] = row.get("current_status", STATUS_OPTS[0])
                st.session_state[f"{pfx}_etc"]    = _mi_parse_date(row.get("etc_date"))
        else:
            st.session_state[sk_na]   = False
            st.session_state[sk_rows] = [0]
            st.session_state[sk_ctr]  = 1
        st.session_state[sk_loaded] = True

    na_val = _na_checkbox("Not Applicable — No major repair this month", sk_na)

    if not na_val:
        row_ids   = st.session_state.get(sk_rows, [])
        to_delete = None

        for i, rid in enumerate(row_ids):
            pfx = f"mi_{T}_{uid}_{month_year}_{rid}"
            st.markdown(_MI_ROW_DIV, unsafe_allow_html=True)
            hc, dc = st.columns([11, 1])
            with hc:
                st.markdown(f"**Row {i + 1}**")
            with dc:
                if st.button("🗑", key=f"{pfx}_del", help="Remove row"):
                    to_delete = rid

            c1, c2 = st.columns(2)
            with c1:
                _tn_cur = st.session_state.get(f"{pfx}_tank")
                if _tn_cur is not None and _tn_cur in tank_opts:
                    _tn_idx = tank_opts.index(_tn_cur)
                else:
                    _tn_idx = None
                st.selectbox("Tank No. *", tank_opts, index=_tn_idx, key=f"{pfx}_tank",
                             help="Select the SAP tank number from Tank Master list")
            with c2:
                if st.session_state.get(f"{pfx}_tank") == "Other Tanks":
                    st.text_input("Other Tank Description *", key=f"{pfx}_other", max_chars=256,
                                  help="Describe the tank if not in the standard Tank Master list")

            st.text_area("Nature of Repair *", key=f"{pfx}_nature", max_chars=256,
                         help="Describe the repair work in detail (max 256 characters)")

            c3, c4, c5, c6 = st.columns(4)
            with c3:
                rc_def = st.session_state.get(f"{pfx}_rc", "Revenue")
                rc_idx = 0 if rc_def == "Revenue" else 1
                st.selectbox("Revenue / Capex *", ["Revenue", "Capex"], index=rc_idx,
                             key=f"{pfx}_rc", help="Select Revenue for opex or Capex for capital expenditure")
            with c4:
                if st.session_state.get(f"{pfx}_rc") == "Capex":
                    st.text_input("AR Code *", key=f"{pfx}_ar", max_chars=50,
                                  help="SAP AR (Activity Request) code for this Capex work")
            with c5:
                s_def = st.session_state.get(f"{pfx}_status", STATUS_OPTS[0])
                s_idx = STATUS_OPTS.index(s_def) if s_def in STATUS_OPTS else 0
                st.selectbox("Current Status *", STATUS_OPTS, index=s_idx, key=f"{pfx}_status",
                             help="Current progress status of the repair work")
            with c6:
                st.date_input("ETC Date *", key=f"{pfx}_etc",
                              value=st.session_state.get(f"{pfx}_etc"), format="DD/MM/YYYY",
                              help="Estimated Time to Complete — expected completion date")

            st.markdown("</div>", unsafe_allow_html=True)

        if to_delete is not None:
            st.session_state[sk_rows] = [r for r in st.session_state[sk_rows] if r != to_delete]
            st.rerun()

        # Summary preview
        _tbl = []
        for rid in st.session_state.get(sk_rows, []):
            pfx = f"mi_{T}_{uid}_{month_year}_{rid}"
            _tbl.append({
                "Tank": st.session_state.get(f"{pfx}_tank",""),
                "Nature of Repair": st.session_state.get(f"{pfx}_nature",""),
                "Rev/Capex": st.session_state.get(f"{pfx}_rc",""),
                "ETC Date": _mi_fmt_date(st.session_state.get(f"{pfx}_etc")),
                "Status": st.session_state.get(f"{pfx}_status",""),
            })
        _mi_summary_table(_tbl, ["Tank","Nature of Repair","Rev/Capex","ETC Date","Status"])

        if st.button("➕ Add Row", key=f"mi_{T}_add_{uid}_{month_year}"):
            new_id = st.session_state.get(sk_ctr, 0)
            st.session_state.setdefault(sk_rows, []).append(new_id)
            st.session_state[sk_ctr] = new_id + 1
            st.rerun()

    if st.button("💾 Save Major Repair Data", key=f"mi_{T}_save_{uid}_{month_year}", type="primary"):
        if na_val:
            rows_to_save = [{"row_no": "1", "na_flag": "Y", "tank_no": "NA",
                             "other_tank_desc": "NA", "nature_of_repair": "NA",
                             "revenue_capex": "NA", "ar_code": "NA",
                             "current_status": "NA", "etc_date": "NA"}]
        else:
            rows_to_save = []
            errors = []
            for i, rid in enumerate(st.session_state.get(sk_rows, [])):
                pfx    = f"mi_{T}_{uid}_{month_year}_{rid}"
                tank   = st.session_state.get(f"{pfx}_tank", "")
                other  = st.session_state.get(f"{pfx}_other", "")
                nature = (st.session_state.get(f"{pfx}_nature") or "").strip()
                rc     = st.session_state.get(f"{pfx}_rc", "Revenue")
                ar     = (st.session_state.get(f"{pfx}_ar") or "").strip()
                status = st.session_state.get(f"{pfx}_status", "")
                etc_d  = _mi_fmt_date(st.session_state.get(f"{pfx}_etc"))
                if tank == "Other Tanks" and not other.strip():
                    errors.append(f"Row {i+1}: Other Tank Description required.")
                if rc == "Capex" and not ar:
                    errors.append(f"Row {i+1}: AR Code is required for Capex.")
                if not all([tank, nature, rc, status, etc_d]):
                    errors.append(f"Row {i+1}: All required fields must be filled.")
                else:
                    rows_to_save.append({
                        "row_no": str(i + 1), "na_flag": "N",
                        "tank_no": tank, "other_tank_desc": other,
                        "nature_of_repair": nature, "revenue_capex": rc,
                        "ar_code": ar, "current_status": status, "etc_date": etc_d,
                    })
            if errors:
                for e in errors:
                    st.error(e)
                return
            if not rows_to_save:
                st.warning("Add at least one row, or mark as Not Applicable.")
                return
        res = sheets.save_mi_data("MI_MAJOR_REPAIR", uid, month_year, rows_to_save)
        if res.get("ok"):
            st.success(f"✅ Saved {res['rows']} row(s).")
            st.session_state.pop(sk_loaded, None)
        else:
            st.error(f"Save failed: {res.get('msg', '')}")


def _mi_tab_vru(uid: str, month_year: str):
    T         = "vr"
    sk_na     = f"mi_{T}_{uid}_{month_year}_na"
    sk_op     = f"mi_{T}_{uid}_{month_year}_op"
    sk_loaded = f"mi_{T}_{uid}_{month_year}_loaded"

    if not st.session_state.get(sk_loaded):
        saved = sheets.load_mi_data("MI_VRU", uid, month_year)
        if saved and saved[0].get("na_flag") == "Y":
            st.session_state[sk_na] = True
        elif saved:
            row = saved[0]
            st.session_state[sk_na] = False
            op_raw = row.get("vru_operational", "Yes")
            st.session_state[sk_op] = op_raw if op_raw in ("Yes", "No") else "Yes"
            for fld in ("date_not_operating", "action_taken", "etc_date",
                        "ms_vol_recovered_kl", "inlet_mfm_start_m3", "inlet_mfm_end_m3",
                        "outlet_mfm_start_m3", "outlet_mfm_end_m3", "vapour_treated_m3",
                        "voc_value_mgcc", "inlet_emission_mgcc",
                        "ms_gasohol_tt_vol_kl", "hsd_tt_vol_kl",
                        "ms_gasohol_tw_vol_kl", "hsd_tw_vol_kl", "vru_uptime_pct"):
                sk_f = f"mi_{T}_{uid}_{month_year}_{fld}"
                if fld in ("date_not_operating", "etc_date"):
                    st.session_state[sk_f] = _mi_parse_date(row.get(fld))
                else:
                    st.session_state[sk_f] = row.get(fld, "")
        else:
            st.session_state[sk_na] = False
        st.session_state[sk_loaded] = True

    na_val = _na_checkbox("Not Applicable — VRU not installed at this location", sk_na)

    if not na_val:
        op_def = st.session_state.get(sk_op, "Yes")
        op_idx = 0 if op_def == "Yes" else 1
        op_val = st.selectbox("VRU Operational this month? *", ["Yes", "No"], index=op_idx,
                              key=sk_op, help="Select Yes if VRU was operational during this month")

        pfx = f"mi_{T}_{uid}_{month_year}"
        if op_val == "No":
            st.markdown("##### Non-Operational Details")
            c1, c2, c3 = st.columns(3)
            with c1:
                st.date_input("Date Since Not Operating *",
                              key=f"{pfx}_date_not_operating",
                              value=st.session_state.get(f"{pfx}_date_not_operating"),
                              format="DD/MM/YYYY",
                              help="Date from which VRU has been non-operational")
            with c2:
                st.text_area("Action Taken *", key=f"{pfx}_action_taken", max_chars=256,
                             help="Describe corrective actions taken to restore VRU operation")
            with c3:
                st.date_input("ETC Date *",
                              key=f"{pfx}_etc_date",
                              value=st.session_state.get(f"{pfx}_etc_date"),
                              format="DD/MM/YYYY",
                              help="Estimated date by which VRU will be restored to operation")
        else:
            st.markdown("##### Operational Readings")
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.text_input("MS Vol Recovered (KL) *", key=f"{pfx}_ms_vol_recovered_kl",
                              help="Volume of MS/Gasohol vapour recovered in kilolitres")
            with c2:
                st.text_input("Inlet MFM Start (m³) *", key=f"{pfx}_inlet_mfm_start_m3",
                              help="Inlet mass flow meter reading at start of month (m³)")
            with c3:
                st.text_input("Inlet MFM End (m³) *", key=f"{pfx}_inlet_mfm_end_m3",
                              help="Inlet mass flow meter reading at end of month (m³)")
            with c4:
                st.text_input("Outlet MFM Start (m³) *", key=f"{pfx}_outlet_mfm_start_m3",
                              help="Outlet mass flow meter reading at start of month (m³)")

            c5, c6, c7, c8 = st.columns(4)
            with c5:
                st.text_input("Outlet MFM End (m³) *", key=f"{pfx}_outlet_mfm_end_m3",
                              help="Outlet mass flow meter reading at end of month (m³)")
            with c6:
                st.text_input("Vapour Treated (m³) *", key=f"{pfx}_vapour_treated_m3",
                              help="Total vapour treated by VRU during the month (m³)")
            with c7:
                st.text_input("VOC Value (mg/cc) *", key=f"{pfx}_voc_value_mgcc",
                              help="VOC concentration at VRU outlet (mg/cc)")
            with c8:
                st.text_input("Inlet Emission (mg/cc) *", key=f"{pfx}_inlet_emission_mgcc",
                              help="VOC concentration at VRU inlet (mg/cc)")

            st.markdown("##### TT / TW Volumes")
            c9, c10, c11, c12 = st.columns(4)
            with c9:
                st.text_input("MS/Gasohol TT Vol (KL)", key=f"{pfx}_ms_gasohol_tt_vol_kl",
                              help="MS/Gasohol volume loaded via tank truck during month (KL)")
            with c10:
                st.text_input("HSD TT Vol (KL)", key=f"{pfx}_hsd_tt_vol_kl",
                              help="HSD volume loaded via tank truck during month (KL)")
            with c11:
                st.text_input("MS/Gasohol TW Vol (KL)", key=f"{pfx}_ms_gasohol_tw_vol_kl",
                              help="MS/Gasohol volume loaded via tank wagon during month (KL)")
            with c12:
                st.text_input("HSD TW Vol (KL)", key=f"{pfx}_hsd_tw_vol_kl",
                              help="HSD volume loaded via tank wagon during month (KL)")

            st.text_input("VRU Uptime % *", key=f"{pfx}_vru_uptime_pct",
                          help="VRU uptime as percentage, e.g. 95.50")

    if st.button("💾 Save VRU Data", key=f"mi_{T}_save_{uid}_{month_year}", type="primary"):
        pfx = f"mi_{T}_{uid}_{month_year}"
        if na_val:
            rec = {"na_flag": "Y", "vru_operational": "NA", "date_not_operating": "NA",
                   "action_taken": "NA", "etc_date": "NA", "ms_vol_recovered_kl": "NA",
                   "inlet_mfm_start_m3": "NA", "inlet_mfm_end_m3": "NA",
                   "outlet_mfm_start_m3": "NA", "outlet_mfm_end_m3": "NA",
                   "vapour_treated_m3": "NA", "voc_value_mgcc": "NA",
                   "inlet_emission_mgcc": "NA", "ms_gasohol_tt_vol_kl": "NA",
                   "hsd_tt_vol_kl": "NA", "ms_gasohol_tw_vol_kl": "NA",
                   "hsd_tw_vol_kl": "NA", "vru_uptime_pct": "NA"}
        else:
            op_val = st.session_state.get(sk_op, "Yes")
            rec = {"na_flag": "N", "vru_operational": op_val}
            if op_val == "No":
                d_off = _mi_fmt_date(st.session_state.get(f"{pfx}_date_not_operating"))
                action = (st.session_state.get(f"{pfx}_action_taken") or "").strip()
                etc_d  = _mi_fmt_date(st.session_state.get(f"{pfx}_etc_date"))
                if not all([d_off, action, etc_d]):
                    st.error("All non-operational fields are required.")
                    return
                rec.update({"date_not_operating": d_off, "action_taken": action,
                             "etc_date": etc_d})
                for fld in ("ms_vol_recovered_kl", "inlet_mfm_start_m3", "inlet_mfm_end_m3",
                            "outlet_mfm_start_m3", "outlet_mfm_end_m3", "vapour_treated_m3",
                            "voc_value_mgcc", "inlet_emission_mgcc", "ms_gasohol_tt_vol_kl",
                            "hsd_tt_vol_kl", "ms_gasohol_tw_vol_kl", "hsd_tw_vol_kl",
                            "vru_uptime_pct"):
                    rec[fld] = "NA"
            else:
                required_op = ["ms_vol_recovered_kl", "inlet_mfm_start_m3", "inlet_mfm_end_m3",
                               "outlet_mfm_start_m3", "outlet_mfm_end_m3", "vapour_treated_m3",
                               "voc_value_mgcc", "inlet_emission_mgcc", "vru_uptime_pct"]
                errors = []
                for fld in required_op:
                    val = (st.session_state.get(f"{pfx}_{fld}") or "").strip()
                    if not val:
                        errors.append(fld.replace("_", " ").title())
                    rec[fld] = val
                if errors:
                    st.error("Required fields missing: " + ", ".join(errors))
                    return
                for fld in ("ms_gasohol_tt_vol_kl", "hsd_tt_vol_kl",
                            "ms_gasohol_tw_vol_kl", "hsd_tw_vol_kl"):
                    rec[fld] = (st.session_state.get(f"{pfx}_{fld}") or "").strip()
                rec["date_not_operating"] = ""
                rec["action_taken"] = ""
                rec["etc_date"] = ""
        res = sheets.save_mi_data("MI_VRU", uid, month_year, [rec])
        if res.get("ok"):
            st.success("✅ VRU data saved.")
            st.session_state.pop(sk_loaded, None)
        else:
            st.error(f"Save failed: {res.get('msg', '')}")


def _mi_tab_audit2526(uid: str, month_year: str):
    T         = "a25"
    sk_na     = f"mi_{T}_{uid}_{month_year}_na"
    sk_loaded = f"mi_{T}_{uid}_{month_year}_loaded"
    pfx       = f"mi_{T}_{uid}_{month_year}"

    if not st.session_state.get(sk_loaded):
        saved = sheets.load_mi_data("MI_AUDIT_2526", uid, month_year)
        if saved and saved[0].get("na_flag") == "Y":
            st.session_state[sk_na] = True
        elif saved:
            row = saved[0]
            st.session_state[sk_na] = False
            st.session_state[f"{pfx}_date"]     = _mi_parse_date(row.get("audit_date"))
            st.session_state[f"{pfx}_no_reco"]  = row.get("no_recommendations", "")
            st.session_state[f"{pfx}_no_pend"]  = row.get("no_pending", "")
            st.session_state[f"{pfx}_score"]    = row.get("external_score", "")
        else:
            st.session_state[sk_na] = False
        st.session_state[sk_loaded] = True

    st.caption("M&I Audit 2025-26 — Enter once; update pending count monthly.")
    na_val = _na_checkbox("Not Applicable — No M&I Audit 2025-26 at this location", sk_na)

    if not na_val:
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.date_input("Audit Date *", key=f"{pfx}_date",
                          value=st.session_state.get(f"{pfx}_date"), format="DD/MM/YYYY",
                          help="Date on which M&I Audit 2025-26 was conducted")
        with c2:
            st.text_input("No. of Recommendations *", key=f"{pfx}_no_reco",
                          help="Total number of recommendations raised in the audit (integer)")
        with c3:
            st.text_input("No. Pending *", key=f"{pfx}_no_pend",
                          help="Number of recommendations still pending closure (update monthly)")
        with c4:
            st.text_input("External Score (0–100) *", key=f"{pfx}_score",
                          help="Audit external score between 0 and 100 (up to 2 decimal places)")

    if st.button("💾 Save M&I Audit 25-26 Data", key=f"mi_{T}_save_{uid}_{month_year}", type="primary"):
        if na_val:
            rec = {"na_flag": "Y", "audit_date": "NA", "no_recommendations": "NA",
                   "no_pending": "NA", "external_score": "NA"}
        else:
            adt  = _mi_fmt_date(st.session_state.get(f"{pfx}_date"))
            reco = (st.session_state.get(f"{pfx}_no_reco") or "").strip()
            pend = (st.session_state.get(f"{pfx}_no_pend") or "").strip()
            scr  = (st.session_state.get(f"{pfx}_score") or "").strip()
            if not all([adt, reco, pend, scr]):
                st.error("All fields are required.")
                return
            try:
                float(scr)
                int(reco)
                int(pend)
            except ValueError:
                st.error("Recommendations and Pending must be integers; Score must be a number.")
                return
            rec = {"na_flag": "N", "audit_date": adt, "no_recommendations": reco,
                   "no_pending": pend, "external_score": scr}
        res = sheets.save_mi_data("MI_AUDIT_2526", uid, month_year, [rec])
        if res.get("ok"):
            st.success("✅ M&I Audit 25-26 data saved.")
            st.session_state.pop(sk_loaded, None)
        else:
            st.error(f"Save failed: {res.get('msg', '')}")


def _mi_tab_audit2627(uid: str, month_year: str):
    T         = "a27"
    sk_na     = f"mi_{T}_{uid}_{month_year}_na"
    sk_co     = f"mi_{T}_{uid}_{month_year}_co"
    sk_loaded = f"mi_{T}_{uid}_{month_year}_loaded"
    pfx       = f"mi_{T}_{uid}_{month_year}"

    if not st.session_state.get(sk_loaded):
        saved = sheets.load_mi_data("MI_AUDIT_2627", uid, month_year)
        if saved and saved[0].get("na_flag") == "Y":
            st.session_state[sk_na] = True
        elif saved:
            row = saved[0]
            st.session_state[sk_na] = False
            co_raw = row.get("audit_carried_out", "No")
            st.session_state[sk_co] = co_raw if co_raw in ("Yes", "No") else "No"
            st.session_state[f"{pfx}_date"]     = _mi_parse_date(row.get("audit_date"))
            st.session_state[f"{pfx}_no_reco"]  = row.get("no_recommendations", "")
            st.session_state[f"{pfx}_no_pend"]  = row.get("no_pending", "")
            st.session_state[f"{pfx}_score"]    = row.get("external_score", "")
        else:
            st.session_state[sk_na] = False
        st.session_state[sk_loaded] = True

    st.caption("M&I Audit 2026-27 — Audit may or may not have been carried out yet.")
    na_val = _na_checkbox("Not Applicable — No M&I Audit 2026-27 at this location", sk_na)

    if not na_val:
        co_def = st.session_state.get(sk_co, "No")
        co_idx = 0 if co_def == "Yes" else 1
        co_val = st.selectbox("Audit Carried Out? *", ["Yes", "No"], index=co_idx,
                              key=sk_co, help="Has M&I Audit 2026-27 been carried out yet?")

        if co_val == "Yes":
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.date_input("Audit Date *", key=f"{pfx}_date",
                              value=st.session_state.get(f"{pfx}_date"), format="DD/MM/YYYY",
                              help="Date on which M&I Audit 2026-27 was conducted")
            with c2:
                st.text_input("No. of Recommendations *", key=f"{pfx}_no_reco",
                              help="Total number of recommendations raised in the audit (integer)")
            with c3:
                st.text_input("No. Pending *", key=f"{pfx}_no_pend",
                              help="Number of recommendations still pending closure (update monthly)")
            with c4:
                st.text_input("External Score (0–100) *", key=f"{pfx}_score",
                              help="Audit external score between 0 and 100 (up to 2 decimal places)")
        else:
            st.info("Audit date and score will be recorded as NA.")

    if st.button("💾 Save M&I Audit 26-27 Data", key=f"mi_{T}_save_{uid}_{month_year}", type="primary"):
        if na_val:
            rec = {"na_flag": "Y", "audit_carried_out": "NA", "audit_date": "NA",
                   "no_recommendations": "NA", "no_pending": "NA", "external_score": "NA"}
        else:
            co_val = st.session_state.get(sk_co, "No")
            if co_val == "No":
                rec = {"na_flag": "N", "audit_carried_out": "No",
                       "audit_date": "NA", "no_recommendations": "NA",
                       "no_pending": "NA", "external_score": "NA"}
            else:
                adt  = _mi_fmt_date(st.session_state.get(f"{pfx}_date"))
                reco = (st.session_state.get(f"{pfx}_no_reco") or "").strip()
                pend = (st.session_state.get(f"{pfx}_no_pend") or "").strip()
                scr  = (st.session_state.get(f"{pfx}_score") or "").strip()
                if not all([adt, reco, pend, scr]):
                    st.error("All fields are required when audit is carried out.")
                    return
                try:
                    float(scr); int(reco); int(pend)
                except ValueError:
                    st.error("Recommendations and Pending must be integers; Score must be a number.")
                    return
                rec = {"na_flag": "N", "audit_carried_out": "Yes",
                       "audit_date": adt, "no_recommendations": reco,
                       "no_pending": pend, "external_score": scr}
        res = sheets.save_mi_data("MI_AUDIT_2627", uid, month_year, [rec])
        if res.get("ok"):
            st.success("✅ M&I Audit 26-27 data saved.")
            st.session_state.pop(sk_loaded, None)
        else:
            st.error(f"Save failed: {res.get('msg', '')}")


def _mi_tab_tech_audit(uid: str, month_year: str):
    T         = "ta"
    sk_na     = f"mi_{T}_{uid}_{month_year}_na"
    sk_rows   = f"mi_{T}_{uid}_{month_year}_rows"
    sk_ctr    = f"mi_{T}_{uid}_{month_year}_ctr"
    sk_loaded = f"mi_{T}_{uid}_{month_year}_loaded"

    if not st.session_state.get(sk_loaded):
        saved = sheets.load_mi_data("MI_TECH_AUDIT", uid, month_year)
        if saved and saved[0].get("na_flag") == "Y":
            st.session_state[sk_na]   = True
            st.session_state[sk_rows] = []
            st.session_state[sk_ctr]  = 0
        elif saved:
            ids = list(range(len(saved)))
            st.session_state[sk_na]   = False
            st.session_state[sk_rows] = ids
            st.session_state[sk_ctr]  = len(saved)
            for rid, row in zip(ids, saved):
                pfx = f"mi_{T}_{uid}_{month_year}_{rid}"
                st.session_state[f"{pfx}_date"]     = _mi_parse_date(row.get("audit_date"))
                st.session_state[f"{pfx}_no_reco"]  = row.get("no_recommendations", "")
                st.session_state[f"{pfx}_no_pend"]  = row.get("no_pending", "")
                st.session_state[f"{pfx}_ref"]      = row.get("ref_no", "")
        else:
            st.session_state[sk_na]   = False
            st.session_state[sk_rows] = [0]
            st.session_state[sk_ctr]  = 1
        st.session_state[sk_loaded] = True

    st.info("📋 UPDATE DETAILS FOR LAST AUDIT CONDUCTED AT LOCATION")

    na_val = _na_checkbox("Not Applicable — No technical audit this month", sk_na)

    if not na_val:
        row_ids   = st.session_state.get(sk_rows, [])
        to_delete = None

        for i, rid in enumerate(row_ids):
            pfx = f"mi_{T}_{uid}_{month_year}_{rid}"
            st.markdown(_MI_ROW_DIV, unsafe_allow_html=True)
            hc, dc = st.columns([11, 1])
            with hc:
                st.markdown(f"**Audit {i + 1}**")
            with dc:
                if st.button("🗑", key=f"{pfx}_del", help="Remove"):
                    to_delete = rid

            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.date_input("Audit Date *", key=f"{pfx}_date",
                              value=st.session_state.get(f"{pfx}_date"), format="DD/MM/YYYY",
                              help="Date on which the technical audit was conducted")
            with c2:
                st.text_input("No. of Recommendations *", key=f"{pfx}_no_reco",
                              help="Total number of recommendations raised in this audit (integer)")
            with c3:
                st.text_input("No. Pending *", key=f"{pfx}_no_pend",
                              help="Number of audit recommendations still pending closure")
            with c4:
                st.text_input("Audit Ref No. as per Audit Portal *", key=f"{pfx}_ref",
                              max_chars=256,
                              help="Audit reference/report number as recorded in the Audit Portal")
            st.markdown("</div>", unsafe_allow_html=True)

        if to_delete is not None:
            st.session_state[sk_rows] = [r for r in st.session_state[sk_rows] if r != to_delete]
            st.rerun()

        # Summary preview
        _tbl = []
        for rid in st.session_state.get(sk_rows, []):
            pfx = f"mi_{T}_{uid}_{month_year}_{rid}"
            _tbl.append({
                "Audit Date": _mi_fmt_date(st.session_state.get(f"{pfx}_date")),
                "Recommendations": st.session_state.get(f"{pfx}_no_reco", ""),
                "Pending": st.session_state.get(f"{pfx}_no_pend", ""),
                "Audit Ref No.": st.session_state.get(f"{pfx}_ref", ""),
            })
        _mi_summary_table(_tbl, ["Audit Date", "Recommendations", "Pending", "Audit Ref No."])

        if st.button("➕ Add Audit", key=f"mi_{T}_add_{uid}_{month_year}"):
            new_id = st.session_state.get(sk_ctr, 0)
            st.session_state.setdefault(sk_rows, []).append(new_id)
            st.session_state[sk_ctr] = new_id + 1
            st.rerun()

    if st.button("💾 Save Technical Audit Data", key=f"mi_{T}_save_{uid}_{month_year}", type="primary"):
        if na_val:
            rows_to_save = [{"row_no": "1", "na_flag": "Y", "audit_date": "NA",
                             "no_recommendations": "NA", "no_pending": "NA", "ref_no": "NA"}]
        else:
            rows_to_save = []
            errors = []
            for i, rid in enumerate(st.session_state.get(sk_rows, [])):
                pfx  = f"mi_{T}_{uid}_{month_year}_{rid}"
                adt  = _mi_fmt_date(st.session_state.get(f"{pfx}_date"))
                reco = (st.session_state.get(f"{pfx}_no_reco") or "").strip()
                pend = (st.session_state.get(f"{pfx}_no_pend") or "").strip()
                ref  = (st.session_state.get(f"{pfx}_ref") or "").strip()
                if not all([adt, reco, pend, ref]):
                    errors.append(f"Audit {i+1}: All fields are required.")
                else:
                    rows_to_save.append({"row_no": str(i + 1), "na_flag": "N",
                                         "audit_date": adt, "no_recommendations": reco,
                                         "no_pending": pend, "ref_no": ref})
            if errors:
                for e in errors:
                    st.error(e)
                return
            if not rows_to_save:
                st.warning("Add at least one audit record, or mark as Not Applicable.")
                return
        res = sheets.save_mi_data("MI_TECH_AUDIT", uid, month_year, rows_to_save)
        if res.get("ok"):
            st.success(f"✅ Saved {res['rows']} audit record(s).")
            st.session_state.pop(sk_loaded, None)
        else:
            st.error(f"Save failed: {res.get('msg', '')}")


def _mi_tab_equip(uid: str, month_year: str):
    T         = "eb"
    sk_na     = f"mi_{T}_{uid}_{month_year}_na"
    sk_rows   = f"mi_{T}_{uid}_{month_year}_rows"
    sk_ctr    = f"mi_{T}_{uid}_{month_year}_ctr"
    sk_loaded = f"mi_{T}_{uid}_{month_year}_loaded"

    EQUIP_OPTS = ["Pipeline", "Pump", "Fire Fighting Equipment",
                  "Fire Engine", "DG Set", "Other"]

    if not st.session_state.get(sk_loaded):
        saved = sheets.load_mi_data("MI_EQUIP_BREAKDOWN", uid, month_year)
        if saved and saved[0].get("na_flag") == "Y":
            st.session_state[sk_na]   = True
            st.session_state[sk_rows] = []
            st.session_state[sk_ctr]  = 0
        elif saved:
            ids = list(range(len(saved)))
            st.session_state[sk_na]   = False
            st.session_state[sk_rows] = ids
            st.session_state[sk_ctr]  = len(saved)
            for rid, row in zip(ids, saved):
                pfx = f"mi_{T}_{uid}_{month_year}_{rid}"
                en_raw = row.get("equipment_name", EQUIP_OPTS[0])
                st.session_state[f"{pfx}_eqname"]  = en_raw if en_raw in EQUIP_OPTS else "Other"
                st.session_state[f"{pfx}_eqother"] = row.get("equipment_name_other", "")
                st.session_state[f"{pfx}_details"] = row.get("equipment_details", "")
                st.session_state[f"{pfx}_sdate"]   = _mi_parse_date(row.get("start_date"))
                st.session_state[f"{pfx}_issue"]   = row.get("issue", "")
                st.session_state[f"{pfx}_propd"]   = _mi_parse_date(row.get("proposed_date"))
                st.session_state[f"{pfx}_actend"]  = _mi_parse_date(row.get("actual_end_date"))
                st.session_state[f"{pfx}_resoln"]  = row.get("resolution_action", "")
        else:
            st.session_state[sk_na]   = False
            st.session_state[sk_rows] = [0]
            st.session_state[sk_ctr]  = 1
        st.session_state[sk_loaded] = True

    na_val = _na_checkbox("Not Applicable — No equipment breakdown this month", sk_na)

    if not na_val:
        row_ids   = st.session_state.get(sk_rows, [])
        to_delete = None

        for i, rid in enumerate(row_ids):
            pfx = f"mi_{T}_{uid}_{month_year}_{rid}"
            st.markdown(_MI_ROW_DIV, unsafe_allow_html=True)
            hc, dc = st.columns([11, 1])
            with hc:
                st.markdown(f"**Breakdown {i + 1}**")
            with dc:
                if st.button("🗑", key=f"{pfx}_del", help="Remove"):
                    to_delete = rid

            c1, c2 = st.columns(2)
            with c1:
                en_def = st.session_state.get(f"{pfx}_eqname", EQUIP_OPTS[0])
                en_idx = EQUIP_OPTS.index(en_def) if en_def in EQUIP_OPTS else len(EQUIP_OPTS) - 1
                st.selectbox("Equipment Name *", EQUIP_OPTS, index=en_idx, key=f"{pfx}_eqname",
                             help="Select the type of equipment that broke down")
            with c2:
                if st.session_state.get(f"{pfx}_eqname") == "Other":
                    st.text_input("Specify Equipment *", key=f"{pfx}_eqother", max_chars=256,
                                  help="Describe the equipment when 'Other' is selected")

            st.text_area("Equipment Details *", key=f"{pfx}_details", max_chars=256,
                         help="Equipment tag number, make, model or other identifying details")

            c3, c4, c5, c6 = st.columns(4)
            with c3:
                st.date_input("Breakdown Start Date *", key=f"{pfx}_sdate",
                              value=st.session_state.get(f"{pfx}_sdate"), format="DD/MM/YYYY",
                              help="Date the equipment breakdown was first reported")
            with c4:
                st.text_area("Issue Description *", key=f"{pfx}_issue", max_chars=256,
                             help="Describe the nature of the breakdown or fault in detail")
            with c5:
                st.date_input("Proposed Resolution Date *", key=f"{pfx}_propd",
                              value=st.session_state.get(f"{pfx}_propd"), format="DD/MM/YYYY",
                              help="Expected date by which the breakdown will be resolved")
            with c6:
                st.date_input("Actual End Date (blank if unresolved)", key=f"{pfx}_actend",
                              value=st.session_state.get(f"{pfx}_actend"), format="DD/MM/YYYY",
                              help="Actual date equipment was restored; leave blank if still unresolved")

            st.text_area("Resolution Action *", key=f"{pfx}_resoln", max_chars=256,
                         help="Actions taken or planned to resolve the equipment breakdown")
            st.markdown("</div>", unsafe_allow_html=True)

        if to_delete is not None:
            st.session_state[sk_rows] = [r for r in st.session_state[sk_rows] if r != to_delete]
            st.rerun()

        # Summary preview
        _tbl = []
        for rid in st.session_state.get(sk_rows, []):
            pfx = f"mi_{T}_{uid}_{month_year}_{rid}"
            _tbl.append({
                "Equipment": st.session_state.get(f"{pfx}_eqname",""),
                "Start Date": _mi_fmt_date(st.session_state.get(f"{pfx}_sdate")),
                "Proposed End": _mi_fmt_date(st.session_state.get(f"{pfx}_propd")),
                "Actual End": _mi_fmt_date(st.session_state.get(f"{pfx}_actend")),
            })
        _mi_summary_table(_tbl, ["Equipment","Start Date","Proposed End","Actual End"])

        if st.button("➕ Add Breakdown", key=f"mi_{T}_add_{uid}_{month_year}"):
            new_id = st.session_state.get(sk_ctr, 0)
            st.session_state.setdefault(sk_rows, []).append(new_id)
            st.session_state[sk_ctr] = new_id + 1
            st.rerun()

    if st.button("💾 Save Equipment Breakdown Data", key=f"mi_{T}_save_{uid}_{month_year}", type="primary"):
        if na_val:
            rows_to_save = [{"row_no": "1", "na_flag": "Y",
                             "equipment_name": "NA", "equipment_name_other": "NA",
                             "equipment_details": "NA", "start_date": "NA",
                             "issue": "NA", "proposed_date": "NA",
                             "actual_end_date": "NA", "resolution_action": "NA"}]
        else:
            rows_to_save = []
            errors = []
            for i, rid in enumerate(st.session_state.get(sk_rows, [])):
                pfx    = f"mi_{T}_{uid}_{month_year}_{rid}"
                eqname = st.session_state.get(f"{pfx}_eqname", "")
                eqoth  = (st.session_state.get(f"{pfx}_eqother") or "").strip()
                if eqname == "Other" and not eqoth:
                    errors.append(f"Breakdown {i+1}: Equipment name specification is required.")
                details = (st.session_state.get(f"{pfx}_details") or "").strip()
                sdate   = _mi_fmt_date(st.session_state.get(f"{pfx}_sdate"))
                issue   = (st.session_state.get(f"{pfx}_issue") or "").strip()
                propd   = _mi_fmt_date(st.session_state.get(f"{pfx}_propd"))
                actend  = _mi_fmt_date(st.session_state.get(f"{pfx}_actend"))
                resoln  = (st.session_state.get(f"{pfx}_resoln") or "").strip()
                if not all([eqname, details, sdate, issue, propd, resoln]):
                    errors.append(f"Breakdown {i+1}: All required fields must be filled.")
                else:
                    rows_to_save.append({
                        "row_no": str(i + 1), "na_flag": "N",
                        "equipment_name": eqname, "equipment_name_other": eqoth,
                        "equipment_details": details, "start_date": sdate,
                        "issue": issue, "proposed_date": propd,
                        "actual_end_date": actend, "resolution_action": resoln,
                    })
            if errors:
                for e in errors:
                    st.error(e)
                return
            if not rows_to_save:
                st.warning("Add at least one breakdown record, or mark as Not Applicable.")
                return
        res = sheets.save_mi_data("MI_EQUIP_BREAKDOWN", uid, month_year, rows_to_save)
        if res.get("ok"):
            st.success(f"✅ Saved {res['rows']} record(s).")
            st.session_state.pop(sk_loaded, None)
        else:
            st.error(f"Save failed: {res.get('msg', '')}")


def _mi_tab_int_pipeline(uid: str, month_year: str):
    T         = "ip"
    sk_na     = f"mi_{T}_{uid}_{month_year}_na"
    sk_loaded = f"mi_{T}_{uid}_{month_year}_loaded"
    pfx       = f"mi_{T}_{uid}_{month_year}"

    DATE_FIELDS = ["last_ut_date", "last_hydrotest_date", "last_dcvg_date", "last_lrut_date"]

    if not st.session_state.get(sk_loaded):
        saved = sheets.load_mi_data("MI_INT_PIPELINE", uid, month_year)
        if saved and saved[0].get("na_flag") == "Y":
            st.session_state[sk_na] = True
        elif saved:
            row = saved[0]
            st.session_state[sk_na] = False
            for fld in DATE_FIELDS:
                st.session_state[f"{pfx}_{fld}"] = _mi_parse_date(row.get(fld))
            st.session_state[f"{pfx}_other_testing"] = row.get("other_testing", "")
        else:
            st.session_state[sk_na] = False
        st.session_state[sk_loaded] = True

    na_val = _na_checkbox("Not Applicable — No internal pipeline at this location", sk_na)

    if not na_val:
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.date_input("Last UT Date *", key=f"{pfx}_last_ut_date",
                          value=st.session_state.get(f"{pfx}_last_ut_date"), format="DD/MM/YYYY",
                          help="Date of last Ultrasonic Thickness (UT) testing of internal pipeline")
        with c2:
            st.date_input("Last Hydrotest Date *", key=f"{pfx}_last_hydrotest_date",
                          value=st.session_state.get(f"{pfx}_last_hydrotest_date"), format="DD/MM/YYYY",
                          help="Date of last hydrostatic pressure test of internal pipeline")
        with c3:
            st.date_input("Last DCVG Date *", key=f"{pfx}_last_dcvg_date",
                          value=st.session_state.get(f"{pfx}_last_dcvg_date"), format="DD/MM/YYYY",
                          help="Date of last Direct Current Voltage Gradient (DCVG) survey")
        with c4:
            st.date_input("Last LRUT Date *", key=f"{pfx}_last_lrut_date",
                          value=st.session_state.get(f"{pfx}_last_lrut_date"), format="DD/MM/YYYY",
                          help="Date of last Long Range Ultrasonic Testing (LRUT) of internal pipeline")
        st.text_area("Other Testing Details", key=f"{pfx}_other_testing", max_chars=256,
                     help="Any other testing or inspection carried out on internal pipeline")

    if st.button("💾 Save Internal Pipeline Data", key=f"mi_{T}_save_{uid}_{month_year}", type="primary"):
        if na_val:
            rec = {"na_flag": "Y", "last_ut_date": "NA", "last_hydrotest_date": "NA",
                   "last_dcvg_date": "NA", "last_lrut_date": "NA", "other_testing": "NA"}
        else:
            dates = {fld: _mi_fmt_date(st.session_state.get(f"{pfx}_{fld}")) for fld in DATE_FIELDS}
            missing = [fld.replace("_", " ").title() for fld, v in dates.items() if not v]
            if missing:
                st.error("Required dates missing: " + ", ".join(missing))
                return
            rec = {"na_flag": "N", **dates,
                   "other_testing": (st.session_state.get(f"{pfx}_other_testing") or "").strip()}
        res = sheets.save_mi_data("MI_INT_PIPELINE", uid, month_year, [rec])
        if res.get("ok"):
            st.success("✅ Internal Pipeline data saved.")
            st.session_state.pop(sk_loaded, None)
        else:
            st.error(f"Save failed: {res.get('msg', '')}")


def _mi_tab_ext_pipeline(uid: str, month_year: str):
    T         = "ep"
    sk_na     = f"mi_{T}_{uid}_{month_year}_na"
    sk_rows   = f"mi_{T}_{uid}_{month_year}_rows"
    sk_ctr    = f"mi_{T}_{uid}_{month_year}_ctr"
    sk_loaded = f"mi_{T}_{uid}_{month_year}_loaded"

    DATE_FIELDS    = ["last_ut_date", "last_hydrotest_date", "last_dcvg_date", "last_lrut_date"]
    PIPE_TYPE_OPTS = ["UG", "AG"]

    if not st.session_state.get(sk_loaded):
        saved = sheets.load_mi_data("MI_EXT_PIPELINE", uid, month_year)
        if saved and saved[0].get("na_flag") == "Y":
            st.session_state[sk_na]   = True
            st.session_state[sk_rows] = []
            st.session_state[sk_ctr]  = 0
        elif saved:
            ids = list(range(len(saved)))
            st.session_state[sk_na]   = False
            st.session_state[sk_rows] = ids
            st.session_state[sk_ctr]  = len(saved)
            for rid, row in zip(ids, saved):
                pfx = f"mi_{T}_{uid}_{month_year}_{rid}"
                pt_raw = row.get("pipeline_type", "UG")
                st.session_state[f"{pfx}_pipeline_type"] = (
                    pt_raw if pt_raw in PIPE_TYPE_OPTS else "UG"
                )
                st.session_state[f"{pfx}_pipeline_details"] = row.get("pipeline_details", "")
                st.session_state[f"{pfx}_length_metres"]    = row.get("length_metres", "")
                st.session_state[f"{pfx}_product"]          = row.get("product", "")
                st.session_state[f"{pfx}_size_inch"]        = row.get("size_inch", "")
                for fld in DATE_FIELDS:
                    st.session_state[f"{pfx}_{fld}"] = _mi_parse_date(row.get(fld))
                st.session_state[f"{pfx}_other_testing"] = row.get("other_testing", "")
        else:
            st.session_state[sk_na]   = False
            st.session_state[sk_rows] = [0]
            st.session_state[sk_ctr]  = 1
        st.session_state[sk_loaded] = True

    na_val = _na_checkbox("Not Applicable — No external pipeline at this location", sk_na)

    if not na_val:
        st.caption("Add one row per external pipeline segment / product.")
        row_ids   = st.session_state.get(sk_rows, [])
        to_delete = None

        for i, rid in enumerate(row_ids):
            pfx = f"mi_{T}_{uid}_{month_year}_{rid}"
            st.markdown(_MI_ROW_DIV, unsafe_allow_html=True)
            hc, dc = st.columns([11, 1])
            with hc:
                st.markdown(f"**Pipeline {i + 1}**")
            with dc:
                if st.button("🗑", key=f"{pfx}_del", help="Remove this pipeline row"):
                    to_delete = rid

            _pt_def = st.session_state.get(f"{pfx}_pipeline_type", "UG")
            _pt_idx = PIPE_TYPE_OPTS.index(_pt_def) if _pt_def in PIPE_TYPE_OPTS else 0
            st.selectbox("Pipeline Type *", PIPE_TYPE_OPTS, index=_pt_idx,
                         key=f"{pfx}_pipeline_type",
                         help="UG = Underground pipeline  /  AG = Above Ground pipeline")

            st.text_area("Pipeline Details *", key=f"{pfx}_pipeline_details", max_chars=256,
                         help="Describe the pipeline segment (route, from-to, purpose)")

            c1, c2, c3 = st.columns(3)
            with c1:
                st.text_input("Length (metres) *", key=f"{pfx}_length_metres",
                              help="Total length of this pipeline segment in metres")
            with c2:
                st.text_input("Product *", key=f"{pfx}_product",
                              help="Product carried — e.g. MS, HSD, ATF, SKO")
            with c3:
                st.text_input("Size (inch) *", key=f"{pfx}_size_inch",
                              help="Pipeline nominal bore in inches")

            c4, c5, c6, c7 = st.columns(4)
            with c4:
                st.date_input("Last UT Date *", key=f"{pfx}_last_ut_date",
                              value=st.session_state.get(f"{pfx}_last_ut_date"),
                              format="DD/MM/YYYY",
                              help="Date of last Ultrasonic Thickness testing")
            with c5:
                st.date_input("Last Hydrotest Date *", key=f"{pfx}_last_hydrotest_date",
                              value=st.session_state.get(f"{pfx}_last_hydrotest_date"),
                              format="DD/MM/YYYY",
                              help="Date of last hydrostatic pressure test")
            with c6:
                st.date_input("Last DCVG Date *", key=f"{pfx}_last_dcvg_date",
                              value=st.session_state.get(f"{pfx}_last_dcvg_date"),
                              format="DD/MM/YYYY",
                              help="Date of last DC Voltage Gradient survey")
            with c7:
                st.date_input("Last LRUT Date *", key=f"{pfx}_last_lrut_date",
                              value=st.session_state.get(f"{pfx}_last_lrut_date"),
                              format="DD/MM/YYYY",
                              help="Date of last Long Range Ultrasonic Testing")
            st.text_area("Other Testing Details", key=f"{pfx}_other_testing", max_chars=256,
                         help="Any other inspection / testing method used for this segment")
            st.markdown("</div>", unsafe_allow_html=True)

        if to_delete is not None:
            st.session_state[sk_rows] = [r for r in st.session_state[sk_rows] if r != to_delete]
            st.rerun()

        # Summary table
        _tbl = []
        for rid in st.session_state.get(sk_rows, []):
            pfx = f"mi_{T}_{uid}_{month_year}_{rid}"
            _tbl.append({
                "Type": st.session_state.get(f"{pfx}_pipeline_type", ""),
                "Pipeline Details": (st.session_state.get(f"{pfx}_pipeline_details") or "")[:40],
                "Product": st.session_state.get(f"{pfx}_product", ""),
                "Length (m)": st.session_state.get(f"{pfx}_length_metres", ""),
                "Size (in)": st.session_state.get(f"{pfx}_size_inch", ""),
                "Last UT": _mi_fmt_date(st.session_state.get(f"{pfx}_last_ut_date")),
            })
        _mi_summary_table(_tbl, ["Type", "Pipeline Details", "Product", "Length (m)", "Size (in)", "Last UT"])

        if st.button("➕ Add Pipeline Segment", key=f"mi_{T}_add_{uid}_{month_year}"):
            new_id = st.session_state.get(sk_ctr, 0)
            st.session_state.setdefault(sk_rows, []).append(new_id)
            st.session_state[sk_ctr] = new_id + 1
            st.rerun()

    if st.button("💾 Save External Pipeline Data", key=f"mi_{T}_save_{uid}_{month_year}",
                 type="primary"):
        if na_val:
            rows_to_save = [{"na_flag": "Y", "pipeline_type": "NA", "pipeline_details": "NA",
                             "length_metres": "NA", "product": "NA", "size_inch": "NA",
                             "last_ut_date": "NA", "last_hydrotest_date": "NA",
                             "last_dcvg_date": "NA", "last_lrut_date": "NA",
                             "other_testing": "NA"}]
        else:
            rows_to_save = []
            errors = []
            for i, rid in enumerate(st.session_state.get(sk_rows, [])):
                pfx  = f"mi_{T}_{uid}_{month_year}_{rid}"
                pt   = (st.session_state.get(f"{pfx}_pipeline_type") or "UG").strip()
                pd_  = (st.session_state.get(f"{pfx}_pipeline_details") or "").strip()
                lm   = (st.session_state.get(f"{pfx}_length_metres") or "").strip()
                prd  = (st.session_state.get(f"{pfx}_product") or "").strip()
                sz   = (st.session_state.get(f"{pfx}_size_inch") or "").strip()
                dates = {fld: _mi_fmt_date(st.session_state.get(f"{pfx}_{fld}")) for fld in DATE_FIELDS}
                missing_d = [fld.replace("_"," ").title() for fld, v in dates.items() if not v]
                if not all([pd_, lm, prd, sz]):
                    errors.append(f"Pipeline {i+1}: Details, length, product and size are required.")
                elif missing_d:
                    errors.append(f"Pipeline {i+1}: Dates missing — {', '.join(missing_d)}")
                else:
                    rows_to_save.append({
                        "row_no": str(i + 1), "na_flag": "N",
                        "pipeline_type": pt, "pipeline_details": pd_,
                        "length_metres": lm, "product": prd, "size_inch": sz, **dates,
                        "other_testing": (st.session_state.get(f"{pfx}_other_testing") or "").strip(),
                    })
            if errors:
                for e in errors:
                    st.error(e)
                return
            if not rows_to_save:
                st.warning("Add at least one pipeline segment, or mark as Not Applicable.")
                return
        res = sheets.save_mi_data("MI_EXT_PIPELINE", uid, month_year, rows_to_save)
        if res.get("ok"):
            st.success(f"✅ Saved {res['rows']} external pipeline record(s).")
            st.session_state.pop(sk_loaded, None)
        else:
            st.error(f"Save failed: {res.get('msg', '')}")


def _mi_tab_tank_status(uid: str, month_year: str, tank_opts: list,
                        zone: str, loc_name: str):
    T         = "ts"
    sk_na     = f"mi_{T}_{uid}_{month_year}_na"
    sk_rows   = f"mi_{T}_{uid}_{month_year}_rows"
    sk_ctr    = f"mi_{T}_{uid}_{month_year}_ctr"
    sk_loaded = f"mi_{T}_{uid}_{month_year}_loaded"

    STATUS_OPTS = ["Operational", "Under Repair", "Under Cleaning",
                   "Idle", "Revamp", "Others"]
    EXT_OPTS    = ["Yes", "No", "NA"]

    if not st.session_state.get(sk_loaded):
        saved = sheets.load_mi_data("MI_TANK_STATUS", uid, month_year)
        if saved and saved[0].get("na_flag") == "Y":
            st.session_state[sk_na]   = True
            st.session_state[sk_rows] = []
            st.session_state[sk_ctr]  = 0
        elif saved:
            ids = list(range(len(saved)))
            st.session_state[sk_na]   = False
            st.session_state[sk_rows] = ids
            st.session_state[sk_ctr]  = len(saved)
            for rid, row in zip(ids, saved):
                pfx = f"mi_{T}_{uid}_{month_year}_{rid}"
                tn  = row.get("tank_no", "")
                st.session_state[f"{pfx}_tank"]  = (
                    tn if tn in tank_opts else (tank_opts[0] if tank_opts else "Other Tanks")
                )
                for df in ("cleaning_completed_date", "cleaning_due_date",
                           "inspection_date", "inspection_due_date",
                           "painting_date", "painting_due_date"):
                    st.session_state[f"{pfx}_{df}"] = _mi_parse_date(row.get(df))
                ext_raw = row.get("extension_taken", "NA")
                st.session_state[f"{pfx}_ext"] = (
                    ext_raw if ext_raw in EXT_OPTS else "NA"
                )
                st.session_state[f"{pfx}_efn"]       = row.get("extension_efn_no", "")
                st_raw = row.get("tank_status", STATUS_OPTS[0])
                st.session_state[f"{pfx}_status"] = (
                    st_raw if st_raw in STATUS_OPTS else STATUS_OPTS[0]
                )
                st.session_state[f"{pfx}_status_other"] = row.get("tank_status_other", "")
        else:
            st.session_state[sk_na]   = False
            st.session_state[sk_rows] = [0]
            st.session_state[sk_ctr]  = 1
        st.session_state[sk_loaded] = True

    na_val = _na_checkbox("Not Applicable — No tanks at this location", sk_na)

    if not na_val:
        row_ids   = st.session_state.get(sk_rows, [])
        to_delete = None

        # ── Bulk-delete toolbar (checkbox per row) ─────────────────────────
        _sel_key = f"mi_{T}_{uid}_{month_year}_sel"
        selected_rids = st.session_state.get(_sel_key, set())
        if selected_rids:
            _bdc1, _bdc2 = st.columns([3, 1])
            with _bdc1:
                st.caption(f"{len(selected_rids)} row(s) selected")
            with _bdc2:
                if st.button("🗑 Delete Selected", key=f"mi_{T}_bulk_del_{uid}_{month_year}",
                             type="secondary", use_container_width=True):
                    st.session_state[sk_rows] = [r for r in row_ids if r not in selected_rids]
                    st.session_state[_sel_key] = set()
                    st.rerun()

        for i, rid in enumerate(row_ids):
            pfx = f"mi_{T}_{uid}_{month_year}_{rid}"
            st.markdown(_MI_ROW_DIV, unsafe_allow_html=True)
            chk_col, hc, dc = st.columns([1, 10, 1])
            with chk_col:
                is_sel = st.checkbox("", key=f"{pfx}_chk", value=(rid in selected_rids),
                                     label_visibility="collapsed",
                                     help="Select to bulk-delete")
                if is_sel:
                    selected_rids.add(rid)
                else:
                    selected_rids.discard(rid)
                st.session_state[_sel_key] = selected_rids
            with hc:
                st.markdown(f"**Tank {i + 1}**")
            with dc:
                if st.button("🗑", key=f"{pfx}_del", help="Remove this row"):
                    to_delete = rid

            # Tank selection
            tn_def = st.session_state.get(f"{pfx}_tank",
                                          tank_opts[0] if tank_opts else "Other Tanks")
            tn_idx = tank_opts.index(tn_def) if tn_def in tank_opts else len(tank_opts) - 1
            st.selectbox("Tank No. *", tank_opts, index=tn_idx, key=f"{pfx}_tank",
                         help="Select tank number from Tank Master list")

            # Cleaning dates (Y, Z)
            c1, c2 = st.columns(2)
            with c1:
                st.date_input("Date of Cleaning Completed (Y)",
                              key=f"{pfx}_cleaning_completed_date",
                              value=st.session_state.get(f"{pfx}_cleaning_completed_date"),
                              format="DD/MM/YYYY",
                              help="Date tank cleaning was completed")
            with c2:
                st.date_input("Due Date of Tank Cleaning (Z)",
                              key=f"{pfx}_cleaning_due_date",
                              value=st.session_state.get(f"{pfx}_cleaning_due_date"),
                              format="DD/MM/YYYY",
                              help="Next scheduled tank cleaning due date")

            # Extension (AA)
            ext_def = st.session_state.get(f"{pfx}_ext", "NA")
            ext_idx = EXT_OPTS.index(ext_def) if ext_def in EXT_OPTS else 2
            c3, c4 = st.columns(2)
            with c3:
                st.selectbox("Extension Taken? (AA)", EXT_OPTS, index=ext_idx,
                             key=f"{pfx}_ext",
                             help="Select Yes if cleaning extension has been taken")
            with c4:
                if st.session_state.get(f"{pfx}_ext") == "Yes":
                    st.text_input("eFN# (Extension Order No.) *",
                                  key=f"{pfx}_efn", max_chars=50,
                                  help="Enter eFN number for the extension taken")

            # Inspection dates (AB, AC)
            c5, c6 = st.columns(2)
            with c5:
                st.date_input("Date of Comprehensive Inspection (AB)",
                              key=f"{pfx}_inspection_date",
                              value=st.session_state.get(f"{pfx}_inspection_date"),
                              format="DD/MM/YYYY",
                              help="Date of last comprehensive inspection")
            with c6:
                st.date_input("Due Date for Comprehensive Inspection (AC)",
                              key=f"{pfx}_inspection_due_date",
                              value=st.session_state.get(f"{pfx}_inspection_due_date"),
                              format="DD/MM/YYYY",
                              help="Next comprehensive inspection due date")

            # Painting dates (AD, AE)
            c7, c8 = st.columns(2)
            with c7:
                st.date_input("Date of Tank Painting (AD)",
                              key=f"{pfx}_painting_date",
                              value=st.session_state.get(f"{pfx}_painting_date"),
                              format="DD/MM/YYYY",
                              help="Date of last tank painting")
            with c8:
                st.date_input("Due Date of Tank Painting (AE)",
                              key=f"{pfx}_painting_due_date",
                              value=st.session_state.get(f"{pfx}_painting_due_date"),
                              format="DD/MM/YYYY",
                              help="Next tank painting due date")

            # Tank Status (AF)
            st_def = st.session_state.get(f"{pfx}_status", STATUS_OPTS[0])
            st_idx = STATUS_OPTS.index(st_def) if st_def in STATUS_OPTS else 0
            c9, c10 = st.columns(2)
            with c9:
                st.selectbox("Tank Status (AF) *", STATUS_OPTS, index=st_idx,
                             key=f"{pfx}_status",
                             help="Current operational status of this tank")
            with c10:
                if st.session_state.get(f"{pfx}_status") == "Others":
                    st.text_input("Tank Status Details *",
                                  key=f"{pfx}_status_other", max_chars=128,
                                  help="Describe the tank status if 'Others' is selected")

            st.markdown("</div>", unsafe_allow_html=True)

        if to_delete is not None:
            st.session_state[sk_rows] = [r for r in st.session_state[sk_rows]
                                          if r != to_delete]
            st.rerun()

        # Summary preview with per-row delete
        _row_ids_snap = list(st.session_state.get(sk_rows, []))
        if _row_ids_snap:
            st.markdown(
                '<div style="font-size:12px;font-weight:700;color:#001a6e;'
                'margin:10px 0 4px;letter-spacing:0.3px;">&#128202; Summary — tanks entered so far</div>',
                unsafe_allow_html=True,
            )
            _hc = st.columns([2, 2, 2, 2, 2, 1])
            for _hi, _hl in enumerate(["Tank No.", "Status", "Cleaning Due",
                                        "Inspection Due", "Painting Due", ""]):
                _hc[_hi].markdown(
                    f'<div style="font-size:10px;font-weight:700;color:#fff;'
                    f'background:#002b8f;padding:4px 6px;border-radius:4px;">{_hl}</div>',
                    unsafe_allow_html=True)
            _del_from_summary = None
            for _sr_idx, _rid in enumerate(_row_ids_snap):
                _pfx = f"mi_{T}_{uid}_{month_year}_{_rid}"
                _rc  = st.columns([2, 2, 2, 2, 2, 1])
                _bg  = "#f0f4ff" if _sr_idx % 2 == 0 else "#ffffff"
                _vals = [
                    st.session_state.get(f"{_pfx}_tank", ""),
                    st.session_state.get(f"{_pfx}_status", ""),
                    _mi_fmt_date(st.session_state.get(f"{_pfx}_cleaning_due_date")) or "—",
                    _mi_fmt_date(st.session_state.get(f"{_pfx}_inspection_due_date")) or "—",
                    _mi_fmt_date(st.session_state.get(f"{_pfx}_painting_due_date")) or "—",
                ]
                for _ci, _v in enumerate(_vals):
                    _rc[_ci].markdown(
                        f'<div style="font-size:11px;padding:4px 6px;background:{_bg};'
                        f'border-radius:3px;">{_v}</div>',
                        unsafe_allow_html=True)
                with _rc[5]:
                    if st.button("🗑", key=f"mi_{T}_sdel_{uid}_{month_year}_{_rid}",
                                 help="Delete this tank row"):
                        _del_from_summary = _rid
            if _del_from_summary is not None:
                st.session_state[sk_rows] = [r for r in st.session_state[sk_rows]
                                              if r != _del_from_summary]
                st.rerun()

        if st.button("➕ Add Tank Row", key=f"mi_{T}_add_{uid}_{month_year}"):
            new_id = st.session_state.get(sk_ctr, 0)
            st.session_state.setdefault(sk_rows, []).append(new_id)
            st.session_state[sk_ctr] = new_id + 1
            st.rerun()

    if st.button("💾 Save Tank Status Data", key=f"mi_{T}_save_{uid}_{month_year}",
                 type="primary"):
        if na_val:
            rows_to_save = [{"row_no": "1", "na_flag": "Y",
                             "zone": zone, "loc_name": loc_name,
                             "tank_no": "NA",
                             "cleaning_completed_date": "NA", "cleaning_due_date": "NA",
                             "extension_taken": "NA", "extension_efn_no": "NA",
                             "inspection_date": "NA", "inspection_due_date": "NA",
                             "painting_date": "NA", "painting_due_date": "NA",
                             "tank_status": "NA", "tank_status_other": "NA"}]
        else:
            rows_to_save = []
            errors = []
            for i, rid in enumerate(st.session_state.get(sk_rows, [])):
                pfx    = f"mi_{T}_{uid}_{month_year}_{rid}"
                tank   = st.session_state.get(f"{pfx}_tank", "")
                ext    = st.session_state.get(f"{pfx}_ext", "NA")
                efn    = (st.session_state.get(f"{pfx}_efn") or "").strip()
                status = st.session_state.get(f"{pfx}_status", "")
                s_other = (st.session_state.get(f"{pfx}_status_other") or "").strip()

                date_fields = [
                    "cleaning_completed_date", "cleaning_due_date",
                    "inspection_date", "inspection_due_date",
                    "painting_date", "painting_due_date",
                ]
                dates = {df: _mi_fmt_date(st.session_state.get(f"{pfx}_{df}"))
                         for df in date_fields}

                if not tank:
                    errors.append(f"Tank {i+1}: Tank No. is required.")
                if ext == "Yes" and not efn:
                    errors.append(f"Tank {i+1}: eFN# is required when Extension = Yes.")
                if not status:
                    errors.append(f"Tank {i+1}: Tank Status is required.")
                if status == "Others" and not s_other:
                    errors.append(f"Tank {i+1}: Tank Status Details required when status = Others.")
                rows_to_save.append({
                    "row_no": str(i + 1), "na_flag": "N",
                    "zone": zone, "loc_name": loc_name,
                    "tank_no": tank,
                    "cleaning_completed_date": dates["cleaning_completed_date"] or "",
                    "cleaning_due_date": dates["cleaning_due_date"] or "",
                    "extension_taken": ext,
                    "extension_efn_no": efn,
                    "inspection_date": dates["inspection_date"] or "",
                    "inspection_due_date": dates["inspection_due_date"] or "",
                    "painting_date": dates["painting_date"] or "",
                    "painting_due_date": dates["painting_due_date"] or "",
                    "tank_status": status,
                    "tank_status_other": s_other,
                })
            if errors:
                for e in errors:
                    st.error(e)
                return
            if not rows_to_save:
                st.warning("Add at least one tank row, or mark as Not Applicable.")
                return
        res = sheets.save_mi_data("MI_TANK_STATUS", uid, month_year, rows_to_save)
        if res.get("ok"):
            st.success(f"✅ Tank Status saved — {res['rows']} tank(s).")
            st.session_state.pop(sk_loaded, None)
            sheets.check_mi_complete.clear()
            sheets.get_full_tank_master_excel.clear()   # force fresh download next time
        else:
            st.error(f"Save failed: {res.get('msg', '')}")


def show_mi_mis_page(user: dict, month_year: str, month_label: str):
    """Full-page M&I MIS entry form with 10 subsection tabs."""
    _dashboard_css()
    uid  = user.get("userId", "")

    with st.sidebar:
        spb = _assets().get("side_panel_banner")
        if spb:
            st.markdown(
                f'<div style="margin:0;padding:0;line-height:0;width:100%;overflow:hidden;">'
                f'<img src="{spb}" style="width:100%;height:auto;display:block;'
                f'border-radius:0;margin:0;padding:0;"></div>',
                unsafe_allow_html=True,
            )

        if st.button("🏠  Back to Dashboard", key="mi_back", use_container_width=True):
            st.session_state.selected_section = None
            st.rerun()

        st.markdown("""
        <div style="padding:10px 16px 4px;">
          <div style="color:#C8D7FF;font-size:9px;font-weight:700;
                      letter-spacing:2px;text-transform:uppercase;">M&amp;I MIS</div>
          <div style="color:#8AABFF;font-size:10px;margin-top:2px;">
            Maintenance &amp; Inspection Data
          </div>
        </div>
        """, unsafe_allow_html=True)

        # ── Tank Master download in sidebar ───────────────────────────────
        st.markdown(
            '<div style="padding:6px 16px 2px;">'
            '<div style="color:#C8D7FF;font-size:9px;font-weight:700;'
            'letter-spacing:2px;text-transform:uppercase;margin-bottom:6px;">'
            'Tank Master</div></div>',
            unsafe_allow_html=True,
        )
        try:
            import base64 as _b64mod
            _tm_bytes = sheets.get_full_tank_master_excel(location_code=uid)
            _tm_b64   = _b64mod.b64encode(_tm_bytes).decode()
            _tm_fname = f"TankMaster_{uid}.xlsx"
            _tm_mime  = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            st.markdown(
                f'<a href="data:{_tm_mime};base64,{_tm_b64}" download="{_tm_fname}"'
                f' style="display:flex;align-items:center;justify-content:flex-start;'
                f'width:calc(100% - 16px);margin:2px 8px;padding:8px 13px;'
                f'background:rgba(255,255,255,0.06);color:#ffffff;'
                f'border:1px solid rgba(255,255,255,0.08);border-radius:8px;'
                f'font-size:11.5px;font-weight:500;text-decoration:none;'
                f'box-sizing:border-box;cursor:pointer;'
                f'transition:background 0.15s ease;">'
                f'&#11015;&#65039;&nbsp; Download Tank Master</a>',
                unsafe_allow_html=True,
            )
        except Exception as _ex:
            st.error(f"Tank Master unavailable: {_ex}")

    _dash_header(user)

    st.markdown(
        f'<div style="background:linear-gradient(90deg,#1a1a6e,#002B8F);color:white;'
        f'border-radius:10px;padding:14px 22px;margin-bottom:16px;">'
        f'<div style="font-size:16px;font-weight:700;">M&amp;I MIS — {month_label}</div>'
        f'<div style="font-size:12px;opacity:0.8;margin-top:3px;">'
        f'Maintenance &amp; Inspection Monthly Information System</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    if not st.session_state.get("_mi_tabs_ensured"):
        with st.spinner("Preparing M&I worksheets…"):
            sheets.ensure_mi_tabs()
        st.session_state["_mi_tabs_ensured"] = True

    zone      = user.get("zone", "")
    loc_name  = user.get("locName", "")
    tank_master = sheets.get_tank_master()
    tank_list   = tank_master.get(uid, []) + ["Other Tanks"]

    # ── Per-tab completion status ─────────────────────────────────────────
    _MI_TAB_KEYS = [
        "MI_TANK_OUTAGE", "MI_MAJOR_REPAIR", "MI_VRU",
        "MI_AUDIT_2526",  "MI_AUDIT_2627",   "MI_TECH_AUDIT",
        "MI_EQUIP_BREAKDOWN", "MI_INT_PIPELINE", "MI_EXT_PIPELINE",
        "MI_TANK_STATUS",
    ]
    _MI_T_CODES = ["to","mr","vr","a25","a27","ta","eb","ip","ep","ts"]
    _comp_cache_key = f"_mi_comp_{uid}_{month_year}"
    # Clear cache if any tab's sk_loaded was popped by a save handler (key missing)
    for _tc in _MI_T_CODES:
        if f"mi_{_tc}_{uid}_{month_year}_loaded" not in st.session_state:
            st.session_state.pop(_comp_cache_key, None)
            break
    if _comp_cache_key not in st.session_state:
        _tab_done = []
        for tk in _MI_TAB_KEYS:
            rows = sheets.load_mi_data(tk, uid, month_year)
            _tab_done.append(bool(rows))
        st.session_state[_comp_cache_key] = _tab_done
    tab_done = st.session_state[_comp_cache_key]

    def _badge(i):
        return "✅" if tab_done[i] else "❌"

    # ── How-to note for NA sections ──────────────────────────────────────
    st.markdown(
        '<div style="background:#fff8e1;border:1px solid #fbc02d;border-radius:8px;'
        'padding:10px 16px;margin-bottom:12px;font-size:12.5px;color:#5d4037;line-height:1.7;">'
        '<b>&#9888; How to complete all 10 sections:</b><br>'
        'For each section that applies to your location, enter the relevant data and save '
        '(the tab will turn ✅). For sections that do <b>not</b> apply — e.g. no VRU '
        'installed, no external pipeline, no tank outage this month — tick the '
        '<b>Not Applicable</b> checkbox visible inside the tab (also turns ✅). '
        'All 10 sections must be ✅ before the Generate M&amp;I Report button becomes active.'
        '</div>',
        unsafe_allow_html=True,
    )

    # Tank Master download is available in the sidebar (⬇️ Download Tank Master button).

    # ── Row 1 tabs ────────────────────────────────────────────────────────
    r1_labels = [
        f"{_badge(0)} 🛢️ Tank Outage",
        f"{_badge(1)} 🔧 Major Repair",
        f"{_badge(2)} 💨 VRU",
        f"{_badge(3)} 📋 M&I Audit 25-26",
        f"{_badge(4)} 📋 M&I Audit 26-27",
    ]
    tabs1 = st.tabs(r1_labels)
    with tabs1[0]:
        _mi_tab_outage(uid, month_year, tank_list)
    with tabs1[1]:
        _mi_tab_repair(uid, month_year, tank_list)
    with tabs1[2]:
        _mi_tab_vru(uid, month_year)
    with tabs1[3]:
        _mi_tab_audit2526(uid, month_year)
    with tabs1[4]:
        _mi_tab_audit2627(uid, month_year)

    st.markdown("<div style='height:4px;'></div>", unsafe_allow_html=True)

    # ── Row 2 tabs ────────────────────────────────────────────────────────
    r2_labels = [
        f"{_badge(5)} 🔍 Tech. Audit",
        f"{_badge(6)} ⚙️ Equip. Breakdown",
        f"{_badge(7)} 🔗 Int. Pipeline",
        f"{_badge(8)} 🔗 Ext. Pipeline",
        f"{_badge(9)} 📊 Tank Status",
    ]
    tabs2 = st.tabs(r2_labels)
    with tabs2[0]:
        _mi_tab_tech_audit(uid, month_year)
    with tabs2[1]:
        _mi_tab_equip(uid, month_year)
    with tabs2[2]:
        _mi_tab_int_pipeline(uid, month_year)
    with tabs2[3]:
        _mi_tab_ext_pipeline(uid, month_year)
    with tabs2[4]:
        _mi_tab_tank_status(uid, month_year, tank_list, zone, loc_name)

    # Inject hover/focus tooltips for S5 standard fields (M&I Index, PM %, etc.)
    _inject_field_enhancements(S5_FIELDS)

    # ── Generate M&I MIS Report button ───────────────────────────────────
    st.markdown(
        '<div style="border-top:1px solid #dde3ed;margin-top:18px;padding-top:12px;"></div>',
        unsafe_allow_html=True,
    )
    mi_done = all(tab_done)
    _mi_rpt_key = f"_mi_rpt_{uid}_{month_year}"
    c1, c2 = st.columns([3, 1])
    with c1:
        if mi_done:
            st.caption("All 10 M&I tabs are complete. You can generate the M&I MIS Report.")
        else:
            pending_tabs = [
                ["Tank Outage","Major Repair","VRU","M&I Audit 25-26","M&I Audit 26-27",
                 "Tech. Audit","Equip. Bkdn","Int. Pipeline","Ext. Pipeline","Tank Status"][i]
                for i, d in enumerate(tab_done) if not d
            ]
            st.caption(f"Pending tabs: {', '.join(pending_tabs)}")
    with c2:
        if st.button("📊 Generate M&I Report", key="btn_gen_mi_rpt",
                     use_container_width=True, disabled=not mi_done):
            with st.spinner("Generating M&I MIS Report…"):
                try:
                    rpt_bytes = sheets.generate_mi_mis_report(uid, month_year, user)
                    st.session_state[_mi_rpt_key] = rpt_bytes
                except Exception as ex:
                    st.error(f"Report error: {ex}")
                    st.session_state[_mi_rpt_key] = None

    if st.session_state.get(_mi_rpt_key):
        st.download_button(
            label="⬇️ Download M&I MIS Report",
            data=st.session_state[_mi_rpt_key],
            file_name=f"MI_MIS_{uid}_{month_year.replace('-','_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key=f"dl_mi_{uid}_{month_year}",
        )

    st.markdown("""
    <div style="margin-top:16px;padding:10px 4px;border-top:1px solid #dde3ed;
                display:flex;justify-content:space-between;font-size:11px;color:#aaa;">
      <span>&#169; 2026 Hindustan Petroleum Corporation Limited.</span>
      <span>HPCL SOD &nbsp;·&nbsp; M&amp;I MIS</span>
    </div>""", unsafe_allow_html=True)


# ── Router ────────────────────────────────────────────────────────────────────

def main():
    # Clear the asset cache once per browser session so a redeployed logo is
    # picked up without requiring a full server restart on localhost.
    if not st.session_state.get("_assets_loaded"):
        _assets.clear()
        st.session_state["_assets_loaded"] = True
    _base_css()
    st.session_state.setdefault("page", "login")
    st.session_state.setdefault("user", None)

    # ── Security checks (only when user is on the dashboard) ─────────────────
    import time as _tm
    _INACTIVITY_TIMEOUT = 1800   # 30 minutes
    _SESSION_CHECK_INTERVAL = 300  # verify token every 5 minutes

    if (st.session_state.get("user") and
            st.session_state.get("page") == "dashboard"):

        # 1. Inactivity timeout
        _last_act = st.session_state.get("last_activity", _tm.time())
        if _tm.time() - _last_act > _INACTIVITY_TIMEOUT:
            _uid = st.session_state.user.get("userId", "")
            if _uid:
                sheets.clear_session(_uid)
            st.session_state.user = None
            st.session_state.page = "login"
            st.session_state["_timeout_msg"] = True
            st.rerun()
        else:
            st.session_state["last_activity"] = _tm.time()

        # 2. Simultaneous login detection (checked every 5 min)
        _last_chk = st.session_state.get("_last_session_check", 0)
        if _tm.time() - _last_chk > _SESSION_CHECK_INTERVAL:
            st.session_state["_last_session_check"] = _tm.time()
            _uid   = st.session_state.user.get("userId", "")
            _token = st.session_state.user.get("_session_token", "")
            if _uid and _token and not sheets.check_session_valid(_uid, _token):
                st.session_state.user = None
                st.session_state.page = "login"
                st.session_state["_displaced_msg"] = True
                st.rerun()

    page = st.session_state.page
    user = st.session_state.user

    role = (user or {}).get("role", "")

    if user is None or page == "login":
        show_login()
    elif page == "change_password":
        show_change_password()
    elif page == "dashboard":
        if role == "Zone":
            sec = st.session_state.get("selected_section")
            if sec == "chatbot":
                show_chatbot_page(user)
            elif sec == "analytics":
                show_analytics_page(user)
            elif sec == "reports":
                show_reports_page(user)
            elif sec == "review":
                show_review(user,
                            st.session_state.get("current_month", ""),
                            st.session_state.get("current_month_label", ""))
            else:
                show_zone_dashboard(user)
        elif role in ("Admin", "Viewer"):
            sec = st.session_state.get("selected_section")
            if sec == "chatbot":
                show_chatbot_page(user)
            elif sec == "analytics":
                show_analytics_page(user)
            elif sec == "reports":
                show_reports_page(user)
            elif sec == "review":
                show_review(user,
                            st.session_state.get("current_month", ""),
                            st.session_state.get("current_month_label", ""))
            elif sec == "email_review" and role == "Admin":
                show_email_review(user)
            else:
                show_hqo_dashboard(user)
        else:
            # Maker / Checker
            sec   = st.session_state.get("selected_section")
            month = st.session_state.get("current_month")
            label = st.session_state.get("current_month_label", "")
            if sec == "chatbot":
                show_chatbot_page(user)
            elif sec == "analytics":
                show_analytics_page(user)
            elif sec == "review" and month:
                show_review(user, month, label)
            elif sec == "mi_mis" and month:
                show_mi_mis_page(user, month, label)
            elif isinstance(sec, int) and 1 <= sec <= 10 and month:
                show_section_form(sec, user, month, label)
            else:
                show_dashboard()
    else:
        st.session_state.page = "login"
        st.rerun()


if __name__ == "__main__":
    main()
