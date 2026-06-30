"""Google Sheets backend for HPCL SOD MIS."""

# ── SSL bypass for corporate networks with SSL-inspection proxies ─────────────
# Must run before any HTTPS-related imports.
import ssl
import urllib3

ssl._create_default_https_context = ssl._create_unverified_context
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

try:
    from requests.adapters import HTTPAdapter as _HA
    _orig_send = _HA.send
    def _send_no_ssl(self, request, **kw):
        kw["verify"] = False
        return _orig_send(self, request, **kw)
    _HA.send = _send_no_ssl
except Exception:
    pass
# ─────────────────────────────────────────────────────────────────────────────

import json
import time
from datetime import datetime, date

import gspread
import streamlit as st
from google.oauth2.service_account import Credentials


def _api_call(fn, *args, retries: int = 6, **kwargs):
    """Execute a Sheets API call with exponential backoff on 429 quota errors.
    Waits 2^n seconds between attempts (1 s, 2 s, 4 s, 8 s, 16 s, 32 s)."""
    for n in range(retries):
        try:
            return fn(*args, **kwargs)
        except gspread.exceptions.APIError as exc:
            resp = getattr(exc, "response", None)
            code = getattr(resp, "status_code", 0) if resp else 0
            if code == 429 and n < retries - 1:
                time.sleep(min(2 ** n, 32))
            else:
                raise

_SCOPES = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]

TABS = {
    "USER_ACCESS":        "UserAccess",
    "LOCATION_MASTER":    "LocationMaster",
    "HELPDESK":           "HelpDesk",
    "AUDIT_LOG":          "AuditLog",
    "MIS_DRAFT":          "MIS_DRAFT",
    "MIS_SUBMITTED":      "MIS_Submitted",
    "SUBMISSION_STATUS":  "SubmissionStatus",
    "RAILWAY_CLAIMS":     "Railway_Claims",
    "IRR_DETAILS":        "IRR_Details",
    "LEGAL_CASES":        "Legal_Cases",
    "REVISION_REQUESTS":    "RevisionRequests",
    "SETTINGS":             "Settings",
    # M&I MIS separate tabs
    "MI_TANK_OUTAGE":       "MI_TANK_OUTAGE",
    "MI_MAJOR_REPAIR":      "MI_MAJOR_REPAIR",
    "MI_VRU":               "MI_VRU",
    "MI_AUDIT_2526":        "MI_AUDIT_2526",
    "MI_AUDIT_2627":        "MI_AUDIT_2627",
    "MI_TECH_AUDIT":        "MI_TECH_AUDIT",
    "MI_EQUIP_BREAKDOWN":   "MI_EQUIP_BREAKDOWN",
    "MI_INT_PIPELINE":      "MI_INT_PIPELINE",
    "MI_EXT_PIPELINE":      "MI_EXT_PIPELINE",
    "MI_TANK_STATUS":       "MI_TANK_STATUS",
    "TANK_MASTER":          "TankMaster",
    "EMAIL_MASTER":         "EmailMaster",
}

# ── Header rows for auto-created tabs ────────────────────────────────────────

_SS_HEADERS = [
    "user_id", "month_year", "status", "completion_pct",
    "submitted_at", "locked_by", "locked_at", "checker_notes", "last_updated",
]

_RR_HEADERS = [
    "request_id", "zone_id", "location_id", "month_year",
    "reason", "status", "actioned_by", "actioned_at", "notes", "created_at",
]

_SETTINGS_HEADERS    = ["key", "value", "updated_by", "updated_at"]
_EMAIL_MASTER_HEADERS = ["type", "code", "name", "email", "cc"]

# Detail-table definitions: sheet_headers (Google Sheet columns) + data_keys (editable fields)
_DETAIL_DEF = {
    "RAILWAY_CLAIMS": {
        "sheet_headers": [
            "Sr#", "Zone", "Location", "Month-Year", "user_id",
            "Claim No.", "Year", "Amount (Rs)", "RR Nos.", "Ex", "To",
            "T/Wagon Nos.", "Product", "Qty.", "Rly.",
            "Pending Stage", "Status of Claim",
            "Last Hearing Date", "Next Hearing Date",
            "RCT Case Status as per Website", "Case Facts",
            "Rejection Reasons", "ShortComings/Discrepancies",
            "Strength of Case", "Recommendation",
        ],
        "data_keys": [
            "claim_no", "year", "amount", "rr_nos", "ex_station", "to_station",
            "wagon_nos", "product", "qty", "rly", "pending_stage", "status_claim",
            "last_hearing", "next_hearing", "rct_status", "case_facts",
            "rejection_reasons", "shortcomings", "strength", "recommendation",
        ],
        "prefix_count": 5,   # Sr#, Zone, Location, Month-Year, user_id
    },
    "IRR_DETAILS": {
        "sheet_headers": [
            "Sr#", "Zone", "Location Code", "Location Name", "Month-Year", "user_id",
            "IRR#", "IRR Date", "IRR Description", "IRR Amount (Rs)",
            "IRR Status (OPEN/CLOSED)", "IRR Closure Date",
        ],
        "data_keys": [
            "irr_no", "irr_date", "description", "amount", "status", "closure_date",
        ],
        "prefix_count": 6,   # Sr#, Zone, Location Code, Location Name, Month-Year, user_id
    },
    "LEGAL_CASES": {
        "sheet_headers": [
            "Sr. No", "Zone", "Location", "Month-Year", "user_id",
            "Court Name", "Case Number", "Cause Title", "Advocate",
            "Nature of Case", "Dealership Name and Location",
            "Background", "Status", "Last Hearing Date", "Next Hearing Date",
        ],
        "data_keys": [
            "court_name", "case_number", "cause_title", "advocate", "nature",
            "dealership", "background", "status", "last_hearing", "next_hearing",
        ],
        "prefix_count": 5,   # Sr. No, Zone, Location, Month-Year, user_id
    },
}

_MONTHS_S = ["Jan","Feb","Mar","Apr","May","Jun",
              "Jul","Aug","Sep","Oct","Nov","Dec"]
MONTHS_LONG = ["January","February","March","April","May","June",
               "July","August","September","October","November","December"]


# ── Sheets client (cached for the app lifetime) ──────────────────────────────

@st.cache_resource
def _client():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"], scopes=_SCOPES
    )
    return gspread.authorize(creds)


@st.cache_resource
def _spreadsheet():
    return _client().open_by_key(st.secrets["sheets"]["spreadsheet_id"])


@st.cache_resource
def _ws_cache() -> dict:
    """App-lifetime dict {tab_name: Worksheet}.
    Avoids repeated ss.worksheet() API calls (each call re-fetches sheet list)."""
    return {}


def _ws(tab_name: str):
    cache = _ws_cache()
    if tab_name not in cache:
        cache[tab_name] = _spreadsheet().worksheet(tab_name)
    return cache[tab_name]


def _ensure_ws(tab_name: str, headers: list = None, force_headers: bool = False):
    """Return worksheet, auto-creating with headers if missing.
    When force_headers=True, also overwrites row 1 if it differs from headers."""
    cache = _ws_cache()
    if tab_name not in cache:
        ss = _spreadsheet()
        try:
            cache[tab_name] = ss.worksheet(tab_name)
        except gspread.exceptions.WorksheetNotFound:
            cols = max(len(headers) if headers else 10, 26)
            ws   = ss.add_worksheet(title=tab_name, rows=2000, cols=cols)
            if headers:
                ws.append_row(headers, value_input_option="RAW")
            cache[tab_name] = ws
    ws = cache[tab_name]
    if force_headers and headers:
        existing = ws.get_all_values()
        if not existing or existing[0] != headers:
            ws.update("A1", [headers])
    return ws


def _field_label_map() -> dict:
    """Return {field_key: field_label} for all non-auto fields (from form_defs)."""
    try:
        from form_defs import SECTION_FIELDS
        return {
            f["key"]: f["label"]
            for fields in SECTION_FIELDS.values()
            for f in fields
            if not f.get("auto")
        }
    except Exception:
        return {}


# ── Month helpers ─────────────────────────────────────────────────────────────

def month_key(d: date = None) -> str:
    """date → 'Apr-2026'"""
    if d is None:
        d = date.today()
    return f"{_MONTHS_S[d.month - 1]}-{d.year}"


def parse_month_key(key: str):
    """'Apr-2026' → (month_0idx, year)"""
    parts = key.split("-")
    return _MONTHS_S.index(parts[0]), int(parts[1])


def compute_deadline(month_year: str) -> dict:
    """Return deadline date, days_left, urgency, and display strings for a month key."""
    month, year = parse_month_key(month_year)
    dl_month = month + 2          # +1 for 1-indexed, +1 for next month
    dl_year  = year
    if dl_month > 12:
        dl_month -= 12
        dl_year  += 1
    deadline  = date(dl_year, dl_month, 5)
    today     = date.today()
    days_left = (deadline - today).days

    if days_left < 0:
        urgency = "overdue"
    elif days_left <= 3:
        urgency = "urgent"
    elif days_left <= 7:
        urgency = "warning"
    else:
        urgency = "ok"

    return {
        "date":        f"{deadline.day}-{_MONTHS_S[deadline.month - 1]}-{deadline.year}",
        "days_left":   days_left,
        "urgency":     urgency,
        "month_label": MONTHS_LONG[month] + " " + str(year),
    }


def _prev_month_key() -> str:
    today = date.today()
    if today.month == 1:
        return month_key(date(today.year - 1, 12, 1))
    return month_key(date(today.year, today.month - 1, 1))


# ── Audit log (best-effort) ──────────────────────────────────────────────────

def audit_log(loc_code: str, action: str, details: str = ""):
    try:
        _ws(TABS["AUDIT_LOG"]).append_row(
            [datetime.now().isoformat(), loc_code, action, details],
            value_input_option="RAW",
        )
    except Exception:
        pass


# ── Location name lookup from master ─────────────────────────────────────────

@st.cache_data(ttl=3600)
def _loc_name_map() -> dict:
    """Return {location_code_upper: (location_name, loc_type, zone)} from LocationMaster.
    Col A = code, B = name, C = loc_type (HPCL|TOP|HMEL), D = zone (optional).
    loc_type is derived from name/code pattern if col C is blank."""
    try:
        ws   = _ws(TABS["LOCATION_MASTER"])
        rows = _api_call(ws.get_all_values)
        m    = {}
        for row in rows[1:]:
            if len(row) < 2:
                continue
            code = str(row[0]).strip()
            name = str(row[1]).strip()
            if not code:
                continue
            # Column C (index 2) = loc_type if present
            raw_type = str(row[2]).strip().upper() if len(row) > 2 else ""
            if raw_type in ("TOP", "HMEL", "HPCL"):
                ltype = raw_type
            else:
                nu = name.upper(); cu = code.upper()
                if "HMEL" in nu or "HMEL" in cu:
                    ltype = "HMEL"
                elif "TOP" in nu or "JAMNAGAR" in nu or "TOP" in cu:
                    ltype = "TOP"
                else:
                    ltype = "HPCL"
            # Column D (index 3) = zone if present
            zone = str(row[3]).strip() if len(row) > 3 else ""
            m[code.upper()] = (name, ltype, zone)
        return m
    except Exception:
        return {}


def get_loc_type(loc_code: str) -> str:
    """Return loc_type ('HPCL'|'TOP'|'HMEL') for a location code."""
    _, ltype = _resolve_loc_info(loc_code, loc_code)
    return ltype


def reset_location_data(loc_code: str) -> dict:
    """Delete ALL MIS data for a location (pre-launch data cleanup).

    Clears: MIS_DRAFT, SubmissionStatus, Railway_Claims, IRR_Details,
    Legal_Cases, MI_TANK_OUTAGE, MI_MAJOR_REPAIR, MI_VRU, MI_AUDIT_2526,
    MI_AUDIT_2627, MI_TECH_AUDIT, MI_EQUIP_BREAKDOWN, MI_INT_PIPELINE,
    MI_EXT_PIPELINE, MI_TANK_STATUS.
    """
    uid = str(loc_code).strip()
    tabs_to_clear = [
        "MIS_DRAFT", "SUBMISSION_STATUS",
        "RAILWAY_CLAIMS", "IRR_DETAILS", "LEGAL_CASES",
        "MI_TANK_OUTAGE", "MI_MAJOR_REPAIR", "MI_VRU",
        "MI_AUDIT_2526", "MI_AUDIT_2627", "MI_TECH_AUDIT",
        "MI_EQUIP_BREAKDOWN", "MI_INT_PIPELINE", "MI_EXT_PIPELINE",
        "MI_TANK_STATUS",
    ]
    deleted_total = 0
    errors = []
    try:
        for tab_key in tabs_to_clear:
            tab_name = TABS.get(tab_key)
            if not tab_name:
                continue
            try:
                ws = _ws(tab_name)
                rows_to_del = []
                for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
                    if row and str(row[0]).strip() == uid:
                        rows_to_del.append(i)
                for row_idx in reversed(rows_to_del):
                    ws.delete_rows(row_idx)
                deleted_total += len(rows_to_del)
            except Exception as tab_err:
                errors.append(f"{tab_name}: {tab_err}")
        # Also clear the cached dashboard data
        try:
            get_dashboard_data.clear()
        except Exception:
            pass
        msg = f"Deleted {deleted_total} row(s) for location {uid}."
        if errors:
            msg += "  Warnings: " + "; ".join(errors)
        return {"ok": True, "deleted": deleted_total, "msg": msg}
    except Exception as exc:
        return {"ok": False, "msg": str(exc)}


def _resolve_loc_info(loc_code: str, stored_name: str) -> tuple:
    """Return (loc_name, loc_type) for a location code."""
    code_up = loc_code.strip().upper()
    master  = _loc_name_map()
    if code_up in master:
        entry = master[code_up]
        name  = entry[0]
        ltype = entry[1]
        final_name = stored_name.strip() if (
            stored_name.strip() and stored_name.strip().upper() != code_up
        ) else name
        return final_name, ltype
    # Not in master: derive loc_type from name/code
    nu, cu = stored_name.upper(), code_up
    if "HMEL" in nu or "HMEL" in cu:
        ltype = "HMEL"
    elif "TOP" in nu or "JAMNAGAR" in nu or "TOP" in cu:
        ltype = "TOP"
    else:
        ltype = "HPCL"
    return stored_name.strip() or loc_code.strip(), ltype


def _resolve_loc_name(loc_code: str, stored_name: str) -> str:
    """Backward-compat wrapper — returns loc_name only."""
    return _resolve_loc_info(loc_code, stored_name)[0]


# ── Authentication ────────────────────────────────────────────────────────────

def check_login(location_code: str, password: str) -> dict:
    try:
        location_code = str(location_code or "").strip()
        password      = str(password      or "").strip()

        if not location_code or not password:
            return {"ok": False, "msg": "Location Code and Password are required."}

        ws   = _ws(TABS["USER_ACCESS"])
        rows = ws.get_all_values()

        if len(rows) < 2:
            return {"ok": False, "msg": "No users configured. Please contact Admin."}

        location_exists = False
        location_code_upper = location_code.upper()

        for sheet_row, row in enumerate(rows[1:], start=2):
            row = (row + [""] * 8)[:8]
            loc_code_cell, loc_name, zone, stored_pass, role, is_first_raw = row[:6]

            if loc_code_cell.strip().upper() != location_code_upper:
                continue

            location_exists = True

            if stored_pass.strip() != password:
                continue

            # Correct location + correct password → update LastLogin
            try:
                ws.update_cell(sheet_row, 7, datetime.now().isoformat())
            except Exception:
                pass

            audit_log(location_code, "Login", f"Successful login as {role.strip() or 'Maker'}")

            _lname, _ltype = _resolve_loc_info(loc_code_cell, loc_name)
            _is_default_pw = stored_pass.strip() == loc_code_cell.strip()
            return {
                "ok":          True,
                "userId":      loc_code_cell.strip(),
                "locName":     _lname,
                "locType":     _ltype,   # HPCL | TOP | HMEL
                "zone":        zone.strip(),
                "role":        role.strip() or "Maker",
                "isFirstLogin": is_first_raw.strip().upper() == "TRUE" or _is_default_pw,
                "_sheet_row":  sheet_row,
                "_password":   password,
            }

        if location_exists:
            return {"ok": False, "msg": "Incorrect password. Please try again."}
        return {"ok": False, "msg": "Location Code not found. Please check and try again."}

    except Exception as e:
        return {"ok": False, "msg": f"System error. Please try again. ({e})"}


def change_password(user_id: str, current_pass: str,
                    new_pass: str, confirm_pass: str) -> dict:
    try:
        user_id      = str(user_id      or "").strip()
        current_pass = str(current_pass or "").strip()
        new_pass     = str(new_pass     or "").strip()
        confirm_pass = str(confirm_pass or "").strip()

        if len(new_pass) < 6:
            return {"ok": False, "msg": "New password must be at least 6 characters."}
        if new_pass != confirm_pass:
            return {"ok": False, "msg": "New password and Confirm password do not match."}
        if new_pass == current_pass:
            return {"ok": False, "msg": "New password must be different from the current password."}

        ws   = _ws(TABS["USER_ACCESS"])
        rows = ws.get_all_values()

        if len(rows) < 2:
            return {"ok": False, "msg": "User record not found."}

        for sheet_row, row in enumerate(rows[1:], start=2):
            row = (row + [""] * 8)[:8]
            if row[0].strip() != user_id:
                continue
            if row[3].strip() != current_pass:
                return {"ok": False, "msg": "Current password is incorrect."}

            ws.update_cell(sheet_row, 4, new_pass)
            ws.update_cell(sheet_row, 6, "FALSE")
            ws.update_cell(sheet_row, 8, datetime.now().isoformat())

            audit_log(user_id, "Password Changed",
                      f"Password changed successfully. IsFirstLogin reset to FALSE. "
                      f"Role: {row[4].strip() or 'Maker'}. Location: {row[1].strip()}")
            return {"ok": True}

        return {"ok": False, "msg": "User record not found."}

    except Exception as e:
        return {"ok": False, "msg": f"System error: {e}"}


# ── Help desk / forgot password ──────────────────────────────────────────────

def request_password_reset(location_code: str) -> dict:
    try:
        location_code = str(location_code or "").strip()
        if not location_code:
            return {"ok": False, "msg": "Please enter your Location Code."}

        ws   = _ws(TABS["USER_ACCESS"])
        rows = ws.get_all_values()
        if len(rows) >= 2:
            codes = [r[0].strip() for r in rows[1:]]
            if location_code not in codes:
                return {
                    "ok":  False,
                    "msg": f'Location Code "{location_code}" is not registered.',
                }

        _ws(TABS["HELPDESK"]).append_row(
            [datetime.now().isoformat(), location_code,
             "Password Reset Request",
             "User requested a password reset via the Forgot Password link.",
             "Pending", ""],
            value_input_option="RAW",
        )
        audit_log(location_code, "Forgot Password", "Reset request logged")
        return {
            "ok":  True,
            "msg": "Your password reset request has been logged. "
                   "Admin will reset your password shortly.",
        }
    except Exception as e:
        return {"ok": False, "msg": f"System error: {e}"}


_HELPDESK_HEADERS = [
    "timestamp", "location_code", "issue_type",
    "issue_desc", "status", "admin_response", "responded_at",
]


@st.cache_data(ttl=120)
def get_helpdesk_tickets() -> list:
    """Return all helpdesk tickets (newest first).

    Each ticket dict includes a 'row' key = 1-based sheet row so
    respond_to_helpdesk_ticket() can update the correct row.
    """
    try:
        ws   = _ensure_ws(TABS["HELPDESK"], _HELPDESK_HEADERS)
        rows = _api_call(ws.get_all_values)
        out  = []
        for i, row in enumerate(rows[1:], start=2):
            row = (row + [""] * 7)[:7]
            out.append({
                "row":            i,
                "timestamp":      row[0],
                "location_code":  row[1],
                "issue_type":     row[2],
                "issue_desc":     row[3],
                "status":         row[4] or "Pending",
                "admin_response": row[5],
                "responded_at":   row[6],
            })
        return list(reversed(out))
    except Exception:
        return []


def respond_to_helpdesk_ticket(row: int, response: str,
                                status: str, updated_by: str) -> dict:
    """Write admin response + status + timestamp back to the sheet row."""
    try:
        ws  = _ensure_ws(TABS["HELPDESK"], _HELPDESK_HEADERS)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        ws.update(f"E{row}:G{row}", [[status, response, now]])
        get_helpdesk_tickets.clear()
        audit_log(updated_by, "HelpDesk Response",
                  f"Row {row} → {status}: {response[:80]}")
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


def log_help_request(location_code: str, issue_desc: str,
                     issue_type: str = "Help Request") -> dict:
    try:
        location_code = str(location_code or "").strip()
        issue_desc    = str(issue_desc    or "").strip()
        issue_type    = str(issue_type    or "Help Request").strip()

        if not location_code:
            return {"ok": False, "msg": "Please enter your Location Code."}
        if len(issue_desc) < 10:
            return {"ok": False, "msg": "Please describe your issue in at least 10 characters."}

        ref_id = "HD-" + str(int(time.time()))[-6:]

        _ws(TABS["HELPDESK"]).append_row(
            [datetime.now().isoformat(), location_code,
             issue_type, issue_desc, "Pending", ""],
            value_input_option="RAW",
        )
        audit_log(location_code, "Help Request", f"{ref_id} [{issue_type}] — {issue_desc[:60]}")
        return {
            "ok":  True,
            "ref": ref_id,
            "msg": f"Ticket raised (Ref: {ref_id}). Admin will respond shortly.",
        }
    except Exception as e:
        return {"ok": False, "msg": f"System error: {e}"}


# ── Submission status ─────────────────────────────────────────────────────────

@st.cache_data(ttl=120)
def _mis_submitted_keys() -> set:
    """Return set of (user_id, month_year) that physically exist in MIS_Submitted (120 s cache)."""
    try:
        ws   = _ws(TABS["MIS_SUBMITTED"])
        rows = _api_call(ws.get_all_values)
        if len(rows) < 2:
            return set()
        hdr   = rows[0]
        uid_c = next((hdr.index(h) for h in ("User ID", "user_id") if h in hdr), 0)
        mon_c = next((hdr.index(h) for h in ("Month-Year", "month_year") if h in hdr), 3)
        return {(r[uid_c].strip(), r[mon_c].strip()) for r in rows[1:] if len(r) > mon_c}
    except Exception:
        return set()


def _revert_if_deleted(user_id: str, month_year: str, status: str,
                       completion_pct: float) -> str:
    """If MIS_Submitted row was deleted externally, revert SubmissionStatus → IN_PROGRESS."""
    if status not in ("SUBMITTED", "LOCKED"):
        return status
    if (user_id, month_year) not in _mis_submitted_keys():
        _update_submission_status(user_id, month_year, "IN_PROGRESS", completion_pct)
        _mis_submitted_keys.clear()   # invalidate cache so next read sees new state
        return "IN_PROGRESS"
    return status


@st.cache_data(ttl=60)
def get_month_status(user_id: str, month_year: str) -> dict:
    try:
        ws   = _ws(TABS["SUBMISSION_STATUS"])
        rows = _api_call(ws.get_all_values)
        if len(rows) >= 2:
            for row in rows[1:]:
                row = (row + [""] * 9)[:9]
                if row[0].strip() == user_id and row[1].strip() == month_year:
                    pct    = float(row[3]) if row[3] else 0.0
                    status = _revert_if_deleted(
                        user_id, month_year, row[2].strip() or "NOT_STARTED", pct
                    )
                    return {
                        "status":         status,
                        "completion_pct": pct,
                        "is_locked":      status in ("SUBMITTED", "LOCKED", "PENDING_REVIEW"),
                        "checker_notes":  row[7].strip() if len(row) > 7 else "",
                    }
    except Exception:
        pass
    return {"status": "NOT_STARTED", "completion_pct": 0.0, "is_locked": False}


def get_fy_months(user_id: str, fy_start_year: int) -> dict:
    """Return all 12 months of a financial year (Apr→Mar) with submission status."""
    try:
        user_id = str(user_id or "").strip()

        status_map: dict = {}
        pct_map:    dict = {}
        try:
            ws   = _ws(TABS["SUBMISSION_STATUS"])
            rows = _api_call(ws.get_all_values)
            if len(rows) >= 2:
                for row in rows[1:]:
                    row = (row + [""] * 4)[:4]
                    if row[0].strip() == user_id:
                        status_map[row[1].strip()] = row[2].strip()
                        pct_map[row[1].strip()]    = float(row[3]) if row[3] else 0.0
        except Exception:
            pass

        months = []
        for i in range(12):
            # i=0→April, i=1→May, …, i=8→December, i=9→January, …, i=11→March
            month_num = (i + 3) % 12 + 1
            yr        = fy_start_year if i < 9 else fy_start_year + 1
            d         = date(yr, month_num, 1)
            key       = month_key(d)
            label     = MONTHS_LONG[d.month - 1] + " " + str(d.year)
            status    = _revert_if_deleted(
                user_id, key, status_map.get(key, "NOT_STARTED"), pct_map.get(key, 0.0)
            )
            months.append({
                "value":     key,
                "label":     label,
                "status":    status,
                "is_locked": status in ("SUBMITTED", "LOCKED", "PENDING_REVIEW"),
            })

        return {"ok": True, "months": months}

    except Exception as e:
        return {"ok": False, "msg": str(e), "months": []}


def get_available_months(user_id: str) -> dict:
    try:
        user_id = str(user_id or "").strip()
        today   = date.today()

        # Fetch submission status — if tab missing just use empty map
        status_map: dict = {}
        try:
            ws   = _ws(TABS["SUBMISSION_STATUS"])
            rows = _api_call(ws.get_all_values)
            if len(rows) >= 2:
                for row in rows[1:]:
                    row = (row + [""] * 3)[:3]
                    if row[0].strip() == user_id:
                        status_map[row[1].strip()] = row[2].strip()
        except Exception:
            pass

        months = []
        for m in range(1, 13):
            month_0 = today.month - 1 - m   # 0-indexed month
            yr      = today.year
            while month_0 < 0:
                month_0 += 12
                yr      -= 1
            d      = date(yr, month_0 + 1, 1)
            key    = month_key(d)
            label  = MONTHS_LONG[d.month - 1] + " " + str(d.year)
            status = status_map.get(key, "NOT_STARTED")
            months.append({
                "value":     key,
                "label":     label,
                "status":    status,
                "is_locked": status in ("SUBMITTED", "LOCKED", "PENDING_REVIEW"),
            })

        return {"ok": True, "months": months}

    except Exception as e:
        return {"ok": False, "msg": str(e), "months": []}


@st.cache_data(ttl=60, show_spinner=False)
def get_dashboard_data(user_id: str, month_year: str = None, loc_type: str = "HPCL") -> dict:
    try:
        user_id = str(user_id or "").strip()
        if not user_id:
            return {"ok": False, "msg": "Session expired. Please log in again."}

        if not month_year:
            month_year = _prev_month_key()

        status_data = get_month_status(user_id, month_year)
        deadline    = compute_deadline(month_year)

        # Derive per-section completion from the draft row (column C)
        secs_done: list = []
        draft: dict = {}
        try:
            draft    = load_draft(user_id, month_year)
            secs_raw = draft.get("_sections_complete", "")
            secs_done = sorted(
                {int(x) for x in secs_raw.split(",") if x.strip().isdigit()}
            )
        except Exception:
            pass

        # Validate stored completion against current SECTION_FIELDS definitions.
        # A section is only complete if every current required non-auto non-excluded
        # field has a non-empty value in the saved draft.
        try:
            from form_defs import SECTION_FIELDS, get_excluded_fields, get_skip_sections
            excl_keys  = get_excluded_fields(loc_type)
            skip_secs  = get_skip_sections(loc_type)
            valid_secs = []
            for sec_num in secs_done:
                if sec_num in skip_secs:
                    valid_secs.append(sec_num)  # skip-sections always count as done
                    continue
                all_filled = True
                for f in SECTION_FIELDS.get(sec_num, []):
                    if f.get("auto") or not f.get("req"):
                        continue
                    if f["key"] in excl_keys:
                        continue
                    # Conditional field — skip if its show_when condition is not met
                    sw = f.get("show_when")
                    if sw and not all(
                        str(draft.get(k) or "") == str(v) for k, v in sw.items()
                    ):
                        continue
                    if draft.get(f["key"]) in (None, ""):
                        all_filled = False
                        break
                if all_filled:
                    valid_secs.append(sec_num)
            secs_done = sorted(valid_secs)
        except Exception:
            pass  # on any import/logic error keep original secs_done

        # S5 only counts as complete when M&I MIS (S5A) is also fully filled
        mi_complete = check_mi_complete(user_id, month_year) if 5 in secs_done else False
        if 5 in secs_done and not mi_complete:
            secs_done = [s for s in secs_done if s != 5]

        eff_pct = len(secs_done) * 10.0

        return {
            "ok":             True,
            "month_year":     month_year,
            "status":         status_data["status"],
            "completion_pct": eff_pct,
            "is_locked":      status_data["is_locked"],
            "checker_notes":  status_data.get("checker_notes", ""),
            "secs_done":      secs_done,
            "mi_complete":    mi_complete,
            "deadline":       deadline,
        }
    except Exception as e:
        return {"ok": False, "msg": f"Error loading dashboard: {e}"}


# ── Draft CRUD ────────────────────────────────────────────────────────────────
# MIS_Draft columns: user_id | month_year | sections_complete | last_updated | data_json

def load_draft(user_id: str, month_year: str) -> dict:
    """Return draft data dict for user+month, or empty dict."""
    try:
        ws   = _ws(TABS["MIS_DRAFT"])
        rows = _api_call(ws.get_all_values)
        for i, row in enumerate(rows[1:], start=2):
            row = (row + [""] * 5)[:5]
            if row[0].strip() == user_id and row[1].strip() == month_year:
                try:
                    data = json.loads(row[4]) if row[4].strip() else {}
                except Exception:
                    data = {}
                data["_sections_complete"] = row[2].strip()
                data["_sheet_row"]         = i
                return data
    except Exception:
        pass
    return {}


def _update_submission_status(user_id: str, month_year: str, status: str, pct: float,
                               submitted_at: str = "", locked_by: str = "",
                               locked_at: str = "", checker_notes: str = ""):
    # _ensure_ws auto-creates the tab with headers if it doesn't exist yet.
    # Any exception propagates to save_draft which surfaces it as an error message.
    ws      = _ensure_ws(TABS["SUBMISSION_STATUS"], _SS_HEADERS)
    rows    = _api_call(ws.get_all_values)
    now_str = datetime.now().isoformat()
    for i, row in enumerate(rows[1:], start=2):
        row_e = (row + [""] * 9)[:9]
        if row_e[0].strip() == user_id and row_e[1].strip() == month_year:
            _api_call(ws.update, f"A{i}:I{i}", [[
                user_id, month_year, status, str(pct),
                submitted_at or row_e[4] or now_str,
                locked_by  or row_e[5],
                locked_at  or row_e[6],
                checker_notes if checker_notes else row_e[7],
                now_str,
            ]])
            get_month_status.clear()
            return
    _api_call(ws.append_row,
        [user_id, month_year, status, str(pct),
         submitted_at or now_str, locked_by, locked_at, checker_notes, now_str],
        value_input_option="RAW",
    )
    get_month_status.clear()


def save_draft(user_id: str, month_year: str,
               section_num: int | None = None,
               field_data: dict | None = None,
               mark_complete: bool = False,
               sections_complete: list | None = None) -> dict:
    """Merge field_data into existing draft and persist to MIS_Draft tab.

    Two calling modes:
      Normal (per-section save):
        save_draft(user_id, month_year, section_num, field_data, mark_complete)
      Bulk upload save:
        save_draft(user_id, month_year, field_data={...}, sections_complete=[1,3,5,...])
    """
    try:
        existing  = load_draft(user_id, month_year)
        sheet_row = existing.pop("_sheet_row", None)
        secs_raw  = existing.pop("_sections_complete", "")

        try:
            secs_done = {int(x) for x in secs_raw.split(",") if x.strip().isdigit()}
        except Exception:
            secs_done = set()

        if field_data:
            existing.update(field_data)

        if sections_complete is not None:
            # Bulk upload: replace section completion entirely
            secs_done = set(sections_complete)
        elif section_num is not None:
            if mark_complete:
                secs_done.add(section_num)
            else:
                secs_done.discard(section_num)

        pct      = len(secs_done) * 10.0
        secs_str = ",".join(str(s) for s in sorted(secs_done))
        now_str  = datetime.now().isoformat()
        data_str = json.dumps(existing)

        ws = _ws(TABS["MIS_DRAFT"])
        if sheet_row:
            _api_call(ws.update, f"A{sheet_row}:E{sheet_row}",
                      [[user_id, month_year, secs_str, now_str, data_str]])
        else:
            _api_call(ws.append_row,
                [user_id, month_year, secs_str, now_str, data_str],
                value_input_option="RAW",
            )

        status = "IN_PROGRESS" if secs_done else "NOT_STARTED"
        _update_submission_status(user_id, month_year, status, pct)
        tag = f"S{section_num}" if section_num else "BulkUpload"
        audit_log(user_id, f"SaveDraft {tag}",
                  f"month={month_year} secs={secs_str} pct={pct}")

        # Invalidate dashboard cache so this user's next dashboard load is fresh
        get_dashboard_data.clear()
        return {"ok": True, "pct": pct, "secs_done": sorted(secs_done)}

    except Exception as e:
        return {"ok": False, "msg": str(e)}


# ── Detail table CRUD (Railway Claims, IRR Details, Legal Cases) ─────────────

def load_detail_table(user_id: str, month_year: str, tab_key: str) -> list:
    """Return list of row-dicts for user+month from a detail sheet."""
    try:
        defn      = _DETAIL_DEF[tab_key]
        ws        = _ensure_ws(TABS[tab_key], defn["sheet_headers"])
        all_rows  = ws.get_all_values()
        if len(all_rows) < 2:
            return []
        headers   = all_rows[0]
        try:
            uid_idx = headers.index("user_id")
            mon_idx = headers.index("Month-Year")
        except ValueError:
            return []
        prefix  = defn["prefix_count"]
        dkeys   = defn["data_keys"]
        result  = []
        for row in all_rows[1:]:
            row = (row + [""] * len(headers))[:len(headers)]
            if row[uid_idx].strip() != user_id or row[mon_idx].strip() != month_year:
                continue
            result.append({k: (row[prefix + i] if prefix + i < len(row) else "")
                           for i, k in enumerate(dkeys)})
        return result
    except Exception:
        return []


def save_detail_table(user_id: str, month_year: str, tab_key: str,
                      rows_data: list, user_info: dict) -> dict:
    """Replace all rows for user+month in a detail sheet with rows_data."""
    try:
        defn      = _DETAIL_DEF[tab_key]
        ws        = _ensure_ws(TABS[tab_key], defn["sheet_headers"])
        all_rows  = ws.get_all_values()
        headers   = all_rows[0] if all_rows else defn["sheet_headers"]
        try:
            uid_idx = headers.index("user_id")
            mon_idx = headers.index("Month-Year")
        except ValueError:
            uid_idx, mon_idx = 4, 3

        # Delete existing rows for this user+month (reverse order to keep indices stable)
        to_del = [i + 2 for i, row in enumerate(all_rows[1:])
                  if (row + [""] * len(headers))[uid_idx].strip() == user_id
                  and (row + [""] * len(headers))[mon_idx].strip() == month_year]
        for idx in reversed(to_del):
            ws.delete_rows(idx)

        zone = user_info.get("zone", "")
        loc  = user_info.get("locName", "")

        for sr, rec in enumerate(rows_data, 1):
            if tab_key == "IRR_DETAILS":
                prefix = [sr, zone, user_id, loc, month_year, user_id]
            else:
                prefix = [sr, zone, loc, month_year, user_id]
            data_vals = [str(rec.get(k, "") or "") for k in defn["data_keys"]]
            ws.append_row(prefix + data_vals, value_input_option="RAW")

        audit_log(user_id, f"SaveDetail {tab_key}",
                  f"month={month_year} rows={len(rows_data)}")
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


# ── Maker-Checker Submit / Approve / Reject workflow ─────────────────────────

def submit_for_review(user_id: str, month_year: str) -> dict:
    """Maker submits completed draft for Checker review."""
    try:
        sd  = get_month_status(user_id, month_year)
        pct = sd.get("completion_pct", 0)
        if pct < 100:
            return {"ok": False,
                    "msg": (f"Cannot submit — completion is {int(pct)}%. "
                            "All 10 sections must be saved first.")}
        if not check_mi_complete(user_id, month_year):
            return {"ok": False,
                    "msg": ("Cannot submit — M&I MIS (S5A) is incomplete. "
                            "Please fill all 10 tabs in S5A and save each one.")}
        if sd["status"] in ("SUBMITTED", "LOCKED", "PENDING_REVIEW"):
            return {"ok": False, "msg": f"Month is already {sd['status'].replace('_',' ').lower()}."}
        _update_submission_status(user_id, month_year, "PENDING_REVIEW", pct,
                                  submitted_at=datetime.now().isoformat())
        audit_log(user_id, "SubmitForReview", f"month={month_year}")
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


def approve_submission(maker_id: str, month_year: str, checker_id: str,
                       flat_data: dict, user_info: dict) -> dict:
    """Checker approves — writes flat row to MIS_Submitted and locks the month."""
    try:
        label_map  = _field_label_map()
        field_keys = list(flat_data.keys())
        col_labels = [label_map.get(k, k) for k in field_keys]
        headers    = (["User ID", "Location Name", "Zone", "Month-Year",
                        "Submitted At", "Approved At", "Approved By"]
                      + col_labels)
        ws      = _ensure_ws(TABS["MIS_SUBMITTED"], headers, force_headers=True)
        all_r   = ws.get_all_values()
        now_str = datetime.now().isoformat()

        row = ([maker_id, user_info.get("locName", ""), user_info.get("zone", ""),
                month_year, now_str, now_str, checker_id]
               + [str(flat_data.get(k, "") or "") for k in field_keys])

        if len(all_r) >= 2:
            hdr    = all_r[0]
            uid_c  = next((hdr.index(h) for h in ("User ID", "user_id") if h in hdr), 0)
            mon_c  = next((hdr.index(h) for h in ("Month-Year", "month_year") if h in hdr), 3)
            for i, r in enumerate(all_r[1:], start=2):
                re = (r + [""] * max(len(hdr), 1))[:max(len(hdr), 1)]
                if re[uid_c].strip() == maker_id and re[mon_c].strip() == month_year:
                    ws.update(f"A{i}", [row])
                    break
            else:
                ws.append_row(row, value_input_option="RAW")
        else:
            ws.append_row(row, value_input_option="RAW")

        sd = get_month_status(maker_id, month_year)
        _update_submission_status(maker_id, month_year, "SUBMITTED",
                                  sd.get("completion_pct", 100),
                                  locked_by=checker_id, locked_at=now_str)
        audit_log(checker_id, "ApproveSubmission", f"maker={maker_id} month={month_year}")
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


def reject_submission(maker_id: str, month_year: str,
                      checker_id: str, note: str) -> dict:
    """Checker rejects — returns to editable state with a note for the Maker."""
    try:
        sd = get_month_status(maker_id, month_year)
        if sd["status"] != "PENDING_REVIEW":
            return {"ok": False, "msg": "Only PENDING_REVIEW submissions can be rejected."}
        _update_submission_status(maker_id, month_year, "REJECTED",
                                  sd.get("completion_pct", 0),
                                  checker_notes=note)
        audit_log(checker_id, "RejectSubmission",
                  f"maker={maker_id} month={month_year} note={note[:60]}")
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


def reset_draft(maker_id: str, month_year: str, checker_id: str, reason: str) -> dict:
    """Checker resets a maker's draft — wipes all field data so maker can start fresh."""
    try:
        ws       = _ws(TABS["MIS_DRAFT"])
        all_rows = ws.get_all_values()
        now_str  = datetime.now().isoformat()
        for i, row in enumerate(all_rows[1:], start=2):
            r = (row + [""] * 5)[:5]
            if r[0].strip() == maker_id and r[1].strip() == month_year:
                ws.update(f"A{i}:E{i}",
                          [[maker_id, month_year, "", now_str, "{}"]])
                break
        _update_submission_status(
            maker_id, month_year, "NOT_STARTED", 0.0,
            checker_notes=f"[RESET by {checker_id}] {reason}",
        )
        audit_log(checker_id, "ResetDraft",
                  f"maker={maker_id} month={month_year} reason={reason[:80]}")
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


# ── Phase-6: Zone / HQO view + Revision workflow ──────────────────────────────

def get_locations_by_zone(zone_name: str) -> list:
    """Return Maker location dicts for a given zone (reads UserAccess)."""
    try:
        ws   = _ws(TABS["USER_ACCESS"])
        rows = ws.get_all_values()
        locs = []
        for row in rows[1:]:
            row = (row + [""] * 8)[:8]
            loc_code, loc_name, zone, _, role = row[:5]
            if zone.strip() == zone_name and role.strip() == "Maker":
                locs.append({
                    "userId":  loc_code.strip(),
                    "locName": loc_name.strip(),
                    "zone":    zone.strip(),
                })
        return locs
    except Exception:
        return []


def get_all_maker_locations() -> list:
    """Return all Maker location dicts across all zones (HQO view)."""
    try:
        ws   = _ws(TABS["USER_ACCESS"])
        rows = ws.get_all_values()
        locs = []
        for row in rows[1:]:
            row = (row + [""] * 8)[:8]
            loc_code, loc_name, zone, _, role = row[:5]
            if role.strip() == "Maker":
                locs.append({
                    "userId":  loc_code.strip(),
                    "locName": loc_name.strip(),
                    "zone":    zone.strip(),
                })
        return locs
    except Exception:
        return []


def get_maker_info(user_id: str) -> dict:
    """Return {"userId", "locName", "zone"} for a Maker, or a minimal dict."""
    try:
        ws   = _ws(TABS["USER_ACCESS"])
        rows = ws.get_all_values()
        for row in rows[1:]:
            row = (row + [""] * 8)[:8]
            loc_code, loc_name, zone, _, role = row[:5]
            if loc_code.strip() == user_id and role.strip() == "Maker":
                return {"userId": loc_code.strip(), "locName": loc_name.strip(),
                        "zone": zone.strip()}
    except Exception:
        pass
    return {"userId": user_id, "locName": user_id, "zone": ""}


def get_submissions_for_locations(locs: list, month_year: str) -> list:
    """Bulk-fetch submission status for a list of location dicts."""
    status_map: dict = {}
    try:
        ws   = _ws(TABS["SUBMISSION_STATUS"])
        rows = ws.get_all_values()
        for row in rows[1:]:
            row = (row + [""] * 9)[:9]
            if row[1].strip() == month_year:
                uid = row[0].strip()
                pct = float(row[3]) if row[3] else 0.0
                status_map[uid] = {
                    "status":         _revert_if_deleted(uid, month_year,
                                                         row[2].strip() or "NOT_STARTED", pct),
                    "completion_pct": pct,
                }
    except Exception:
        pass

    results = []
    for loc in locs:
        sd = status_map.get(loc["userId"],
                            {"status": "NOT_STARTED", "completion_pct": 0.0})
        results.append({**loc, **sd})
    return results


def create_revision_request(zone_id: str, location_id: str,
                            month_year: str, reason: str) -> dict:
    """Zone user raises a correction request for an already-submitted month."""
    try:
        if not reason.strip():
            return {"ok": False, "msg": "Please provide a reason for the revision request."}

        ws   = _ensure_ws(TABS["REVISION_REQUESTS"], _RR_HEADERS)
        rows = ws.get_all_values()
        for row in rows[1:]:
            row = (row + [""] * 10)[:10]
            if (row[2].strip() == location_id and
                    row[3].strip() == month_year and
                    row[5].strip() == "PENDING_HQO"):
                return {"ok": False,
                        "msg": "A revision request for this location/month is already pending HQO approval."}

        import uuid
        req_id  = str(uuid.uuid4())[:8].upper()
        now_str = datetime.now().isoformat()
        ws.append_row(
            [req_id, zone_id, location_id, month_year,
             reason, "PENDING_HQO", "", "", "", now_str],
            value_input_option="RAW",
        )
        audit_log(zone_id, "RevisionRequest",
                  f"req={req_id} loc={location_id} month={month_year}")
        return {"ok": True, "msg": f"Revision request {req_id} submitted to HQO."}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


def get_revision_requests(zone_filter: str = "") -> list:
    """Return revision requests; optionally filtered to a specific zone_id."""
    try:
        ws   = _ensure_ws(TABS["REVISION_REQUESTS"], _RR_HEADERS)
        rows = ws.get_all_values()
        result = []
        for i, row in enumerate(rows[1:], start=2):
            row = (row + [""] * 10)[:10]
            if zone_filter and row[1].strip() != zone_filter:
                continue
            result.append({
                "row":         i,
                "request_id":  row[0].strip(),
                "zone_id":     row[1].strip(),
                "location_id": row[2].strip(),
                "month_year":  row[3].strip(),
                "reason":      row[4].strip(),
                "status":      row[5].strip() or "PENDING_HQO",
                "actioned_by": row[6].strip(),
                "actioned_at": row[7].strip(),
                "notes":       row[8].strip(),
                "created_at":  row[9].strip(),
            })
        return result
    except Exception:
        return []


def approve_revision_request(request_id: str, actioned_by: str) -> dict:
    """HQO approves revision — unlocks the location's month for re-editing."""
    try:
        ws   = _ensure_ws(TABS["REVISION_REQUESTS"], _RR_HEADERS)
        rows = ws.get_all_values()
        for i, row in enumerate(rows[1:], start=2):
            row = (row + [""] * 10)[:10]
            if row[0].strip() == request_id:
                location_id = row[2].strip()
                month_year  = row[3].strip()
                now_str     = datetime.now().isoformat()
                ws.update(f"F{i}:I{i}",
                          [["APPROVED", actioned_by, now_str, ""]])
                sd = get_month_status(location_id, month_year)
                _update_submission_status(
                    location_id, month_year, "REJECTED",
                    sd.get("completion_pct", 0),
                    checker_notes=(
                        "Correction approved by HQO. "
                        "Please update the data and resubmit."
                    ),
                )
                audit_log(actioned_by, "ApproveRevision",
                          f"req={request_id} loc={location_id} month={month_year}")
                return {"ok": True}
        return {"ok": False, "msg": f"Request ID {request_id} not found."}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


def reject_revision_request(request_id: str,
                             actioned_by: str, note: str) -> dict:
    """HQO rejects revision request."""
    try:
        ws   = _ensure_ws(TABS["REVISION_REQUESTS"], _RR_HEADERS)
        rows = ws.get_all_values()
        for i, row in enumerate(rows[1:], start=2):
            row = (row + [""] * 10)[:10]
            if row[0].strip() == request_id:
                now_str = datetime.now().isoformat()
                ws.update(f"F{i}:I{i}",
                          [["REJECTED", actioned_by, now_str, note]])
                audit_log(actioned_by, "RejectRevision",
                          f"req={request_id} note={note[:60]}")
                return {"ok": True}
        return {"ok": False, "msg": f"Request ID {request_id} not found."}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


# ── Fix-1: Auto-populate Zone & HQO accounts ──────────────────────────────────

_UA_HEADERS = [
    "user_id", "location_name", "zone", "password", "role",
    "email", "mobile", "active",
]

def setup_zone_accounts() -> dict:
    """Auto-generate Zone and HQO accounts in UserAccess from existing Maker rows.

    Zone account convention:
      user_id  = first-3-chars-of-zone-name (uppercase) + "ZONE"  e.g. "BHOZONE"
      password = first-3-chars-of-zone-name (uppercase) + "MIS"   e.g. "BHOMIS"
      role     = "Zone"

    HQO account:
      user_id = "SODSBU", password = "SODMIS", role = "Admin"

    Skips rows that already exist (matched on user_id).
    Returns {"ok": True, "added": [...newly added user_ids...]}
    """
    try:
        ws   = _ensure_ws(TABS["USER_ACCESS"], _UA_HEADERS)
        rows = ws.get_all_values()
        if not rows:
            return {"ok": False, "msg": "UserAccess sheet is empty."}

        hdr = rows[0]
        existing_ids = {
            (r + [""] * len(hdr))[0].strip().upper()
            for r in rows[1:]
        }

        # Collect unique zone names from Maker rows
        zones_seen: dict[str, str] = {}   # zone_name -> derived zone_id
        for r in rows[1:]:
            r = (r + [""] * 8)[:8]
            zone_name = r[2].strip()
            role      = r[4].strip()
            if role == "Maker" and zone_name:
                prefix   = zone_name[:3].upper()
                zone_uid = prefix + "ZONE"
                zones_seen[zone_name] = zone_uid

        now_str = datetime.now().isoformat()
        added   = []

        for zone_name, zone_uid in zones_seen.items():
            if zone_uid.upper() in existing_ids:
                continue
            prefix = zone_uid[:3]
            pw     = prefix + "MIS"
            ws.append_row(
                [zone_uid, zone_name + " Zone", zone_name, pw, "Zone",
                 "", "", "Y"],
                value_input_option="RAW",
            )
            added.append(zone_uid)
            audit_log("SYSTEM", "SetupZoneAccount",
                      f"Added zone account {zone_uid} for zone {zone_name}")

        # HQO / Admin account
        if "SODSBU" not in existing_ids:
            ws.append_row(
                ["SODSBU", "HQ Operations", "ALL", "SODMIS", "Admin",
                 "", "", "Y"],
                value_input_option="RAW",
            )
            added.append("SODSBU")
            audit_log("SYSTEM", "SetupZoneAccount", "Added HQO account SODSBU")

        # View-only account
        if "SODVIEW" not in existing_ids:
            ws.append_row(
                ["SODVIEW", "View Only Access", "ALL", "VIEWMIS", "Viewer",
                 "", "", "Y"],
                value_input_option="RAW",
            )
            added.append("SODVIEW")
            audit_log("SYSTEM", "SetupZoneAccount", "Added Viewer account SODVIEW")

        try:
            _spreadsheet.clear()   # force next login to re-read UserAccess
        except Exception:
            pass
        return {"ok": True, "added": added}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


def get_zone_admin_accounts() -> list:
    """Return all Zone and Admin rows from UserAccess for diagnostics."""
    try:
        ws   = _ws(TABS["USER_ACCESS"])
        rows = ws.get_all_values()
        out  = []
        for r in rows[1:]:
            r = (r + [""] * 8)[:8]
            role = r[4].strip()
            if role in ("Zone", "Admin"):
                out.append({
                    "user_id":  r[0].strip(),
                    "loc_name": r[1].strip(),
                    "zone":     r[2].strip(),
                    "password": r[3].strip(),
                    "role":     role,
                })
        return out
    except Exception:
        return []


def get_all_maker_credentials() -> list:
    """Return list of {userId, locName, zone, password} for all Maker accounts."""
    try:
        ws   = _ws(TABS["USER_ACCESS"])
        rows = _api_call(ws.get_all_values)
        out  = []
        for r in rows[1:]:
            r = (r + [""] * 8)[:8]
            if r[4].strip() == "Maker" and r[0].strip():
                out.append({
                    "userId":  r[0].strip(),
                    "locName": r[1].strip(),
                    "zone":    r[2].strip(),
                    "password": r[3].strip(),
                })
        return out
    except Exception:
        return []


def get_all_checker_credentials() -> list:
    """Return list of {userId, locName, zone, password} for all Checker accounts."""
    try:
        ws   = _ws(TABS["USER_ACCESS"])
        rows = _api_call(ws.get_all_values)
        out  = []
        for r in rows[1:]:
            r = (r + [""] * 8)[:8]
            if r[4].strip() == "Checker" and r[0].strip():
                out.append({
                    "userId":  r[0].strip(),
                    "locName": r[1].strip(),
                    "zone":    r[2].strip(),
                    "password": r[3].strip(),
                })
        return out
    except Exception:
        return []


def get_all_zone_credentials() -> list:
    """Return list of {userId, locName, zone, password} for all Zone accounts.
    Deduplicates by zone name — keeps the first occurrence to prevent double emails.
    """
    try:
        ws   = _ws(TABS["USER_ACCESS"])
        rows = _api_call(ws.get_all_values)
        out  = []
        seen_zones = set()
        for r in rows[1:]:
            r = (r + [""] * 8)[:8]
            zone = r[2].strip()
            if r[4].strip() == "Zone" and r[0].strip():
                if zone in seen_zones:
                    continue  # skip duplicate zone row
                seen_zones.add(zone)
                out.append({
                    "userId":   r[0].strip(),
                    "locName":  r[1].strip(),
                    "zone":     zone,
                    "password": r[3].strip(),
                })
        return out
    except Exception:
        return []


def upsert_zone_account(zone_name: str, new_user_id: str, new_password: str) -> dict:
    """Create or update a Zone account matched by zone_name.

    - If a row with role=Zone and zone=zone_name exists → updates user_id + password.
    - If no such row exists → appends a new row.
    Always returns {"ok": True/False, "action": "updated"/"created"}.
    """
    try:
        ws   = _ws(TABS["USER_ACCESS"])
        rows = ws.get_all_values()

        for i, row in enumerate(rows[1:], start=2):
            row = (row + [""] * 8)[:8]
            stored_zone = row[2].strip()
            stored_role = row[4].strip()
            if stored_role == "Zone" and stored_zone == zone_name:
                # Update user_id (col A) and password (col D) in-place
                ws.update(f"A{i}", [[new_user_id]])
                ws.update(f"D{i}", [[new_password]])
                ws.update(f"B{i}", [[zone_name]])   # normalise loc_name too
                audit_log("SYSTEM", "UpdateZoneAccount",
                          f"zone={zone_name} new_id={new_user_id}")
                return {"ok": True, "action": "updated"}

        # Not found → create fresh row
        ws.append_row(
            [new_user_id, zone_name, zone_name, new_password, "Zone", "", "", "Y"],
            value_input_option="RAW",
        )
        audit_log("SYSTEM", "CreateZoneAccount",
                  f"zone={zone_name} id={new_user_id}")
        return {"ok": True, "action": "created"}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


def sync_missing_maker_accounts(default_password: str = "") -> dict:
    """Add Maker and Checker UserAccess rows for every LocationMaster location
    that is missing either role.

    Reads LocationMaster (col A=code, B=name, C=loc_type, D=zone).
    Checks separately for Maker and Checker rows — a location already having
    a Maker account will still get a Checker row added (and vice-versa).
    Default passwords: Maker = loc_code, Checker = loc_code + "C".
    Sets is_first = TRUE so users change password on first login.
    Returns {"ok": True, "added": [...descriptions...], "skipped": count}.
    """
    try:
        # Read LocationMaster
        lm_ws   = _ws(TABS["LOCATION_MASTER"])
        lm_rows = _api_call(lm_ws.get_all_values)

        # Build (code_upper, role) set of what already exists
        ua_ws   = _ensure_ws(TABS["USER_ACCESS"], _UA_HEADERS)
        ua_rows = _api_call(ua_ws.get_all_values)
        existing = set()   # {(code_upper, role_upper)}
        for r in ua_rows[1:]:
            if not r or not str(r[0]).strip():
                continue
            c = str(r[0]).strip().upper()
            role_r = str(r[4]).strip().upper() if len(r) > 4 else "MAKER"
            existing.add((c, role_r))

        added   = []
        skipped = 0
        now_str = datetime.now().isoformat()

        for row in lm_rows[1:]:
            if not row or len(row) < 2:
                continue
            code = str(row[0]).strip()
            name = str(row[1]).strip()
            if not code:
                continue
            zone = str(row[3]).strip() if len(row) > 3 else ""
            code_up = code.upper()
            base_pw = str(default_password).strip() or code

            # Add Maker if missing
            if (code_up, "MAKER") not in existing:
                _api_call(
                    ua_ws.append_row,
                    [code, name, zone, base_pw, "Maker", "TRUE", now_str, "Y"],
                    value_input_option="RAW",
                )
                existing.add((code_up, "MAKER"))
                added.append(f"{code} (Maker)")
                audit_log("SYSTEM", "SyncAccount",
                          f"Added Maker for {code} ({name})")
            else:
                skipped += 1

            # Add Checker if missing
            if (code_up, "CHECKER") not in existing:
                _api_call(
                    ua_ws.append_row,
                    [code, name, zone, base_pw + "C", "Checker", "TRUE", now_str, "Y"],
                    value_input_option="RAW",
                )
                existing.add((code_up, "CHECKER"))
                added.append(f"{code} (Checker)")
                audit_log("SYSTEM", "SyncAccount",
                          f"Added Checker for {code} ({name})")
            else:
                skipped += 1

        return {"ok": True, "added": added, "skipped": skipped}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


def hqo_account_exists() -> bool:
    """Quick check: is there at least one Admin-role row in UserAccess?"""
    try:
        ws   = _ws(TABS["USER_ACCESS"])
        rows = ws.get_all_values()
        return any(
            (r + [""] * 5)[4].strip() == "Admin"
            for r in rows[1:]
        )
    except Exception:
        return False


# ── App Settings (admin-controlled feature flags) ────────────────────────────

@st.cache_data(ttl=60)
def register_session(user_id: str, token: str) -> dict:
    """Write session token for user to Settings (key=sess_<uid>)."""
    import time as _t
    return set_setting(f"sess_{user_id}", f"{token}|{int(_t.time())}", user_id)


def check_session_valid(user_id: str, token: str) -> bool:
    """Return True if stored session token matches the one in this session."""
    stored = get_setting(f"sess_{user_id}", "")
    if not stored or "|" not in stored:
        return False
    stored_token = stored.rsplit("|", 1)[0]
    return stored_token == token


def clear_session(user_id: str) -> None:
    """Remove the active session token for a user (on logout/timeout)."""
    set_setting(f"sess_{user_id}", "", user_id)


def get_setting(key: str, default: str = "FALSE") -> str:
    """Read a single value from the Settings tab (cached 60 s)."""
    try:
        ws   = _ensure_ws(TABS["SETTINGS"], _SETTINGS_HEADERS)
        rows = ws.get_all_values()
        for row in rows[1:]:
            row = (row + [""] * 4)[:4]
            if row[0].strip() == key:
                return row[1].strip() or default
        return default
    except Exception:
        return default


def set_setting(key: str, value: str, updated_by: str = "system") -> dict:
    """Write/update a value in the Settings tab and clear the read cache."""
    try:
        ws   = _ensure_ws(TABS["SETTINGS"], _SETTINGS_HEADERS)
        rows = ws.get_all_values()
        now  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for i, row in enumerate(rows[1:], start=2):
            row = (row + [""] * 4)[:4]
            if row[0].strip() == key:
                ws.update(f"B{i}:D{i}", [[value, updated_by, now]])
                get_setting.clear()
                return {"ok": True}
        ws.append_row([key, value, updated_by, now], value_input_option="RAW")
        get_setting.clear()
        return {"ok": True}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


# ── EmailMaster: dynamic email maps ──────────────────────────────────────────

@st.cache_data(ttl=300)
def get_email_master_maps() -> tuple:
    """Read EmailMaster sheet → (loc_map, zone_map).

    Returns (None, None) if the tab is empty or missing — callers should
    fall back to the hardcoded dicts in emails.py.
    ttl=300 means email changes in the sheet take effect within 5 minutes.
    """
    try:
        ws   = _ensure_ws(TABS["EMAIL_MASTER"], _EMAIL_MASTER_HEADERS)
        rows = _api_call(ws.get_all_values)
        if len(rows) < 2:
            return None, None
        loc_map  = {}
        zone_map = {}
        for row in rows[1:]:
            row = (row + [""] * 5)[:5]
            t, code, _name, email, cc = [c.strip() for c in row]
            if not code or not email:
                continue
            if t.lower() == "location":
                loc_map[code] = email
            elif t.lower() == "zone":
                zone_map[code] = {"to": email, "cc": cc}
        return (loc_map or None), (zone_map or None)
    except Exception:
        return None, None


def seed_email_master(location_map: dict, zone_map: dict) -> dict:
    """Populate (or overwrite) the EmailMaster tab from the given dicts.

    Called once from the Mail Trigger page to migrate hardcoded data to the
    sheet so the admin can edit it there going forward.
    """
    try:
        ws = _ensure_ws(TABS["EMAIL_MASTER"], _EMAIL_MASTER_HEADERS)
        ws.batch_clear(["A2:E2000"])
        rows = []
        for code, email in sorted(location_map.items()):
            rows.append(["Location", code, "", email, ""])
        for zone, v in sorted(zone_map.items()):
            rows.append(["Zone", zone, zone, v.get("to", ""), v.get("cc", "")])
        if rows:
            ws.append_rows(rows, value_input_option="RAW")
        get_email_master_maps.clear()
        return {"ok": True, "count": len(rows)}
    except Exception as e:
        return {"ok": False, "msg": str(e)}


# ── Phase-7: Excel Template Download & Upload ─────────────────────────────────

def generate_mis_template(
    user_id: str,
    month_year: str,
    user_info: dict,
    existing_draft: dict | None = None,
    loc_type: str = "HPCL",
) -> bytes:
    """Build a validated, protected .xlsx workbook with 4 sheets.

    Sheet 1 – MIS Data:
      Row 1  Section banners (HPCL gold, merged per section)
      Row 2  Field labels (HPCL blue; auto fields green)
      Row 3  Hints row (grey italic; includes unit / range / options)
      Row 4  Data entry — pre-filled from draft
             • Identity cols + auto-calc cols → LOCKED
             • User-input cols → UNLOCKED
             • Excel IFERROR formulas for all auto-calc fields
             • Data-validation rules (decimal / whole / list) per field
             • Input-message tooltip on every editable cell
             • Error alert (STOP) on every validated cell

    Sheets 2-4 – Railway Claims, IRR Details, Legal Cases:
      Header row locked; data rows unlocked, pre-filled from saved detail tables.

    Sheet protection password: HPCL@MIS (unlock in Excel if bulk edit needed).
    """
    import io as _io
    import re as _re
    from openpyxl import Workbook
    from openpyxl.styles import (
        Font, PatternFill, Alignment, Border, Side, Protection
    )
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    from form_defs import SECTION_FIELDS, SECTION_NAMES

    PROTECT_PW = "HPCL@MIS"
    BIG_NUM    = "1E+15"          # effective "no upper limit"

    wb    = Workbook()
    draft = existing_draft or {}

    # ── Shared styles ────────────────────────────────────────────────────
    _thin   = Side(style="thin", color="CCCCCC")
    _border = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

    def _fill(hex6):
        return PatternFill("solid", fgColor=hex6)

    def _font(bold=False, color="000000", size=10, italic=False):
        return Font(bold=bold, color=color, size=size, italic=italic)

    BLUE_FILL  = _fill("0033A0")
    GOLD_FILL  = _fill("C6A64A")
    AUTO_FILL  = _fill("D6F5E0")   # light green — auto-calc
    LOCK_FILL  = _fill("F0F4FF")   # soft blue — locked identity
    HINT_FILL  = _fill("F8F9FA")
    WHITE_FILL = _fill("FFFFFF")
    NA_FILL    = _fill("D0D0D0")   # grey — N/A for this location type

    W_FONT  = _font(bold=True,  color="FFFFFF", size=9)
    AU_FONT = _font(italic=True, color="1a7a3c", size=9)
    HT_FONT = _font(italic=True, color="888888", size=8)
    NM_FONT = _font(size=10)
    ID_FONT = _font(bold=True,  color="0033A0", size=10)

    CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    LOCKED_CELL   = Protection(locked=True)
    UNLOCKED_CELL = Protection(locked=False)

    def _cell(ws, row, col, value=None, font=None, fill=None,
              align=None, bdr=True, lock=True):
        c = ws.cell(row=row, column=col, value=value)
        if font:  c.font       = font
        if fill:  c.fill       = fill
        if align: c.alignment  = align
        if bdr:   c.border     = _border
        c.protection = LOCKED_CELL if lock else UNLOCKED_CELL
        return c

    # ── Location-type exclusions ─────────────────────────────────────────
    from form_defs import get_excluded_fields as _get_excl
    _excl_keys = _get_excl(loc_type)   # frozenset of field keys not applicable

    # ── Collect all fields in section order ──────────────────────────────
    all_fields = []
    for sn in sorted(SECTION_FIELDS):
        for f in SECTION_FIELDS[sn]:
            all_fields.append((sn, f))

    ID_COLS = ["User ID", "Location Name", "Zone", "Month-Year"]
    ID_VALS = [user_id, user_info.get("locName", ""),
               user_info.get("zone", ""), month_year]
    N_ID    = len(ID_COLS)

    # ── Map field key → Excel column number (row 4) ──────────────────────
    field_col = {}
    for idx, (sn, f) in enumerate(all_fields):
        field_col[f["key"]] = N_ID + 1 + idx

    def _to_xl_formula(expr: str, row: int = 4) -> str:
        """Convert Python field expression (e.g. 'f17+f18+f19') to Excel formula."""
        def _sub(m):
            col = field_col.get(m.group(0))
            return f"{get_column_letter(col)}{row}" if col else "0"
        xl = _re.sub(r'f\d+', _sub, expr)
        return f'=IFERROR({xl},"")' if "/" in xl else f"={xl}"

    # ── Sheet 1: MIS Data ────────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "MIS Data"

    # Row 1 — Section banners (merged, HPCL gold)
    for ci, h in enumerate(ID_COLS, 1):
        _cell(ws1, 1, ci, h, font=W_FONT, fill=BLUE_FILL, align=CENTER)
        ws1.column_dimensions[get_column_letter(ci)].width = 16

    sec_col = N_ID + 1
    for sn in sorted(SECTION_FIELDS):
        fields = SECTION_FIELDS[sn]
        end_c  = sec_col + len(fields) - 1
        _cell(ws1, 1, sec_col, SECTION_NAMES[sn],
              font=_font(bold=True, color="FFFFFF", size=9),
              fill=GOLD_FILL, align=CENTER)
        if len(fields) > 1:
            ws1.merge_cells(start_row=1, start_column=sec_col,
                            end_row=1,   end_column=end_c)
        sec_col += len(fields)
    ws1.row_dimensions[1].height = 30

    # Row 2 — Field labels
    for ci, h in enumerate(ID_COLS, 1):
        _cell(ws1, 2, ci, h, font=W_FONT, fill=BLUE_FILL, align=CENTER)

    for ci, (sn, f) in enumerate(all_fields, N_ID + 1):
        is_auto = bool(f.get("auto"))
        is_na   = f["key"] in _excl_keys
        lbl = f["label"] + (" *" if f.get("req") and not is_auto else "")
        if is_auto:
            lbl += "  [Auto-Calc]"
        if is_na:
            lbl = f["label"] + "  [N/A — Not Applicable]"
        _cell(ws1, 2, ci, lbl,
              font=_font(italic=True, color="666666", size=9) if is_na
                   else (AU_FONT if is_auto else W_FONT),
              fill=NA_FILL if is_na else (AUTO_FILL if is_auto else BLUE_FILL),
              align=CENTER)
        ws1.column_dimensions[get_column_letter(ci)].width = 22
    ws1.row_dimensions[2].height = 60

    # Row 3 — Hints / validation notes
    for ci in range(1, N_ID + 1):
        _cell(ws1, 3, ci, "Pre-filled  —  do not edit",
              font=HT_FONT, fill=HINT_FILL, align=CENTER)

    for ci, (sn, f) in enumerate(all_fields, N_ID + 1):
        if f["key"] in _excl_keys:
            _cell(ws1, 3, ci, f"Not applicable for {loc_type} locations — leave blank",
                  font=HT_FONT, fill=NA_FILL, align=CENTER)
            continue
        parts = [f.get("hint", "")]
        if f.get("min") is not None:
            parts.append(f"Min: {f['min']}")
        if f.get("max") is not None:
            parts.append(f"Max: {f['max']}")
        if f["type"] == "number" and f.get("dec") is not None:
            parts.append(f"Decimals: up to {f['dec']}")
        if f["type"] == "int":
            parts.append("Whole numbers only")
        if f.get("opts"):
            parts.append("Options: " + " / ".join(f["opts"]))
        if f.get("auto"):
            parts.append("[Auto-calculated — do not enter manually]")
        if f["type"] == "textarea":
            parts.append("Press Alt+Enter to add a new line")
        hint_text = "  |  ".join(p for p in parts if p)
        _cell(ws1, 3, ci, hint_text, font=HT_FONT, fill=HINT_FILL, align=LEFT)
    ws1.row_dimensions[3].height = 30

    # Row 4 — Data (pre-filled + protected + validated)
    for ci, (h, v) in enumerate(zip(ID_COLS, ID_VALS), 1):
        _cell(ws1, 4, ci, v, font=ID_FONT, fill=LOCK_FILL, align=CENTER, lock=True)

    for ci, (sn, f) in enumerate(all_fields, N_ID + 1):
        is_auto   = bool(f.get("auto"))
        is_na     = f["key"] in _excl_keys
        raw       = draft.get(f["key"])
        cell_ref  = f"{get_column_letter(ci)}4"

        if is_na:
            _cell(ws1, 4, ci, "N/A",
                  font=_font(italic=True, color="888888", size=10),
                  fill=NA_FILL, align=CENTER, lock=True)
            continue

        if is_auto:
            formula = _to_xl_formula(f["auto"])
            c = _cell(ws1, 4, ci, formula,
                      font=AU_FONT, fill=AUTO_FILL, align=CENTER, lock=True)
            # Apply decimal number format so auto-calc results display correctly
            _dec = f.get("dec") or 2
            c.number_format = "0." + "0" * _dec
        else:
            val = None
            if raw not in (None, ""):
                try:
                    val = float(raw) if f["type"] in ("number", "int") else str(raw)
                except Exception:
                    val = str(raw)
            c = _cell(ws1, 4, ci, val,
                      font=NM_FONT, fill=WHITE_FILL, align=CENTER, lock=False)

            # ── Data validation ─────────────────────────────────────────
            ftype = f["type"]
            mn    = f.get("min")
            mx    = f.get("max")
            opts  = f.get("opts")
            dec   = f.get("dec")
            lbl   = f["label"]
            hint  = f.get("hint", "")

            # Build a descriptive range string for messages
            range_txt = ""
            if mn is not None and mx is not None:
                range_txt = f"between {mn} and {mx}"
            elif mn is not None:
                range_txt = f"≥ {mn}"
            elif mx is not None:
                range_txt = f"≤ {mx}"

            if ftype in ("number", "int") and (mn is not None or mx is not None):
                xl_type  = "whole" if ftype == "int" else "decimal"
                f1       = str(mn) if mn is not None else "0"
                f2       = str(mx) if mx is not None else BIG_NUM

                if mn is not None and mx is not None:
                    op = "between"
                elif mn is not None:
                    op, f2 = "greaterThanOrEqual", None
                else:
                    op, f1 = "lessThanOrEqual", f2

                dv_kwargs = dict(
                    type=xl_type, operator=op,
                    formula1=f1,
                    allow_blank=True,
                    showErrorMessage=True,
                    errorStyle="stop",
                    errorTitle="Invalid Value",
                    error=(
                        f'"{lbl}" must be a '
                        f'{"whole number" if ftype=="int" else "decimal number"}'
                        + (f" {range_txt}" if range_txt else "")
                        + ("." if not range_txt else "")
                        + (f" (up to {dec} decimal places)" if ftype == "number" and dec is not None else "")
                    ),
                    showInputMessage=True,
                    promptTitle=lbl[:32],
                    prompt=(
                        hint
                        + (f"\nRange: {range_txt}" if range_txt else "")
                        + (f"\nDecimals: up to {dec}" if ftype == "number" and dec is not None else "")
                        + ("\nWhole numbers only" if ftype == "int" else "")
                    ),
                )
                if f2 is not None:
                    dv_kwargs["formula2"] = f2

                dv = DataValidation(**dv_kwargs)
                dv.sqref = cell_ref
                ws1.add_data_validation(dv)

            elif ftype == "select" and opts:
                opts_str = ",".join(opts)
                dv = DataValidation(
                    type="list",
                    formula1=f'"{opts_str}"',
                    allow_blank=True,
                    showErrorMessage=True,
                    errorStyle="stop",
                    errorTitle="Invalid Selection",
                    error=f'Select one of: {opts_str}',
                    showInputMessage=True,
                    promptTitle=lbl[:32],
                    prompt=f"{hint}\nSelect: {opts_str}",
                )
                dv.sqref = cell_ref
                ws1.add_data_validation(dv)

            elif ftype == "date":
                # Store as text in DD/MM/YYYY format; enforce format via custom validation
                c.number_format = "@"  # force text so Excel doesn't auto-convert to serial
                _cr0 = cell_ref       # e.g. "AZ4"
                formula = (
                    f'=AND(LEN({_cr0})=10,'
                    f'ISNUMBER(VALUE(LEFT({_cr0},2))),'
                    f'MID({_cr0},3,1)="/",'
                    f'ISNUMBER(VALUE(MID({_cr0},4,2))),'
                    f'MID({_cr0},6,1)="/",'
                    f'ISNUMBER(VALUE(RIGHT({_cr0},4))),'
                    f'VALUE(LEFT({_cr0},2))>=1,'
                    f'VALUE(LEFT({_cr0},2))<=31,'
                    f'VALUE(MID({_cr0},4,2))>=1,'
                    f'VALUE(MID({_cr0},4,2))<=12)'
                )
                dv = DataValidation(
                    type="custom", formula1=formula,
                    allow_blank=True,
                    showErrorMessage=True,
                    errorStyle="warning",
                    errorTitle="Invalid Date Format",
                    error='Enter date as DD/MM/YYYY — e.g. 25/06/2026',
                    showInputMessage=True,
                    promptTitle=lbl[:32],
                    prompt=f"{hint}\nFormat: DD/MM/YYYY (e.g. 25/06/2026)",
                )
                dv.sqref = cell_ref
                ws1.add_data_validation(dv)

            else:
                # textarea / unconstrained — hint is in row 3; no validation rule needed
                pass

    # Taller data row so textarea cells (multi-line text) are readable
    ws1.row_dimensions[4].height = 60
    ws1.freeze_panes = ws1.cell(row=4, column=N_ID + 1)

    # ── Protect MIS Data sheet ───────────────────────────────────────────
    ws1.protection.sheet             = True
    ws1.protection.password          = PROTECT_PW
    ws1.protection.selectLockedCells   = False   # allow clicking locked cells
    ws1.protection.selectUnlockedCells = False

    # Date field keys per detail tab — these get DD/MM/YYYY formula validation
    _DETAIL_DATE_KEYS = {
        "RAILWAY_CLAIMS": {"last_hearing", "next_hearing"},
        "IRR_DETAILS":    {"irr_date", "closure_date"},
        "LEGAL_CASES":    {"last_hearing", "next_hearing"},
    }

    # ── Helper: protected detail sheet ──────────────────────────────────
    def _detail_sheet(sheet_name: str, tab_key: str):
        ddef      = _DETAIL_DEF[tab_key]
        data_keys = ddef["data_keys"]
        col_hdrs  = ddef["sheet_headers"][ddef["prefix_count"]:]
        existing  = load_detail_table(user_id, month_year, tab_key)
        date_keys = _DETAIL_DATE_KEYS.get(tab_key, set())

        ws2 = wb.create_sheet(sheet_name)

        # Header row — locked
        for ci, lbl in enumerate(col_hdrs, 1):
            _cell(ws2, 1, ci, lbl, font=W_FONT, fill=BLUE_FILL, align=CENTER, lock=True)
            ws2.column_dimensions[get_column_letter(ci)].width = 22
        ws2.row_dimensions[1].height = 30

        # Data rows — unlocked
        data_rows = existing if existing else [{}]
        for ri, rec in enumerate(data_rows, 2):
            for ci, key in enumerate(data_keys, 1):
                val = rec.get(key) or None
                _cell(ws2, ri, ci, val,
                      font=NM_FONT, fill=WHITE_FILL, align=LEFT, lock=False)
            ws2.row_dimensions[ri].height = 20

        # Extend blank unlocked rows to row 200
        start_blank = 2 + len(data_rows)
        for ri in range(start_blank, 201):
            for ci in range(1, len(col_hdrs) + 1):
                ws2.cell(row=ri, column=ci).protection = UNLOCKED_CELL

        # Data validation per column — DD/MM/YYYY for dates, input-message for others
        for ci, (lbl, key) in enumerate(zip(col_hdrs, data_keys), 1):
            cl      = get_column_letter(ci)
            col_ref = f"{cl}2:{cl}200"
            if key in date_keys:
                cell0   = f"{cl}2"
                formula = (f'=AND(LEN({cell0})=10,'
                           f'MID({cell0},3,1)="/",'
                           f'MID({cell0},6,1)="/")')
                dv = DataValidation(
                    type="custom", formula1=formula,
                    allow_blank=True, showErrorMessage=True,
                    errorStyle="warning", errorTitle="Invalid Date Format",
                    error='Enter date as DD/MM/YYYY (e.g. 25/06/2025) or "NA" if unknown.',
                    showInputMessage=True, promptTitle=lbl[:32],
                    prompt="DD/MM/YYYY — e.g. 25/06/2025  (or NA)")
            else:
                dv = DataValidation(
                    allow_blank=True, showInputMessage=True,
                    promptTitle=lbl[:32],
                    prompt=f'Enter: {lbl}  (or NA if not applicable)',
                    showErrorMessage=False)
            dv.sqref = col_ref
            ws2.add_data_validation(dv)

        ws2.freeze_panes = ws2["A2"]

        ws2.protection.sheet             = True
        ws2.protection.password          = PROTECT_PW
        ws2.protection.selectLockedCells   = False
        ws2.protection.selectUnlockedCells = False

    _detail_sheet("Railway Claims", "RAILWAY_CLAIMS")
    _detail_sheet("IRR Details",    "IRR_DETAILS")
    _detail_sheet("Legal Cases",    "LEGAL_CASES")

    # ── S5A: all 10 M&I MIS subsection sheets ────────────────────────────
    # Each entry: (tab_key, sheet_name, headers, hints, keys, dropdowns)
    # dropdowns: dict of key → "opt1,opt2,..." for list-validation columns

    # Build tank dropdown for tank_no columns — inline list (reliable) with
    # hidden-sheet fallback for locations that have more than ~30 tanks.
    _loc_tanks    = get_tank_master().get(user_id, [])
    _loc_tanks_all = _loc_tanks + ["Other Tanks"]
    _tank_inline  = ",".join(_loc_tanks_all)          # e.g. "T-001,T-002,Other Tanks"
    _use_inline   = len(_tank_inline) <= 250           # Excel list-validation limit
    _n_tanks      = len(_loc_tanks_all)
    # Hidden TankList sheet — used when inline exceeds 250 chars
    ws_tl = wb.create_sheet("TankList")
    ws_tl.sheet_state = "hidden"
    for _ti, _tn in enumerate(_loc_tanks_all, 1):
        ws_tl.cell(row=_ti, column=1, value=_tn)

    _MI_TAB_DEFS = [
        ("MI_TANK_OUTAGE", "S5A-1 Tank Outage",
         ["Tank No.", "Other Tank Desc.", "Planned Start", "Planned End",
          "Actual Start", "Actual End", "Outage For", "Current Status"],
         ["Select tank number", "Describe if 'Other Tanks'",
          "DD/MM/YYYY", "DD/MM/YYYY", "DD/MM/YYYY", "DD/MM/YYYY",
          "Reason for outage", "Current status of outage"],
         ["tank_no", "other_tank_desc", "planned_start", "planned_end",
          "actual_start", "actual_end", "outage_for", "current_status"],
         {}),

        ("MI_MAJOR_REPAIR", "S5A-2 Major Repair",
         ["Tank No.", "Other Tank Desc.", "Nature of Repair",
          "Revenue / Capex", "AR Code", "Status", "ETC Date"],
         ["Select tank number", "Describe if 'Other Tanks'",
          "Describe nature of repair", "Select: Revenue or Capex",
          "AR code if applicable", "Current repair status", "DD/MM/YYYY — Expected completion"],
         ["tank_no", "other_tank_desc", "nature_of_repair",
          "revenue_capex", "ar_code", "current_status", "etc_date"],
         {"revenue_capex": "Revenue,Capex"}),

        ("MI_VRU", "S5A-3 VRU",
         ["VRU Operational", "Date Not Operating", "Action Taken", "ETC Date",
          "MS Vol Recovered (KL)", "Inlet MFM Start (m³)", "Inlet MFM End (m³)",
          "Outlet MFM Start (m³)", "Outlet MFM End (m³)", "Vapour Treated (m³)",
          "VOC Value (mg/cc)", "Inlet Emission (mg/cc)",
          "MS/Gasohol TT Vol (KL)", "HSD TT Vol (KL)",
          "MS/Gasohol TW Vol (KL)", "HSD TW Vol (KL)", "VRU Uptime (%)"],
         ["Yes/No", "DD/MM/YYYY if not operating", "Action taken if not operating",
          "DD/MM/YYYY — Expected completion", "Numeric",
          "Numeric (m³)", "Numeric (m³)", "Numeric (m³)", "Numeric (m³)",
          "Numeric (m³)", "VOC concentration at VRU outlet mg/cc",
          "VOC concentration at VRU inlet mg/cc",
          "Numeric (KL)", "Numeric (KL)", "Numeric (KL)", "Numeric (KL)", "0–100"],
         ["vru_operational", "date_not_operating", "action_taken", "etc_date",
          "ms_vol_recovered_kl", "inlet_mfm_start_m3", "inlet_mfm_end_m3",
          "outlet_mfm_start_m3", "outlet_mfm_end_m3", "vapour_treated_m3",
          "voc_value_mgcc", "inlet_emission_mgcc",
          "ms_gasohol_tt_vol_kl", "hsd_tt_vol_kl",
          "ms_gasohol_tw_vol_kl", "hsd_tw_vol_kl", "vru_uptime_pct"],
         {"vru_operational": "Yes,No"}),

        ("MI_AUDIT_2526", "S5A-4 M&I Audit 25-26",
         ["Audit Date", "No. of Recommendations", "No. Pending", "External Score"],
         ["DD/MM/YYYY", "Total recommendations from audit",
          "Pending recommendations", "Score from external auditor"],
         ["audit_date", "no_recommendations", "no_pending", "external_score"],
         {}),

        ("MI_AUDIT_2627", "S5A-5 M&I Audit 26-27",
         ["Audit Carried Out", "Audit Date", "No. of Recommendations",
          "No. Pending", "External Score"],
         ["Yes/No", "DD/MM/YYYY", "Total recommendations",
          "Pending recommendations", "Score from external auditor"],
         ["audit_carried_out", "audit_date", "no_recommendations",
          "no_pending", "external_score"],
         {"audit_carried_out": "Yes,No"}),

        ("MI_TECH_AUDIT", "S5A-6 Tech. Audit",
         ["Audit Date", "No. of Recommendations", "No. Pending", "Ref. No."],
         ["DD/MM/YYYY", "Total recommendations", "Pending count", "Reference number"],
         ["audit_date", "no_recommendations", "no_pending", "ref_no"],
         {}),

        ("MI_EQUIP_BREAKDOWN", "S5A-7 Equip. Breakdown",
         ["Equipment Name", "Equipment Other", "Equipment Details",
          "Start Date", "Issue", "Proposed Date", "Actual End Date",
          "Resolution Action"],
         ["Select equipment type", "Specify if 'Other'", "Details of breakdown",
          "DD/MM/YYYY", "Describe the issue",
          "DD/MM/YYYY — Proposed fix date", "DD/MM/YYYY — Actual resolution",
          "Action taken to resolve"],
         ["equipment_name", "equipment_other", "equipment_details",
          "start_date", "issue", "proposed_date", "actual_end_date",
          "resolution_action"],
         {"equipment_name": "Pipeline,Pump,Fire Fighting Equipment,Fire Engine,DG Set,Other"}),

        ("MI_INT_PIPELINE", "S5A-8 Int. Pipeline",
         ["Last UT Date", "Last Hydrotest Date", "Last DCVG Date",
          "Last LRUT Date", "Other Testing"],
         ["DD/MM/YYYY", "DD/MM/YYYY", "DD/MM/YYYY", "DD/MM/YYYY",
          "Describe any other testing done"],
         ["last_ut_date", "last_hydrotest_date", "last_dcvg_date",
          "last_lrut_date", "other_testing"],
         {}),

        ("MI_EXT_PIPELINE", "S5A-9 Ext. Pipeline",
         ["Pipeline Type", "Pipeline Details", "Length Metres", "Product", "Size Inch",
          "Last UT Date", "Last Hydrotest Date", "Last DCVG Date",
          "Last LRUT Date", "Other Testing"],
         ["UG = Underground / AG = Above Ground",
          "Describe pipeline segment (route / from-to)", "Length in metres",
          "Product carried e.g. MS HSD ATF",
          "Nominal bore in inches", "DD/MM/YYYY", "DD/MM/YYYY", "DD/MM/YYYY",
          "DD/MM/YYYY", "Describe any other testing"],
         ["pipeline_type", "pipeline_details", "length_metres", "product", "size_inch",
          "last_ut_date", "last_hydrotest_date", "last_dcvg_date",
          "last_lrut_date", "other_testing"],
         {"pipeline_type": "UG,AG"}),

        ("MI_TANK_STATUS", "S5A-10 Tank Status",
         ["Tank No", "Cleaning Completed Date", "Cleaning Due Date",
          "Extension Taken", "Extension EFN No",
          "Inspection Date", "Inspection Due Date",
          "Painting Date", "Painting Due Date",
          "Tank Status", "Tank Status Other"],
         ["Select tank number from Tank Master",
          "DD/MM/YYYY", "DD/MM/YYYY",
          "Yes / No / NA", "Required if Extension = Yes",
          "DD/MM/YYYY", "DD/MM/YYYY",
          "DD/MM/YYYY", "DD/MM/YYYY",
          "Operational / Under Repair / Under Cleaning / Idle / Revamp / Others",
          "Required if Tank Status = Others"],
         ["tank_no",
          "cleaning_completed_date", "cleaning_due_date",
          "extension_taken", "extension_efn_no",
          "inspection_date", "inspection_due_date",
          "painting_date", "painting_due_date",
          "tank_status", "tank_status_other"],
         {"extension_taken": "Yes,No,NA",
          "tank_status":     "Operational,Under Repair,Under Cleaning,Idle,Revamp,Others"}),
    ]

    BANNER_FILL = _fill("1a1a6e")
    BANNER_FONT = _font(bold=True, color="FFFFFF", size=11)
    NA_FONT     = _font(italic=True, color="888888", size=10)
    NA_FILL     = _fill("F5F5F5")

    # Date field keys per M&I tab — these get DD/MM/YYYY formula validation
    _MI_DATE_KEYS: dict[str, set] = {
        "MI_TANK_OUTAGE":     {"planned_start", "planned_end", "actual_start", "actual_end"},
        "MI_MAJOR_REPAIR":    {"etc_date"},
        "MI_VRU":             {"date_not_operating", "etc_date"},
        "MI_AUDIT_2526":      {"audit_date"},
        "MI_AUDIT_2627":      {"audit_date"},
        "MI_TECH_AUDIT":      {"audit_date"},
        "MI_EQUIP_BREAKDOWN": {"start_date", "proposed_date", "actual_end_date"},
        "MI_INT_PIPELINE":    {"last_ut_date", "last_hydrotest_date",
                               "last_dcvg_date", "last_lrut_date"},
        "MI_EXT_PIPELINE":    {"last_ut_date", "last_hydrotest_date",
                               "last_dcvg_date", "last_lrut_date"},
        "MI_TANK_STATUS":     {"cleaning_completed_date", "cleaning_due_date",
                               "inspection_date", "inspection_due_date",
                               "painting_date", "painting_due_date"},
    }

    for tab_key, sheet_name, hdrs, hints, keys, dropdowns in _MI_TAB_DEFS:
        ws_mi  = wb.create_sheet(sheet_name)
        n_cols = len(hdrs)
        date_keys_tab = _MI_DATE_KEYS.get(tab_key, set())

        # Row 1 — banner
        _cell(ws_mi, 1, 1, sheet_name,
              font=BANNER_FONT, fill=BANNER_FILL, align=CENTER)
        ws_mi.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
        ws_mi.row_dimensions[1].height = 28

        # Row 2 — headers
        for ci, h in enumerate(hdrs, 1):
            _cell(ws_mi, 2, ci, h, font=W_FONT, fill=BLUE_FILL, align=CENTER, lock=True)
            ws_mi.column_dimensions[get_column_letter(ci)].width = 24
        ws_mi.row_dimensions[2].height = 40

        # Row 3 — hints (date fields show DD/MM/YYYY prominently)
        for ci, (h, key) in enumerate(zip(hints, keys), 1):
            display_hint = (f"DD/MM/YYYY — {h}" if key in date_keys_tab
                            else ("NA if not applicable — " + h if not h.startswith("NA") else h))
            _cell(ws_mi, 3, ci, display_hint, font=HT_FONT, fill=HINT_FILL, align=LEFT, lock=True)
        ws_mi.row_dimensions[3].height = 22

        # Pre-fill saved data rows starting at row 4
        existing   = load_mi_data(tab_key, user_id, month_year)
        is_na_tab  = bool(existing and existing[0].get("na_flag") == "Y")
        saved_rows = ([r for r in existing if r.get("na_flag") != "Y"]
                      if existing and not is_na_tab else [])

        if is_na_tab:
            # Write "Not Applicable" in first cell so users know tab is marked NA
            c = ws_mi.cell(row=4, column=1, value="Not Applicable — marked NA in application")
            c.font = NA_FONT; c.fill = NA_FILL; c.alignment = LEFT
            ws_mi.merge_cells(start_row=4, start_column=1, end_row=4, end_column=n_cols)
            ws_mi.row_dimensions[4].height = 20
            start_blank = 5
        else:
            for ri, rec in enumerate(saved_rows, 4):
                for ci, key in enumerate(keys, 1):
                    val = rec.get(key) or None
                    _cell(ws_mi, ri, ci, val, font=NM_FONT, fill=WHITE_FILL,
                          align=LEFT, lock=False)
                    if key in date_keys_tab:
                        ws_mi.cell(row=ri, column=ci).number_format = "@"
                ws_mi.row_dimensions[ri].height = 20
            start_blank = 4 + len(saved_rows)

        # Blank unlocked rows to row 150 — date columns forced to text format
        for ri in range(start_blank, 151):
            for ci in range(1, n_cols + 1):
                ws_mi.cell(row=ri, column=ci).protection = UNLOCKED_CELL
                if keys[ci - 1] in date_keys_tab:
                    ws_mi.cell(row=ri, column=ci).number_format = "@"

        # Data validations per column (DD/MM/YYYY formula for dates; list for dropdowns)
        for ci, (hdr, hint, key) in enumerate(zip(hdrs, hints, keys), 1):
            cl      = get_column_letter(ci)
            col_ref = f"{cl}4:{cl}150"
            if key in date_keys_tab:
                # Custom formula: check DD/MM/YYYY text format
                cell0   = f"{cl}4"
                formula = (f'=AND(LEN({cell0})=10,'
                           f'MID({cell0},3,1)="/",'
                           f'MID({cell0},6,1)="/")')
                dv = DataValidation(
                    type="custom", formula1=formula,
                    allow_blank=True, showErrorMessage=True,
                    errorStyle="warning", errorTitle="Invalid Date Format",
                    error='Enter date as DD/MM/YYYY (e.g. 25/06/2025) or "NA" if not applicable.',
                    showInputMessage=True, promptTitle=hdr[:32],
                    prompt=f"DD/MM/YYYY (e.g. 25/06/2025) or NA")
            elif key == "tank_no":
                _f1 = (f'"{_tank_inline}"' if _use_inline
                       else f"'TankList'!$A$1:$A${_n_tanks}")
                dv = DataValidation(
                    type="list", formula1=_f1,
                    allow_blank=True, showErrorMessage=True,
                    errorStyle="warning",
                    error="Select a tank number from your location's Tank Master list, or 'Other Tanks'",
                    showInputMessage=True, promptTitle="Tank No.",
                    prompt="Select tank from Tank Master list for this location")
            elif key in dropdowns:
                dv = DataValidation(
                    type="list", formula1=f'"{dropdowns[key]}"',
                    allow_blank=True, showErrorMessage=True,
                    errorStyle="warning", error=f"Select: {dropdowns[key]}",
                    showInputMessage=True, promptTitle=hdr[:32], prompt=hint)
            else:
                dv = DataValidation(
                    allow_blank=True, showInputMessage=True,
                    promptTitle=hdr[:32],
                    prompt=f'Enter value, or "NA" if not applicable.',
                    showErrorMessage=False)
            dv.sqref = col_ref
            ws_mi.add_data_validation(dv)

        ws_mi.freeze_panes = ws_mi["A4"]
        ws_mi.protection.sheet             = True
        ws_mi.protection.password          = PROTECT_PW
        ws_mi.protection.selectLockedCells   = False
        ws_mi.protection.selectUnlockedCells = False

    buf = _io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def generate_filled_mis_report(
    user_id: str,
    month_year: str,
    user_info: dict,
    existing_draft: dict | None = None,
) -> bytes:
    """Generate a read-only filled MIS report Excel — all cells locked.

    Intended for download AFTER Checker approval (status = SUBMITTED).
    Same layout as generate_mis_template but fully locked and pre-filled.
    """
    import io as _io
    import re as _re
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
    from openpyxl.utils import get_column_letter
    from form_defs import SECTION_FIELDS, SECTION_NAMES

    draft = existing_draft or {}
    wb    = Workbook()

    _thin   = Side(style="thin", color="CCCCCC")
    _border = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

    def _fill(h):
        return PatternFill("solid", fgColor=h)
    def _font(bold=False, color="000000", size=10, italic=False):
        return Font(bold=bold, color=color, size=size, italic=italic)

    BLUE_FILL  = _fill("0033A0")
    GOLD_FILL  = _fill("C6A64A")
    AUTO_FILL  = _fill("E8F5E9")
    HINT_FILL  = _fill("F8F9FA")
    WHITE_FILL = _fill("FFFFFF")
    W_FONT     = _font(bold=True, color="FFFFFF", size=9)
    HT_FONT    = _font(italic=True, color="888888", size=8)
    NM_FONT    = _font(size=10)
    AU_FONT    = _font(italic=True, color="1a7a3c", size=9)
    CENTER     = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT       = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    LOCKED     = Protection(locked=True)

    def _cell(ws, row, col, value=None, font=None, fill=None, align=None):
        c = ws.cell(row=row, column=col, value=value)
        if font:  c.font      = font
        if fill:  c.fill      = fill
        if align: c.alignment = align
        c.border     = _border
        c.protection = LOCKED
        return c

    all_fields = []
    for sn in sorted(SECTION_FIELDS):
        for f in SECTION_FIELDS[sn]:
            all_fields.append((sn, f))

    ID_COLS = ["User ID", "Location Name", "Zone", "Month-Year"]
    ID_VALS = [user_id, user_info.get("locName", ""),
               user_info.get("zone", ""), month_year]
    N_ID    = len(ID_COLS)

    field_col = {f["key"]: N_ID + 1 + idx for idx, (_, f) in enumerate(all_fields)}

    def _to_xl(expr, row=4):
        def _sub(m):
            col = field_col.get(m.group(0))
            return f"{get_column_letter(col)}{row}" if col else "0"
        xl = _re.sub(r'f\d+', _sub, expr)
        return f'=IFERROR({xl},"")' if "/" in xl else f"={xl}"

    ws1 = wb.active
    ws1.title = "MIS Report"

    # Row 1 — report banner
    total_cols = N_ID + len(all_fields)
    _cell(ws1, 1, 1,
          f"HPCL SOD MIS REPORT — {user_info.get('locName','')} | {month_year}",
          font=_font(bold=True, color="FFFFFF", size=11),
          fill=_fill("001060"), align=CENTER)
    ws1.merge_cells(start_row=1, start_column=1, end_row=1, end_column=min(total_cols, 50))
    ws1.row_dimensions[1].height = 28

    # Row 2 — Section banners
    for ci, h in enumerate(ID_COLS, 1):
        _cell(ws1, 2, ci, h, font=W_FONT, fill=BLUE_FILL, align=CENTER)
        ws1.column_dimensions[get_column_letter(ci)].width = 16

    sec_col = N_ID + 1
    for sn in sorted(SECTION_FIELDS):
        fields = SECTION_FIELDS[sn]
        end_c  = sec_col + len(fields) - 1
        _cell(ws1, 2, sec_col, SECTION_NAMES[sn],
              font=_font(bold=True, color="FFFFFF", size=9),
              fill=GOLD_FILL, align=CENTER)
        if len(fields) > 1:
            ws1.merge_cells(start_row=2, start_column=sec_col,
                            end_row=2,   end_column=end_c)
        sec_col += len(fields)
    ws1.row_dimensions[2].height = 28

    # Row 3 — Field labels
    for ci, h in enumerate(ID_COLS, 1):
        _cell(ws1, 3, ci, h, font=W_FONT, fill=BLUE_FILL, align=CENTER)

    for ci, (sn, f) in enumerate(all_fields, N_ID + 1):
        is_auto = bool(f.get("auto"))
        lbl     = f["label"] + ("  [Auto]" if is_auto else "")
        _cell(ws1, 3, ci, lbl,
              font=AU_FONT if is_auto else W_FONT,
              fill=AUTO_FILL if is_auto else BLUE_FILL,
              align=CENTER)
        ws1.column_dimensions[get_column_letter(ci)].width = 18
    ws1.row_dimensions[3].height = 30

    # Row 4 — Hint row
    for ci, _ in enumerate(ID_COLS, 1):
        _cell(ws1, 4, ci, "", font=HT_FONT, fill=HINT_FILL, align=LEFT)
    for ci, (_, f) in enumerate(all_fields, N_ID + 1):
        hint = f.get("hint", "")
        _cell(ws1, 4, ci, hint, font=HT_FONT, fill=HINT_FILL, align=LEFT)
    ws1.row_dimensions[4].height = 20

    # Row 5 — Data values
    for ci, v in enumerate(ID_VALS, 1):
        _cell(ws1, 5, ci, v, font=_font(bold=True, color="0033A0", size=10),
              fill=_fill("F0F4FF"), align=LEFT)

    for ci, (sn, f) in enumerate(all_fields, N_ID + 1):
        if f.get("auto") and f.get("auto"):
            val = _to_xl(f["auto"], row=5)
        else:
            raw = draft.get(f["key"], "")
            val = raw if raw not in ("", None) else ""
        _cell(ws1, 5, ci, val, font=NM_FONT, fill=WHITE_FILL, align=LEFT)
    ws1.row_dimensions[5].height = 22

    ws1.freeze_panes = ws1["A5"]

    ws1.protection.sheet             = True
    ws1.protection.password          = "HPCL@MIS"
    ws1.protection.selectLockedCells   = False
    ws1.protection.selectUnlockedCells = False

    buf = _io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def generate_mi_mis_report(
    user_id: str,
    month_year: str,
    user_info: dict,
) -> bytes:
    """Generate filled M&I MIS Excel report with all 10 subsection tabs."""
    import io as _io
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
    from openpyxl.utils import get_column_letter

    wb = Workbook()

    _thin   = Side(style="thin", color="CCCCCC")
    _border = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

    def _fill(h):
        return PatternFill("solid", fgColor=h)
    def _font(bold=False, color="000000", size=10, italic=False):
        return Font(bold=bold, color=color, size=size, italic=italic)

    BLUE_FILL  = _fill("1a1a6e")
    HDR_FONT   = _font(bold=True, color="FFFFFF", size=9)
    NM_FONT    = _font(size=10)
    HINT_FILL  = _fill("F8F9FA")
    HT_FONT    = _font(italic=True, color="888888", size=8)
    WHITE_FILL = _fill("FFFFFF")
    CENTER     = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT       = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    LOCKED     = Protection(locked=True)

    def _hdr_row(ws, headers, row=1):
        for ci, h in enumerate(headers, 1):
            c = ws.cell(row=row, column=ci, value=h)
            c.font = HDR_FONT; c.fill = BLUE_FILL
            c.alignment = CENTER; c.border = _border
            c.protection = LOCKED
            ws.column_dimensions[get_column_letter(ci)].width = 22
        ws.row_dimensions[row].height = 28

    def _data_row(ws, row, values):
        for ci, v in enumerate(values, 1):
            c = ws.cell(row=row, column=ci, value=v)
            c.font = NM_FONT; c.fill = WHITE_FILL
            c.alignment = LEFT; c.border = _border
            c.protection = LOCKED
        ws.row_dimensions[row].height = 18

    def _sheet(tab_key: str, sheet_name: str, display_headers: list, data_keys: list):
        ws = wb.create_sheet(sheet_name)
        rows = load_mi_data(tab_key, user_id, month_year)
        _hdr_row(ws, display_headers)
        if not rows:
            ws.cell(row=2, column=1, value="No data saved for this tab.").font = HT_FONT
        elif rows and rows[0].get("na_flag") == "Y":
            ws.cell(row=2, column=1, value="Not Applicable (marked NA)").font = HT_FONT
        else:
            for ri, rec in enumerate(rows, 2):
                vals = [rec.get(k, "") for k in data_keys]
                _data_row(ws, ri, vals)
        ws.freeze_panes = ws["A2"]
        ws.protection.sheet   = True
        ws.protection.password = "HPCL@MIS"
        ws.protection.selectLockedCells   = False
        ws.protection.selectUnlockedCells = False

    # ── Cover sheet ────────────────────────────────────────────────────────
    ws0 = wb.active
    ws0.title = "Cover"
    cover_data = [
        ("Location", user_info.get("locName", user_id)),
        ("Zone",     user_info.get("zone", "")),
        ("Month",    month_year),
        ("Report",   "M&I MIS — Maintenance & Inspection Monthly Information System"),
    ]
    for ri, (k, v) in enumerate(cover_data, 1):
        ws0.cell(row=ri, column=1, value=k).font  = _font(bold=True, color="0033A0", size=11)
        ws0.cell(row=ri, column=2, value=v).font  = _font(size=11)
        ws0.column_dimensions["A"].width = 18
        ws0.column_dimensions["B"].width = 50

    # ── 10 subsection sheets ───────────────────────────────────────────────
    _sheet("MI_TANK_OUTAGE", "Tank Outage",
           ["Tank No.", "Other Tank", "Planned Start", "Planned End",
            "Actual Start", "Actual End", "Outage For", "Current Status"],
           ["tank_no", "other_tank_desc", "planned_start", "planned_end",
            "actual_start", "actual_end", "outage_for", "current_status"])

    _sheet("MI_MAJOR_REPAIR", "Major Repair",
           ["Tank No.", "Other Tank", "Nature of Repair",
            "Revenue/Capex", "AR Code", "Status", "ETC Date"],
           ["tank_no", "other_tank_desc", "nature_of_repair",
            "revenue_capex", "ar_code", "current_status", "etc_date"])

    _sheet("MI_VRU", "VRU",
           ["VRU Operational", "Date Not Operating", "Action Taken", "ETC Date",
            "MS Vol Recovered (KL)", "Inlet MFM Start", "Inlet MFM End",
            "Outlet MFM Start", "Outlet MFM End", "Vapour Treated (m³)",
            "VOC Value (mg/cc)", "Inlet Emission (mg/cc)",
            "MS/Gasohol TT Vol", "HSD TT Vol", "MS/Gasohol TW Vol",
            "HSD TW Vol", "VRU Uptime %"],
           ["vru_operational", "date_not_operating", "action_taken", "etc_date",
            "ms_vol_recovered_kl", "inlet_mfm_start_m3", "inlet_mfm_end_m3",
            "outlet_mfm_start_m3", "outlet_mfm_end_m3", "vapour_treated_m3",
            "voc_value_mgcc", "inlet_emission_mgcc",
            "ms_gasohol_tt_vol_kl", "hsd_tt_vol_kl",
            "ms_gasohol_tw_vol_kl", "hsd_tw_vol_kl", "vru_uptime_pct"])

    _sheet("MI_AUDIT_2526", "Audit 25-26",
           ["Audit Date", "No. of Recommendations", "No. Pending", "External Score"],
           ["audit_date", "no_recommendations", "no_pending", "external_score"])

    _sheet("MI_AUDIT_2627", "Audit 26-27",
           ["Audit Carried Out", "Audit Date", "No. of Recommendations",
            "No. Pending", "External Score"],
           ["audit_carried_out", "audit_date", "no_recommendations",
            "no_pending", "external_score"])

    _sheet("MI_TECH_AUDIT", "Tech. Audit",
           ["Audit Date", "No. of Recommendations", "No. Pending", "Ref. No."],
           ["audit_date", "no_recommendations", "no_pending", "ref_no"])

    _sheet("MI_EQUIP_BREAKDOWN", "Equip. Breakdown",
           ["Equipment Name", "Equipment Other", "Equipment Details",
            "Start Date", "Issue", "Proposed Date", "Actual End Date",
            "Resolution Action"],
           ["equipment_name", "equipment_name_other", "equipment_details",
            "start_date", "issue", "proposed_date", "actual_end_date",
            "resolution_action"])

    _sheet("MI_INT_PIPELINE", "Int. Pipeline",
           ["Last UT Date", "Last Hydrotest Date", "Last DCVG Date",
            "Last LRUT Date", "Other Testing"],
           ["last_ut_date", "last_hydrotest_date", "last_dcvg_date",
            "last_lrut_date", "other_testing"])

    _sheet("MI_EXT_PIPELINE", "Ext. Pipeline",
           ["Type", "Pipeline Details", "Length (m)", "Product", "Size (inch)",
            "Last UT Date", "Last Hydrotest Date", "Last DCVG Date",
            "Last LRUT Date", "Other Testing"],
           ["pipeline_type", "pipeline_details", "length_metres", "product", "size_inch",
            "last_ut_date", "last_hydrotest_date", "last_dcvg_date",
            "last_lrut_date", "other_testing"])

    _sheet("MI_TANK_STATUS", "Tank Status",
           ["Zone", "Location", "Tank No.", "Cleaning Completed Date",
            "Cleaning Due Date", "Extension Taken", "eFN No.",
            "Inspection Date", "Inspection Due Date",
            "Painting Date", "Painting Due Date",
            "Tank Status", "Tank Status (Others)"],
           ["zone", "loc_name", "tank_no",
            "cleaning_completed_date", "cleaning_due_date",
            "extension_taken", "extension_efn_no",
            "inspection_date", "inspection_due_date",
            "painting_date", "painting_due_date",
            "tank_status", "tank_status_other"])

    buf = _io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def parse_mis_upload(file_bytes: bytes) -> dict:
    """Parse an uploaded MIS template (.xlsx) and return structured data.

    Returns:
      {
        "fields":         {field_key: value, ...},   # all 140 MIS fields
        "railway_claims": [{data_key: value}, ...],
        "irr_details":    [{data_key: value}, ...],
        "legal_cases":    [{data_key: value}, ...],
        "errors":         [str, ...],                # non-fatal warnings
      }
    """
    import io as _io
    from datetime import datetime as _dt, date as _date
    from openpyxl import load_workbook
    from form_defs import SECTION_FIELDS

    result = {"fields": {}, "railway_claims": [], "irr_details": [],
              "legal_cases": [], "errors": []}

    def _norm(v):
        """Normalise a cell value: datetime/date → DD/MM/YYYY; else stringify."""
        if v is None:
            return None
        if isinstance(v, (_dt, _date)):
            try:
                return v.strftime("%d/%m/%Y")
            except Exception:
                return str(v)
        sv = str(v).strip()
        # Normalise D/M/YYYY or D/M/YY text strings to zero-padded DD/MM/YYYY
        import re as _re2
        _dm = _re2.match(r'^(\d{1,2})/(\d{1,2})/(\d{2,4})$', sv)
        if _dm:
            _d, _m, _y = _dm.groups()
            if 1 <= int(_m) <= 12 and 1 <= int(_d) <= 31:
                _y = "20" + _y if len(_y) == 2 else _y
                sv = f"{int(_d):02d}/{int(_m):02d}/{_y}"
        return sv if sv else None

    # Build label → key inverse map (case-insensitive, strip * and [Auto])
    label_map = {}
    for fields in SECTION_FIELDS.values():
        for f in fields:
            clean = f["label"].strip().rstrip(" *").replace(" [Auto]", "").strip().lower()
            label_map[clean] = f["key"]

    try:
        wb = load_workbook(_io.BytesIO(file_bytes), data_only=True)
    except Exception as e:
        result["errors"].append(f"Cannot open file: {e}")
        return result

    # ── MIS Data sheet ───────────────────────────────────────────────────
    if "MIS Data" not in wb.sheetnames:
        result["errors"].append("Sheet 'MIS Data' not found in uploaded file.")
    else:
        ws = wb["MIS Data"]
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 4:
            result["errors"].append("MIS Data sheet has fewer than 4 rows — expected header + hint + data.")
        else:
            hdr_row  = [str(v or "").strip().rstrip(" *").replace(" [Auto]", "").strip() for v in rows[1]]
            data_row = rows[3]   # row index 3 = row 4 (data)
            N_ID = 4             # first 4 cols are identity (User ID, Location, Zone, Month-Year)
            for ci, hdr in enumerate(hdr_row):
                if ci < N_ID:
                    continue    # skip identity cols
                key = label_map.get(hdr.lower())
                if key and ci < len(data_row):
                    val = _norm(data_row[ci])
                    if val:
                        result["fields"][key] = val

    # ── Detail sheets ────────────────────────────────────────────────────
    def _parse_detail(sheet_name, tab_key, out_key):
        if sheet_name not in wb.sheetnames:
            result["errors"].append(f"Sheet '{sheet_name}' not found — skipped.")
            return
        ddef      = _DETAIL_DEF[tab_key]
        data_keys = ddef["data_keys"]
        col_hdrs  = ddef["sheet_headers"][ddef["prefix_count"]:]

        # Build header → data_key map
        hdr_to_key = {h.strip().lower(): k for h, k in zip(col_hdrs, data_keys)}

        ws2   = wb[sheet_name]
        rows2 = list(ws2.iter_rows(values_only=True))
        if len(rows2) < 2:
            return

        file_hdrs = [str(v or "").strip().lower() for v in rows2[0]]
        col_map   = {ci: hdr_to_key[h] for ci, h in enumerate(file_hdrs) if h in hdr_to_key}

        for raw_row in rows2[1:]:
            if all(v is None or str(v).strip() == "" for v in raw_row):
                continue  # skip blank rows
            rec = {col_map[ci]: _norm(v)
                   for ci, v in enumerate(raw_row)
                   if ci in col_map and _norm(v)}
            if rec:
                result[out_key].append(rec)

    _parse_detail("Railway Claims", "RAILWAY_CLAIMS", "railway_claims")
    _parse_detail("IRR Details",    "IRR_DETAILS",    "irr_details")
    _parse_detail("Legal Cases",    "LEGAL_CASES",    "legal_cases")

    # ── S5A M&I subsection sheets ────────────────────────────────────────
    # Mapping: (Excel sheet name, tab_key, data_keys list)
    _MI_UPLOAD_DEFS = [
        ("S5A-1 Tank Outage",    "MI_TANK_OUTAGE",
         ["tank_no","other_tank_desc","planned_start","planned_end",
          "actual_start","actual_end","outage_for","current_status"]),
        ("S5A-2 Major Repair",   "MI_MAJOR_REPAIR",
         ["tank_no","other_tank_desc","nature_of_repair",
          "revenue_capex","ar_code","current_status","etc_date"]),
        ("S5A-3 VRU",            "MI_VRU",
         ["vru_operational","date_not_operating","action_taken","etc_date",
          "ms_vol_recovered_kl","inlet_mfm_start_m3","inlet_mfm_end_m3",
          "outlet_mfm_start_m3","outlet_mfm_end_m3","vapour_treated_m3",
          "voc_value_mgcc","inlet_emission_mgcc",
          "ms_gasohol_tt_vol_kl","hsd_tt_vol_kl",
          "ms_gasohol_tw_vol_kl","hsd_tw_vol_kl","vru_uptime_pct"]),
        ("S5A-4 M&I Audit 25-26", "MI_AUDIT_2526",
         ["audit_date","no_recommendations","no_pending","external_score"]),
        ("S5A-5 M&I Audit 26-27", "MI_AUDIT_2627",
         ["audit_carried_out","audit_date","no_recommendations",
          "no_pending","external_score"]),
        ("S5A-6 Tech. Audit",    "MI_TECH_AUDIT",
         ["audit_date","no_recommendations","no_pending","ref_no"]),
        ("S5A-7 Equip. Breakdown","MI_EQUIP_BREAKDOWN",
         ["equipment_name","equipment_other","equipment_details",
          "start_date","issue","proposed_date","actual_end_date","resolution_action"]),
        ("S5A-8 Int. Pipeline",  "MI_INT_PIPELINE",
         ["last_ut_date","last_hydrotest_date","last_dcvg_date",
          "last_lrut_date","other_testing"]),
        ("S5A-9 Ext. Pipeline",  "MI_EXT_PIPELINE",
         ["pipeline_type","pipeline_details","length_metres","product","size_inch",
          "last_ut_date","last_hydrotest_date","last_dcvg_date",
          "last_lrut_date","other_testing"]),
        ("S5A-10 Tank Status",   "MI_TANK_STATUS",
         ["tank_no","cleaning_completed_date","cleaning_due_date",
          "extension_taken","extension_efn_no",
          "inspection_date","inspection_due_date",
          "painting_date","painting_due_date",
          "tank_status","tank_status_other"]),
    ]

    result["mi_tabs"] = {}   # {tab_key: [row_dict, ...] or "NA"}

    import re as _re

    def _clean_h(s: str) -> str:
        """Normalise a header for fuzzy matching: strip dots, parens, slashes,
        unicode superscripts, standalone 'of', then remove all non-alphanumeric."""
        s = s.lower().replace("³", "3").replace("²", "2")
        s = _re.sub(r'\bof\b', '', s)
        return _re.sub(r'[^a-z0-9]', '', s)

    for sheet_name, tab_key, data_keys in _MI_UPLOAD_DEFS:
        if sheet_name not in wb.sheetnames:
            result["errors"].append(f"M&I sheet '{sheet_name}' not found — skipped.")
            continue
        ws_mi   = wb[sheet_name]
        mi_rows = list(ws_mi.iter_rows(values_only=True))
        if len(mi_rows) < 4:
            continue  # banner + header + hint rows; data starts at row 4
        hdr_row  = [str(v or "").strip() for v in mi_rows[1]]  # row 2 = headers
        # Build normalised-label → actual-key lookup (both key form and label form)
        clean_to_key: dict[str, str] = {}
        for dk in data_keys:
            clean_to_key[_clean_h(dk)]                 = dk   # e.g. "norecommendations"
            clean_to_key[_clean_h(dk.replace("_", " "))] = dk  # "no recommendations"
        col_map: dict[int, str] = {}
        for ci, h in enumerate(hdr_row):
            hc = _clean_h(h)
            if hc in clean_to_key:
                col_map[ci] = clean_to_key[hc]

        tab_data = []
        for raw in mi_rows[3:]:   # skip banner(0), header(1), hint(2)
            if all(v is None or str(v).strip() == "" for v in raw):
                continue
            # Detect "Not Applicable" marker row
            first_val = str(raw[0] or "").strip().lower()
            if "not applicable" in first_val or first_val == "na":
                tab_data = "NA"
                break
            rec = {}
            for ci, v in enumerate(raw):
                if ci in col_map:
                    nv = _norm(v)
                    if nv:
                        rec[col_map[ci]] = nv
            if rec:
                tab_data.append(rec)

        result["mi_tabs"][tab_key] = tab_data

    return result


# ── Phase-8/9: Reports & Email ──────────────────────────────────────────────

def get_all_status_for_month(month_year: str) -> list:
    """Return submission status for every Maker location for a given month."""
    locs = get_all_maker_locations()
    return get_submissions_for_locations(locs, month_year)


def download_submitted_data_excel(month_year: str):
    """Build and return an Excel workbook (bytes) of approved MIS data for month_year.

    Returns None if there are no submitted rows for the requested month.
    """
    import io as _io
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    try:
        ws_sheet = _ws(TABS["MIS_SUBMITTED"])
        all_vals = ws_sheet.get_all_values()
    except Exception:
        return None

    if len(all_vals) < 2:
        return None

    headers = all_vals[0]
    month_col_idx = None
    for ci, h in enumerate(headers):
        if "month" in h.lower():
            month_col_idx = ci
            break

    if month_col_idx is None:
        return None

    data_rows = [row for row in all_vals[1:] if len(row) > month_col_idx and row[month_col_idx].strip() == month_year]
    if not data_rows:
        return None

    wb = Workbook()
    ws = wb.active
    ws.title = f"MIS {month_year}"

    hdr_fill = PatternFill(fill_type="solid", fgColor="002B8F")
    hdr_font = Font(bold=True, color="FFFFFF", size=11)
    alt_fill = PatternFill(fill_type="solid", fgColor="F5F5F5")

    for ci, col_name in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=ci, value=col_name)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for ri, row in enumerate(data_rows, start=2):
        for ci, val in enumerate(row, start=1):
            cell = ws.cell(row=ri, column=ci, value=val)
            if ri % 2 == 0:
                cell.fill = alt_fill
            cell.alignment = Alignment(vertical="center")

    for ci, _ in enumerate(headers, start=1):
        col_letter = get_column_letter(ci)
        max_len = len(str(headers[ci - 1]))
        for row in data_rows:
            if ci - 1 < len(row):
                max_len = max(max_len, len(str(row[ci - 1])))
        ws.column_dimensions[col_letter].width = min(max_len + 3, 45)

    buf = _io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def download_pending_list_excel(pending_rows: list, month_year: str) -> bytes:
    """Build and return an Excel workbook (bytes) listing pending/non-submitted locations."""
    import io as _io
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    col_headers = [
        "Location Code", "Location Name", "Zone",
        "Status", "Completion %", "Days Overdue / Remark",
    ]

    wb = Workbook()
    ws = wb.active
    ws.title = f"Pending {month_year}"

    hdr_fill = PatternFill(fill_type="solid", fgColor="B71C1C")
    hdr_font = Font(bold=True, color="FFFFFF", size=11)

    for ci, col_name in enumerate(col_headers, start=1):
        cell = ws.cell(row=1, column=ci, value=col_name)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    alt_fill = PatternFill(fill_type="solid", fgColor="F5F5F5")

    for ri, rec in enumerate(pending_rows, start=2):
        vals = [
            rec.get("userId", ""),
            rec.get("locName", ""),
            rec.get("zone", ""),
            rec.get("status", ""),
            rec.get("completion_pct", 0),
            rec.get("remark", ""),
        ]
        for ci, val in enumerate(vals, start=1):
            cell = ws.cell(row=ri, column=ci, value=val)
            if ri % 2 == 0:
                cell.fill = alt_fill
            cell.alignment = Alignment(vertical="center")

    for ci, col_name in enumerate(col_headers, start=1):
        max_len = len(col_name)
        for rec in pending_rows:
            vals = [rec.get("userId",""), rec.get("locName",""), rec.get("zone",""),
                    rec.get("status",""), str(rec.get("completion_pct",0)), rec.get("remark","")]
            if ci - 1 < len(vals):
                max_len = max(max_len, len(str(vals[ci - 1])))
        ws.column_dimensions[get_column_letter(ci)].width = min(max_len + 3, 45)

    buf = _io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Analytics helpers ─────────────────────────────────────────────────────────

def _fy_month_years(fy_year: int) -> list:
    """12 'Mon-YYYY' strings for FY starting April fy_year through March fy_year+1."""
    out = []
    for m in range(4, 13):
        out.append(f"{MONTHS_LONG[m - 1][:3]}-{fy_year}")
    for m in range(1, 4):
        out.append(f"{MONTHS_LONG[m - 1][:3]}-{fy_year + 1}")
    return out


_AN_FIELD_LABELS = {
    "f1":  "MS (MT)",
    "f2":  "HSD (MT)",
    "f3":  "Total (MT) incl. Other Products",
    "f4":  "Thruput Target (MT)",
    "f5":  "MEB (Rs in Lakhs)",
    "f6":  "MEB % w.r.t Budget",
    "f7":  "OPEX (Rs/MT)",
    "f8":  "OPEX Target (Rs/MT)",
    "f12": "Electricity Expenditure (Rs in Lakhs)",
    "f21": "SEC (KWH/MT)",
    "f24": "AIM Holds (Nos.)",
    "f26": "Auto-Reconciliation (% of Tanks on Auto Reco)",
    "f38": "CAPEX (Lakhs)",
    "f39": "Capex Target as per AOP (Lakhs)",
    "f46": "MDP Qty Target (KL)",
    "f47": "MDP Qty Actual (KL)",
    "f50": "EBP – Ethanol Blending Percentage",
    "f54": "M&I Index",
    "f55": "PM Percentage",
    "f59": "HSE Index vs Target",
    "f60": "Water Consumed in Month (KL)",
    "f61": "SWC (KL/MT)",
}


@st.cache_data(ttl=300)
def get_compliance_analytics(role: str, zone: str, fy_year: int) -> dict:
    """Fetch submission status for all 12 FY months.

    Returns {month_year: {user_id: {"status", "completion_pct", "submitted_at", "loc"}}}
    """
    months  = _fy_month_years(fy_year)
    locs    = get_all_maker_locations() if role == "Admin" else get_locations_by_zone(zone)
    loc_ids = {l["userId"] for l in locs}
    loc_map  = {l["userId"]: l for l in locs}
    result   = {m: {} for m in months}
    try:
        ws   = _ws(TABS["SUBMISSION_STATUS"])
        rows = ws.get_all_values()
        for row in rows[1:]:
            row    = (row + [""] * 9)[:9]
            uid    = row[0].strip()
            mon    = row[1].strip()
            status = row[2].strip()
            pct    = row[3]
            sub_at = row[4].strip()
            if uid in loc_ids and mon in result:
                result[mon][uid] = {
                    "status":         status or "NOT_STARTED",
                    "completion_pct": float(pct) if pct else 0.0,
                    "submitted_at":   sub_at,
                    "loc":            loc_map.get(uid, {}),
                }
    except Exception:
        pass
    for mon in months:
        for uid in loc_ids:
            if uid not in result[mon]:
                result[mon][uid] = {
                    "status":         "NOT_STARTED",
                    "completion_pct": 0.0,
                    "submitted_at":   "",
                    "loc":            loc_map.get(uid, {}),
                }
    return result


@st.cache_data(ttl=300)
def get_analytics_field_data(role: str, zone: str, fy_year: int) -> list:
    """Fetch approved MIS numeric field values for FY from MIS_Submitted.

    Returns list of dicts: user_id, loc_name, zone_name, month_year + field labels as keys.
    """
    months  = set(_fy_month_years(fy_year))
    locs    = get_all_maker_locations() if role == "Admin" else get_locations_by_zone(zone)
    loc_ids = {l["userId"] for l in locs}
    loc_map  = {l["userId"]: l for l in locs}
    needed   = set(_AN_FIELD_LABELS.values())
    try:
        ws       = _ws(TABS["MIS_SUBMITTED"])
        all_vals = ws.get_all_values()
        if len(all_vals) < 2:
            return []
        headers = all_vals[0]
        col_map = {h: ci for ci, h in enumerate(headers) if h in needed}
        uid_c   = next((i for i, h in enumerate(headers) if h in ("User ID", "user_id")), 0)
        mon_c   = next((i for i, h in enumerate(headers) if "month" in h.lower()), 3)
        out     = []
        for row in all_vals[1:]:
            if len(row) <= max(uid_c, mon_c):
                continue
            uid = row[uid_c].strip()
            mon = row[mon_c].strip()
            if uid not in loc_ids or mon not in months:
                continue
            rec = {
                "user_id":    uid,
                "month_year": mon,
                "loc_name":   loc_map.get(uid, {}).get("locName", uid),
                "zone_name":  loc_map.get(uid, {}).get("zone", ""),
            }
            for label, ci in col_map.items():
                raw = row[ci].strip() if ci < len(row) else ""
                try:
                    rec[label] = float(raw) if raw else None
                except ValueError:
                    rec[label] = None
            out.append(rec)
        return out
    except Exception:
        return []


# ── M&I MIS helpers ───────────────────────────────────────────────────────────

_MI_TAB_HEADERS = {
    "MI_TANK_OUTAGE": [
        "user_id", "month_year", "row_no", "na_flag",
        "tank_no", "other_tank_desc",
        "planned_start", "planned_end", "actual_start", "actual_end",
        "outage_for", "current_status", "saved_at",
    ],
    "MI_MAJOR_REPAIR": [
        "user_id", "month_year", "row_no", "na_flag",
        "tank_no", "other_tank_desc",
        "nature_of_repair", "revenue_capex", "ar_code",
        "current_status", "etc_date", "saved_at",
    ],
    "MI_VRU": [
        "user_id", "month_year", "na_flag", "vru_operational",
        "date_not_operating", "action_taken", "etc_date",
        "ms_vol_recovered_kl",
        "inlet_mfm_start_m3", "inlet_mfm_end_m3",
        "outlet_mfm_start_m3", "outlet_mfm_end_m3",
        "vapour_treated_m3", "voc_value_mgcc", "inlet_emission_mgcc",
        "ms_gasohol_tt_vol_kl", "hsd_tt_vol_kl",
        "ms_gasohol_tw_vol_kl", "hsd_tw_vol_kl",
        "vru_uptime_pct", "saved_at",
    ],
    "MI_AUDIT_2526": [
        "user_id", "month_year", "na_flag",
        "audit_date", "no_recommendations", "no_pending",
        "external_score", "saved_at",
    ],
    "MI_AUDIT_2627": [
        "user_id", "month_year", "na_flag",
        "audit_carried_out", "audit_date", "no_recommendations",
        "no_pending", "external_score", "saved_at",
    ],
    "MI_TECH_AUDIT": [
        "user_id", "month_year", "row_no", "na_flag",
        "audit_date", "no_recommendations", "no_pending",
        "ref_no", "saved_at",
    ],
    "MI_EQUIP_BREAKDOWN": [
        "user_id", "month_year", "row_no", "na_flag",
        "equipment_name", "equipment_name_other", "equipment_details",
        "start_date", "issue", "proposed_date", "actual_end_date",
        "resolution_action", "saved_at",
    ],
    "MI_INT_PIPELINE": [
        "user_id", "month_year", "na_flag",
        "last_ut_date", "last_hydrotest_date", "last_dcvg_date",
        "last_lrut_date", "other_testing", "saved_at",
    ],
    "MI_EXT_PIPELINE": [
        "user_id", "month_year", "na_flag",
        "pipeline_type", "pipeline_details", "length_metres", "product", "size_inch",
        "last_ut_date", "last_hydrotest_date", "last_dcvg_date",
        "last_lrut_date", "other_testing", "saved_at",
    ],
    "MI_TANK_STATUS": [
        "user_id", "month_year", "row_no", "na_flag",
        "zone", "loc_name", "tank_no",
        "cleaning_completed_date", "cleaning_due_date",
        "extension_taken", "extension_efn_no",
        "inspection_date", "inspection_due_date",
        "painting_date", "painting_due_date",
        "tank_status", "tank_status_other",
        "saved_at",
    ],
}


def ensure_mi_tabs():
    """Auto-create any missing M&I worksheet tabs with their headers."""
    for key, headers in _MI_TAB_HEADERS.items():
        _ensure_ws(TABS[key], headers)


_MI_ALL_TABS = (
    "MI_TANK_OUTAGE", "MI_MAJOR_REPAIR", "MI_VRU",
    "MI_AUDIT_2526", "MI_AUDIT_2627", "MI_TECH_AUDIT",
    "MI_EQUIP_BREAKDOWN", "MI_INT_PIPELINE", "MI_EXT_PIPELINE",
    "MI_TANK_STATUS",
)


@st.cache_data(ttl=30)
def check_mi_complete(user_id: str, month_year: str) -> bool:
    """Return True if all 10 M&I MIS tabs have at least one saved row for user+month.

    Cached for 30 s to avoid 10 API calls on every dashboard render.
    """
    for tab_key in _MI_ALL_TABS:
        if not load_mi_data(tab_key, user_id, month_year):
            return False
    return True


_TM_HEADERS = [
    "Sr. No.", "Zone", "Location Code", "Location Name", "SAP Loc Code",
    "Tank No.", "String", "Type", "Year of Commissioning", "Age",
    "Safe Capacity (KL)", "Type2", "SAP Tank No.", "Product",
    "Diameter (m)", "Height (m)",
    "Last Tank Cleaning Date", "Tank Cleaning Due Date",
    "Due for Cleaning 2026-27",
    "Last Comprehensive Inspection Date", "Inspection Due Date",
    "Due for Inspection 2026-27",
    "Last Painted Date", "Due for Painting 2026-27",
    "Cleaning Completed Date", "Cleaning Due Date (Current)",
    "Extension Taken (Yes/No/NA)",
    "Inspection Date (Current)", "Inspection Due Date (Current)",
    "Painting Date (Current)", "Painting Due Date (Current)",
    "Tank Status",
]


@st.cache_data(ttl=7200)
def get_tank_master() -> dict:
    """Return {location_code: [sap_tank_no, ...]} — tries Google Sheet first, then Excel.

    Uses _api_call so 429 quota errors are retried with exponential backoff
    (up to 6 attempts, max ~63 s total) instead of silently returning {}.
    TTL raised to 2 h — Tank Master changes only when admin re-seeds.
    """
    # Try Google Sheet (with retry on 429)
    try:
        ws   = _ws(TABS["TANK_MASTER"])
        rows = _api_call(ws.get_all_values)
        if len(rows) >= 2:
            hdr = rows[0]
            try:
                loc_idx  = hdr.index("Location Code")
                tank_idx = hdr.index("SAP Tank No.")
            except ValueError:
                loc_idx, tank_idx = 2, 12
            result: dict = {}
            for row in rows[1:]:
                row = (row + [""] * max(loc_idx, tank_idx))
                loc_code = str(row[loc_idx]).strip()
                tank_no  = str(row[tank_idx]).strip()
                if not loc_code or loc_code.lower() in ("none", ""):
                    continue
                if not tank_no or tank_no.lower() in ("none", ""):
                    continue
                result.setdefault(loc_code, [])
                if tank_no not in result[loc_code]:
                    result[loc_code].append(tank_no)
            if result:
                return result
    except Exception:
        pass

    # Fallback: local Excel
    try:
        import openpyxl
        path = r"D:\SHOAIB\VS CODE PROJECTS\SOD MIS\M&I Separate Block.xlsx"
        wb_xl = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws_xl = wb_xl["Tank Master"]
        result = {}
        for row in ws_xl.iter_rows(min_row=2, values_only=True):
            if not row or len(row) <= 12:
                continue
            loc_code = str(row[2]).strip() if row[2] is not None else ""
            tank_no  = str(row[12]).strip() if row[12] is not None else ""
            if not loc_code or loc_code.lower() in ("none", ""):
                continue
            if not tank_no or tank_no.lower() in ("none", ""):
                continue
            result.setdefault(loc_code, [])
            if tank_no not in result[loc_code]:
                result[loc_code].append(tank_no)
        return result
    except Exception:
        return {}


def sync_tank_master_to_sheet() -> dict:
    """Read Tank Master from local Excel, map zones from UserAccess, write to Google Sheet.

    Returns {"ok": bool, "rows": int, "msg": str}
    """
    try:
        import openpyxl
        from datetime import datetime as _dt

        # Build zone lookup: loc_code_str → full_zone_name from UserAccess
        zone_map: dict = {}
        try:
            ua_ws  = _ws(TABS["USER_ACCESS"])
            ua_rows = ua_ws.get_all_values()
            for row in ua_rows[1:]:
                row = (row + [""] * 6)[:6]
                loc_code, _loc_name, zone = row[0].strip(), row[1].strip(), row[2].strip()
                if loc_code and zone:
                    zone_map[loc_code] = zone
        except Exception:
            pass

        # Read Excel Tank Master
        path  = r"D:\SHOAIB\VS CODE PROJECTS\SOD MIS\M&I Separate Block.xlsx"
        wb_xl = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws_xl = wb_xl["Tank Master"]

        data_rows = []
        for ri, row in enumerate(ws_xl.iter_rows(min_row=2, values_only=True), 1):
            if not row or all(v is None for v in row):
                continue
            # Convert row to list; stringify dates and numbers
            cells = []
            for v in row:
                if v is None:
                    cells.append("")
                elif isinstance(v, _dt):
                    cells.append(v.strftime("%d/%m/%Y"))
                else:
                    cells.append(str(v))

            # Pad / truncate to 32 columns
            while len(cells) < 32:
                cells.append("")
            cells = cells[:32]

            # Replace abbreviated zone (index 1) with full name from UserAccess
            loc_code_str = cells[2].strip()
            if loc_code_str and loc_code_str in zone_map:
                cells[1] = zone_map[loc_code_str]

            data_rows.append(cells)

        # Write to Google Sheet (clear + rewrite)
        tm_ws = _ensure_ws(TABS["TANK_MASTER"], _TM_HEADERS)
        all_existing = tm_ws.get_all_values()
        if len(all_existing) > 1:
            tm_ws.delete_rows(2, len(all_existing) - 1)

        # Batch write (500 rows at a time)
        BATCH = 500
        for start in range(0, len(data_rows), BATCH):
            tm_ws.append_rows(data_rows[start:start + BATCH], value_input_option="USER_ENTERED")

        get_tank_master.clear()
        return {"ok": True, "rows": len(data_rows),
                "msg": f"Tank Master synced: {len(data_rows)} rows written to Google Sheet."}
    except Exception as exc:
        return {"ok": False, "rows": 0, "msg": str(exc)}


_TM_DATE_COLS = {
    "Last Tank Cleaning Date", "Tank Cleaning Due Date",
    "Last Comprehensive Inspection Date", "Inspection Due Date",
    "Last Painted Date",
    "Cleaning Completed Date", "Cleaning Due Date (Current)",
    "Inspection Date (Current)", "Inspection Due Date (Current)",
    "Painting Date (Current)", "Painting Due Date (Current)",
}


def _normalize_tm_date(val: str) -> str:
    """Convert DD.MM.YY or DD.MM.YYYY → DD/MM/YYYY; leave other values unchanged."""
    import re as _re
    if not val or not isinstance(val, str):
        return val
    v = val.strip()
    m = _re.fullmatch(r'(\d{1,2})\.(\d{1,2})\.(\d{2})', v)
    if m:
        d, mo, y = m.groups()
        return f"{d.zfill(2)}/{mo.zfill(2)}/20{y}"
    m = _re.fullmatch(r'(\d{1,2})\.(\d{1,2})\.(\d{4})', v)
    if m:
        d, mo, y = m.groups()
        return f"{d.zfill(2)}/{mo.zfill(2)}/{y}"
    m = _re.fullmatch(r'(\d{1,2})/(\d{1,2})/(\d{4})', v)
    if m:
        d, mo, y = m.groups()
        return f"{d.zfill(2)}/{mo.zfill(2)}/{y}"
    return val


@st.cache_data(ttl=7200, show_spinner=False)
def get_full_tank_master_excel(
    location_code: str | None = None,
    zone: str | None = None,
) -> bytes:
    """Return xlsx bytes for Tank Master filtered by location_code, zone, or all rows."""
    import io as _io
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    try:
        ws_sheet = _ws(TABS["TANK_MASTER"])
        all_rows = _api_call(ws_sheet.get_all_values)
    except Exception as exc:
        raise ValueError(f"Cannot read Tank Master: {exc}") from exc

    if len(all_rows) < 2:
        raise ValueError("Tank Master sheet has no data rows.")

    hdr  = all_rows[0]
    data = all_rows[1:]

    try:
        loc_idx = hdr.index("Location Code")
    except ValueError:
        loc_idx = 2
    try:
        zone_idx = hdr.index("Zone")
    except ValueError:
        zone_idx = 1

    if location_code:
        data = [r for r in data if (r + [""] * (loc_idx + 1))[loc_idx].strip() == location_code]
    elif zone:
        data = [r for r in data if (r + [""] * (zone_idx + 1))[zone_idx].strip() == zone]

    # Identify which column indices are date columns
    date_col_indices = {ci for ci, h in enumerate(hdr) if h in _TM_DATE_COLS}

    wb  = Workbook()
    ws1 = wb.active
    ws1.title = "Tank Master"

    BLUE = PatternFill("solid", fgColor="0033A0")
    W    = Font(bold=True, color="FFFFFF", size=10)
    NM   = Font(size=10)
    CTR  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LFT  = Alignment(horizontal="left",   vertical="center", wrap_text=False)

    n_cols = len(hdr)
    for ci, h in enumerate(hdr, 1):
        c = ws1.cell(row=1, column=ci, value=h)
        c.font = W; c.fill = BLUE; c.alignment = CTR

    for ri, row in enumerate(data, 2):
        row_p = (row + [""] * n_cols)[:n_cols]
        for ci, val in enumerate(row_p, 1):
            if (ci - 1) in date_col_indices:
                val = _normalize_tm_date(val)
            c = ws1.cell(row=ri, column=ci, value=val)
            c.font = NM; c.alignment = LFT

    for ci, h in enumerate(hdr, 1):
        ws1.column_dimensions[get_column_letter(ci)].width = min(len(str(h)) + 4, 35)

    ws1.freeze_panes = "A2"
    ws1.row_dimensions[1].height = 30

    buf = _io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def get_approved_mis_excel(
    zone: str | None = None,
    month_year: str | None = None,
) -> bytes:
    """Return xlsx bytes from MIS_SUBMITTED optionally filtered by zone and/or month."""
    import io as _io
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    try:
        ws_sheet = _ws(TABS["MIS_SUBMITTED"])
        all_rows = _api_call(ws_sheet.get_all_values)
    except Exception as exc:
        raise ValueError(f"Cannot read MIS_SUBMITTED: {exc}") from exc

    if len(all_rows) < 2:
        raise ValueError("No approved MIS submissions found.")

    hdr  = all_rows[0]
    data = all_rows[1:]

    try:
        zone_idx = hdr.index("Zone")
    except ValueError:
        zone_idx = 2
    try:
        mon_idx = hdr.index("Month-Year")
    except ValueError:
        mon_idx = 3

    if zone:
        data = [r for r in data if (r + [""] * (zone_idx + 1))[zone_idx].strip() == zone]
    if month_year:
        data = [r for r in data if (r + [""] * (mon_idx + 1))[mon_idx].strip() == month_year]

    if not data:
        raise ValueError("No approved MIS records found for the selected filters.")

    wb  = Workbook()
    ws1 = wb.active
    ws1.title = "Approved MIS"

    BLUE = PatternFill("solid", fgColor="0033A0")
    W    = Font(bold=True, color="FFFFFF", size=10)
    NM   = Font(size=10)
    CTR  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LFT  = Alignment(horizontal="left",   vertical="center", wrap_text=False)

    n_cols = len(hdr)
    for ci, h in enumerate(hdr, 1):
        c = ws1.cell(row=1, column=ci, value=h)
        c.font = W; c.fill = BLUE; c.alignment = CTR

    for ri, row in enumerate(data, 2):
        row_p = (row + [""] * n_cols)[:n_cols]
        for ci, val in enumerate(row_p, 1):
            c = ws1.cell(row=ri, column=ci, value=val)
            c.font = NM; c.alignment = LFT

    for ci, h in enumerate(hdr, 1):
        ws1.column_dimensions[get_column_letter(ci)].width = min(len(str(h)) + 4, 40)

    ws1.freeze_panes = "A2"
    ws1.row_dimensions[1].height = 30

    buf = _io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


@st.cache_data(ttl=300, show_spinner=False)
def load_mi_data(tab_key: str, user_id: str, month_year: str) -> list:
    """Return list of row-dicts for user+month from an M&I tab (5-min cache)."""
    try:
        headers = _MI_TAB_HEADERS[tab_key]
        ws      = _ensure_ws(TABS[tab_key], headers)
        all_rows = _api_call(ws.get_all_values)
        if len(all_rows) < 2:
            return []
        hdr = all_rows[0]
        try:
            uid_idx = hdr.index("user_id")
            mon_idx = hdr.index("month_year")
        except ValueError:
            return []
        result = []
        for row in all_rows[1:]:
            row = (row + [""] * len(hdr))[:len(hdr)]
            if row[uid_idx].strip() != user_id or row[mon_idx].strip() != month_year:
                continue
            result.append({hdr[i]: row[i] for i in range(len(hdr))})
        return result
    except Exception:
        return []


def save_mi_data(tab_key: str, user_id: str, month_year: str, rows: list) -> dict:
    """Replace all rows for user+month in an M&I tab with the provided rows list."""
    try:
        headers  = _MI_TAB_HEADERS[tab_key]
        ws       = _ensure_ws(TABS[tab_key], headers)
        all_rows = _api_call(ws.get_all_values)
        hdr = all_rows[0] if all_rows else headers
        try:
            uid_idx = hdr.index("user_id")
            mon_idx = hdr.index("month_year")
        except ValueError:
            uid_idx, mon_idx = 0, 1

        to_del = [i + 2 for i, row in enumerate(all_rows[1:])
                  if (row + [""] * len(hdr))[uid_idx].strip() == user_id
                  and (row + [""] * len(hdr))[mon_idx].strip() == month_year]
        for idx in reversed(to_del):
            _api_call(ws.delete_rows, idx)

        now_str = datetime.now().isoformat()
        for rec in rows:
            out_row = []
            for col in headers:
                if col == "user_id":
                    out_row.append(user_id)
                elif col == "month_year":
                    out_row.append(month_year)
                elif col == "saved_at":
                    out_row.append(now_str)
                else:
                    out_row.append(str(rec.get(col, "") or ""))
            _api_call(ws.append_row, out_row, value_input_option="RAW")

        load_mi_data.clear()  # invalidate read cache so next load fetches fresh data
        audit_log(user_id, f"SaveMI {tab_key}", f"month={month_year} rows={len(rows)}")
        return {"ok": True, "rows": len(rows)}
    except Exception as e:
        return {"ok": False, "msg": str(e)}
