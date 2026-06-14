"""HPCL SOD MIS — Email notifications via local Microsoft Outlook (Phase 9).

Emails are sent through the Outlook desktop application running on the
same Windows machine as this Streamlit app. No passwords are required —
Outlook uses its existing HPCL account authentication.

Requirements:
  • Microsoft Outlook must be installed and OPEN on this PC
  • shoaibrehman@hpcl.in must be configured as an account in Outlook
  • pywin32 installed:  pip install pywin32
"""

import platform
from datetime import date

# ── Constants ─────────────────────────────────────────────────────────────────

SENDER_EMAIL = "shoaibrehman@hpcl.in"

BCC_EMAILS = [
    "SOD.OPNS.HQO@hpcl.in",
    "bhsgk@hpcl.in",
    "shubham.tayal@hpcl.in",
]

# ── Zone → recipient map ──────────────────────────────────────────────────────

ZONE_EMAIL_MAP = {
    "Bengaluru Zone":       {"to": "brijeshkumar@hpcl.in",                 "cc": "BLR.OND.IC@hpcl.in"},
    "Bhopal Zone":          {"to": "agajare@hpcl.in;twinacore@hpcl.in",    "cc": "CZ.OND.IC@hpcl.in"},
    "Bhubneshwar Zone":     {"to": "smarak.lenka@hpcl.in",                 "cc": "ECZ.OND.IC@hpcl.in"},
    "Chandigarh Zone":      {"to": "haroonhamid@hpcl.in",                  "cc": "NFZ.OND.IC@hpcl.in"},
    "Cochin Zone":          {"to": "kathir@hpcl.in",                       "cc": "kbanothu@hpcl.in"},
    "East Zone":            {"to": "sray@hpcl.in",                         "cc": "EZ.OND.IC@hpcl.in"},
    "Guwahati Zone":        {"to": "lodyuo@hpcl.in",                       "cc": "gurubachansingha@hpcl.in"},
    "Jaipur Zone":          {"to": "rjprasad@hpcl.in",                     "cc": "NWF.OND.IC@hpcl.in"},
    "Noida (UP-West) Zone": {"to": "aradhnat@hpcl.in",                     "cc": "chraghu@hpcl.in"},
    "North Central Zone":   {"to": "rvpandey@hpcl.in",                     "cc": "adeshmukh@hpcl.in"},
    "North West Zone":      {"to": "sanjaykdewangan@hpcl.in",              "cc": "NWZ.OND.IC@hpcl.in"},
    "North Zone (NZ)":      {"to": "ajaygr@hpcl.in",                      "cc": "NZ.OND.IC@hpcl.in"},
    "Patna Zone":           {"to": "ajaisingh@hpcl.in",                    "cc": "dastidar@hpcl.in"},
    "South Central Zone":   {"to": "suryabv@hpcl.in",                      "cc": "SCRZ.OND.IC@hpcl.in"},
    "South Zone":           {"to": "venkates@hpcl.in",                     "cc": "SZ.OND.IC@hpcl.in"},
    "West Zone (WZ)":       {"to": "ntgajbiye@hpcl.in",                    "cc": "WZ.OND.IC@hpcl.in"},
}


# ── Outlook availability check ────────────────────────────────────────────────

def outlook_available() -> tuple:
    """Return (True, '') if Outlook COM is reachable, else (False, reason_str)."""
    if platform.system() != "Windows":
        return False, (
            "Outlook email requires Windows. "
            "Run the app locally on your Windows PC with Outlook open."
        )
    try:
        import win32com.client  # noqa: F401
        import pythoncom        # noqa: F401
    except ImportError:
        return False, (
            "pywin32 is not installed. "
            "Run:  pip install pywin32  then restart the app."
        )
    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        try:
            win32com.client.Dispatch("Outlook.Application")
        finally:
            pythoncom.CoUninitialize()
    except Exception as exc:
        return False, f"Outlook is not running or not installed: {exc}"
    return True, ""


def _get_sender_account(outlook):
    """Return the Outlook Account object for SENDER_EMAIL, or None."""
    try:
        accounts = outlook.GetNamespace("MAPI").Accounts
        for i in range(1, accounts.Count + 1):
            acct = accounts.Item(i)
            if acct.SmtpAddress.lower() == SENDER_EMAIL.lower():
                return acct
    except Exception:
        pass
    return None


# ── HTML email builder ────────────────────────────────────────────────────────

def _build_email_html(zone_name: str, month_year: str,
                      pending_locs: list, due_date: date) -> str:
    today   = date.today()
    overdue = today > due_date
    due_str = due_date.strftime("%d %b %Y")

    rows_html = ""
    for loc in pending_locs:
        uid    = loc.get("userId", "")
        name   = loc.get("locName", "")
        status = loc.get("status", "NOT_STARTED").replace("_", " ").title()
        pct    = int(float(loc.get("completion_pct", 0)))
        bg     = "#fff8f8" if overdue else "#ffffff"
        rows_html += (
            f"<tr style='background:{bg};'>"
            f"<td style='padding:8px 12px;border-bottom:1px solid #e8ecf4;font-size:13px;'>{uid}</td>"
            f"<td style='padding:8px 12px;border-bottom:1px solid #e8ecf4;font-size:13px;'>{name}</td>"
            f"<td style='padding:8px 12px;border-bottom:1px solid #e8ecf4;font-size:13px;'>{status}</td>"
            f"<td style='padding:8px 12px;border-bottom:1px solid #e8ecf4;font-size:13px;"
            f"text-align:center;'>{pct}%</td>"
            f"</tr>"
        )

    overdue_banner = ""
    if overdue:
        overdue_banner = (
            f"<div style='background:#fdecea;border:2px solid #e53935;border-radius:8px;"
            f"padding:12px 18px;margin:16px 0;color:#b71c1c;font-weight:700;font-size:14px;'>"
            f"&#9888;&nbsp; OVERDUE — Submission deadline was {due_str}. "
            f"Immediate action is required."
            f"</div>"
        )

    return f"""<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;margin:0;padding:0;background:#f4f6fb;">
  <div style="max-width:720px;margin:30px auto;background:white;border-radius:14px;
              box-shadow:0 2px 12px rgba(0,0,0,0.10);overflow:hidden;">

    <div style="background:#002B8F;padding:22px 30px;">
      <div style="color:white;font-size:20px;font-weight:700;letter-spacing:0.5px;">
        HPCL SOD &mdash; MIS Portal
      </div>
      <div style="color:#a8bfe8;font-size:12px;margin-top:4px;">
        Supply, Operations &amp; Distribution &nbsp;&middot;&nbsp; HQO Notification
      </div>
    </div>

    <div style="padding:28px 30px;">
      <p style="font-size:15px;color:#333;margin-top:0;">
        Dear <strong>{zone_name} Team</strong>,
      </p>
      <p style="font-size:14px;color:#444;line-height:1.8;">
        This is a reminder that the MIS submission for <strong>{month_year}</strong>
        is <strong>pending</strong> for the following locations in your zone.
        Please ensure submissions are completed at the earliest.
      </p>

      {overdue_banner}

      <table style="width:100%;border-collapse:collapse;margin-top:16px;border-radius:8px;overflow:hidden;">
        <thead>
          <tr style="background:#002B8F;">
            <th style="padding:10px 12px;color:white;text-align:left;font-size:13px;">Location Code</th>
            <th style="padding:10px 12px;color:white;text-align:left;font-size:13px;">Location Name</th>
            <th style="padding:10px 12px;color:white;text-align:left;font-size:13px;">Status</th>
            <th style="padding:10px 12px;color:white;text-align:center;font-size:13px;">Completion</th>
          </tr>
        </thead>
        <tbody>
          {rows_html}
        </tbody>
      </table>

      <p style="font-size:14px;color:#333;margin-top:22px;line-height:1.8;">
        Submission deadline: <strong>{due_str}</strong><br>
        Please log in to the HPCL SOD MIS Portal and complete your submission.
      </p>

      <p style="font-size:12px;color:#aaa;border-top:1px solid #eee;
                padding-top:16px;margin-top:24px;line-height:1.7;">
        This is an auto-generated message from the HPCL SOD MIS Portal.<br>
        Sent by: {SENDER_EMAIL} &nbsp;&middot;&nbsp; Supply, Operations &amp; Distribution HQO<br>
        Please do not reply directly to this email.
      </p>
    </div>
  </div>
</body>
</html>"""


# ── Single-zone sender ────────────────────────────────────────────────────────

def send_zone_reminder(zone_name: str, month_year: str,
                       pending_locs: list, due_date: date) -> dict:
    """Send one reminder email for a zone through Outlook COM."""
    contacts = ZONE_EMAIL_MAP.get(zone_name)
    if not contacts:
        return {"ok": False, "msg": f"No email contacts configured for: {zone_name}"}

    try:
        import win32com.client
        import pythoncom
    except ImportError:
        return {"ok": False, "msg": "pywin32 is not installed. Run: pip install pywin32"}

    try:
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")

            mail = outlook.CreateItem(0)  # 0 = olMailItem

            # Use shoaibrehman@hpcl.in account if present
            acct = _get_sender_account(outlook)
            if acct:
                mail.SendUsingAccount = acct

            to_list  = [e.strip() for e in contacts["to"].split(";") if e.strip()]
            cc_list  = [e.strip() for e in contacts["cc"].split(";") if e.strip()]
            bcc_list = list(BCC_EMAILS)

            mail.To      = "; ".join(to_list)
            mail.CC      = "; ".join(cc_list)
            mail.BCC     = "; ".join(bcc_list)

            today   = date.today()
            overdue = today > due_date
            mail.Subject = (
                f"⚠ OVERDUE — MIS Submission Pending | {zone_name} | {month_year}"
                if overdue else
                f"Reminder — MIS Submission Pending | {zone_name} | {month_year}"
            )
            mail.HTMLBody = _build_email_html(zone_name, month_year, pending_locs, due_date)
            mail.Send()
        finally:
            pythoncom.CoUninitialize()
        return {"ok": True}

    except Exception as exc:
        return {"ok": False, "msg": str(exc)}


# ── All-zones batch sender ────────────────────────────────────────────────────

def send_all_reminders(month_year: str,
                       all_location_rows: list,
                       due_date: date) -> dict:
    """Send reminder emails for every zone that has non-submitted locations."""
    avail, err = outlook_available()
    if not avail:
        return {"ok": False, "sent": 0, "failed": 0, "errors": [], "msg": err}

    pending_by_zone: dict = {}
    for loc in all_location_rows:
        if loc.get("status") != "SUBMITTED":
            zone = loc.get("zone", "Unknown")
            pending_by_zone.setdefault(zone, []).append(loc)

    if not pending_by_zone:
        return {
            "ok": True, "sent": 0, "failed": 0, "errors": [],
            "msg": "All locations have submitted. No reminder emails sent.",
        }

    sent_count, failed_count, errors = 0, 0, []
    for zone_name, locs in sorted(pending_by_zone.items()):
        if zone_name not in ZONE_EMAIL_MAP:
            errors.append(f"{zone_name}: no email config")
            failed_count += 1
            continue
        res = send_zone_reminder(zone_name, month_year, locs, due_date)
        if res["ok"]:
            sent_count += 1
        else:
            failed_count += 1
            errors.append(f"{zone_name}: {res.get('msg', 'unknown error')}")

    all_ok = (failed_count == 0)
    if all_ok:
        msg = f"Reminder emails sent to {sent_count} zone(s) from {SENDER_EMAIL}."
    else:
        msg = (
            f"Sent: {sent_count} zone(s). Failed: {failed_count}. "
            + "Errors: " + "; ".join(errors)
        )
    return {
        "ok": all_ok,
        "sent": sent_count,
        "failed": failed_count,
        "errors": errors,
        "msg": msg,
    }
