"""HPCL SOD MIS — Email notifications via Microsoft Outlook COM.

Emails are sent through the Outlook desktop application running on the
same Windows machine as this Streamlit app.  No SMTP credentials needed —
Outlook uses the currently signed-in HPCL account (shoaibrehman@hpcl.in).

Requirements (local only):
  • Microsoft Outlook must be installed and OPEN on this PC
  • shoaibrehman@hpcl.in must be the active Outlook account
  • pywin32 installed:  pip install pywin32

When running on Streamlit Cloud (Linux), email is not available and the
UI shows an informational notice — this is expected.
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
    "North Zone (NZ)":      {"to": "ajaygr@hpcl.in",                       "cc": "NZ.OND.IC@hpcl.in"},
    "Patna Zone":           {"to": "ajaisingh@hpcl.in",                    "cc": "dastidar@hpcl.in"},
    "South Central Zone":   {"to": "suryabv@hpcl.in",                      "cc": "SCRZ.OND.IC@hpcl.in"},
    "South Zone":           {"to": "venkates@hpcl.in",                     "cc": "SZ.OND.IC@hpcl.in"},
    "West Zone (WZ)":       {"to": "ntgajbiye@hpcl.in",                    "cc": "WZ.OND.IC@hpcl.in"},
}


# ── Outlook availability check ────────────────────────────────────────────────

def email_configured() -> tuple:
    """Return (True, '') if Outlook COM is reachable, else (False, reason)."""
    if platform.system() != "Windows":
        return False, "local_only"   # sentinel — not an error, just local feature
    try:
        import win32com.client  # noqa: F401
        import pythoncom        # noqa: F401
    except ImportError:
        return False, "pywin32 not installed — run: pip install pywin32"
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        try:
            win32com.client.Dispatch("Outlook.Application")
        finally:
            pythoncom.CoUninitialize()
        return True, ""
    except Exception as exc:
        return False, f"Outlook is not running or not installed: {exc}"


# backward-compat alias
def outlook_available() -> tuple:
    return email_configured()


# ── HTML email builder ────────────────────────────────────────────────────────

_DEFAULT_INTRO = (
    "This is a reminder that the MIS submission for <strong>{month_year}</strong> "
    "is <strong>pending</strong> for the following locations in your zone. "
    "Please ensure submissions are completed at the earliest."
)

def _build_email_html(zone_name: str, month_year: str,
                      pending_locs: list, due_date: date,
                      custom_intro: str = "") -> str:
    today   = date.today()
    overdue = today > due_date
    due_str = due_date.strftime("%d %b %Y")

    intro_html = (
        custom_intro.strip()
        if custom_intro.strip()
        else _DEFAULT_INTRO.format(month_year=month_year)
    )

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
        {intro_html}
      </p>

      {overdue_banner}

      <table style="width:100%;border-collapse:collapse;margin-top:16px;
                    border-radius:8px;overflow:hidden;">
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


def build_preview_html(zone_name: str, month_year: str,
                       pending_locs: list, due_date: date,
                       custom_intro: str = "") -> str:
    """Public wrapper — returns full HTML string for preview in the app."""
    return _build_email_html(zone_name, month_year, pending_locs, due_date, custom_intro)


def get_zone_recipients(zone_name: str) -> dict:
    """Return {"to": ..., "cc": ..., "bcc": ...} for a zone, or empty strings."""
    c = ZONE_EMAIL_MAP.get(zone_name, {})
    return {
        "to":  c.get("to", ""),
        "cc":  c.get("cc", ""),
        "bcc": "; ".join(BCC_EMAILS),
    }


# ── Single-zone sender ────────────────────────────────────────────────────────

def send_zone_reminder(zone_name: str, month_year: str,
                       pending_locs: list, due_date: date,
                       custom_intro: str = "",
                       test_mode: bool = False,
                       test_email: str = "") -> dict:
    """Send one reminder email for a zone through Outlook COM.
    Falls back to saving as Outlook draft if direct Send() fails.

    test_mode=True overrides all recipients with test_email only.
    """
    contacts = ZONE_EMAIL_MAP.get(zone_name)
    if not contacts:
        return {"ok": False, "msg": f"No email contacts configured for: {zone_name}"}

    try:
        import pythoncom
        import win32com.client as win32
    except ImportError:
        return {"ok": False, "msg": "pywin32 not installed. Run: pip install pywin32"}

    try:
        today   = date.today()
        overdue = today > due_date
        pfx     = "[TEST] " if test_mode else ""
        subject = (
            f"{pfx}OVERDUE — MIS Submission Pending | {zone_name} | {month_year}"
            if overdue else
            f"{pfx}Reminder — MIS Submission Pending | {zone_name} | {month_year}"
        )
        html_body = _build_email_html(zone_name, month_year, pending_locs,
                                      due_date, custom_intro)

        if test_mode:
            to_str  = test_email or SENDER_EMAIL
            cc_str  = ""
            bcc_str = ""
        else:
            to_str  = contacts["to"]
            cc_str  = contacts["cc"]
            bcc_str = "; ".join(BCC_EMAILS)

        pythoncom.CoInitialize()
        try:
            outlook   = win32.Dispatch("Outlook.Application")
            mail_item = outlook.CreateItem(0)   # 0 = olMailItem

            mail_item.To       = to_str
            if cc_str:
                mail_item.CC   = cc_str
            if bcc_str:
                mail_item.BCC  = bcc_str
            mail_item.Subject  = subject
            mail_item.HTMLBody = html_body

            try:
                mail_item.Send()
                return {"ok": True, "mode": "sent"}
            except Exception as send_exc:
                try:
                    mail_item.Save()
                    return {
                        "ok": True,
                        "mode": "draft",
                        "msg": f"Saved as Outlook draft (send error: {send_exc})",
                    }
                except Exception as draft_exc:
                    return {"ok": False, "msg": f"Send and draft both failed: {draft_exc}"}
        finally:
            pythoncom.CoUninitialize()

    except Exception as exc:
        return {"ok": False, "msg": f"Outlook error: {exc}"}


# ── All-zones batch sender ────────────────────────────────────────────────────

def send_all_reminders(month_year: str,
                       all_location_rows: list,
                       due_date: date,
                       custom_intro: str = "",
                       test_mode: bool = False,
                       test_email: str = "") -> dict:
    """Send reminder emails for every zone that has non-submitted locations.

    test_mode=True sends a single combined email to test_email instead of
    sending to all zone contacts.
    """
    ok, err = email_configured()
    if not ok:
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
    zones_to_send = sorted(pending_by_zone.items())

    if test_mode:
        # Send only the first zone's email to the test address
        first_zone, first_locs = zones_to_send[0]
        res = send_zone_reminder(
            first_zone, month_year, first_locs, due_date,
            custom_intro=custom_intro,
            test_mode=True,
            test_email=test_email or SENDER_EMAIL,
        )
        if res["ok"]:
            sent_count = 1
            mode = res.get("mode", "sent")
            msg = (
                f"[TEST] Email {mode} to {test_email or SENDER_EMAIL} "
                f"(sample: {first_zone}, {len(first_locs)} location(s)). "
                f"In production this would send to {len(zones_to_send)} zone(s)."
            )
        else:
            failed_count = 1
            errors = [res.get("msg", "")]
            msg = f"[TEST] Send failed: {errors[0]}"
        return {"ok": res["ok"], "sent": sent_count, "failed": failed_count,
                "errors": errors, "msg": msg}

    for zone_name, locs in zones_to_send:
        if zone_name not in ZONE_EMAIL_MAP:
            errors.append(f"{zone_name}: no email config")
            failed_count += 1
            continue
        res = send_zone_reminder(zone_name, month_year, locs, due_date,
                                 custom_intro=custom_intro)
        if res["ok"]:
            sent_count += 1
        else:
            failed_count += 1
            errors.append(f"{zone_name}: {res.get('msg', 'unknown')}")

    all_ok = (failed_count == 0)
    msg = (
        f"Emails sent to {sent_count} zone(s) via Outlook ({SENDER_EMAIL})."
        if all_ok else
        f"Sent: {sent_count}. Failed: {failed_count}. " + "; ".join(errors)
    )
    return {
        "ok": all_ok,
        "sent": sent_count,
        "failed": failed_count,
        "errors": errors,
        "msg": msg,
    }
