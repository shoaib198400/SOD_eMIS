"""HPCL SOD MIS — Email notifications via SMTP (cloud-compatible).

Works on Streamlit Cloud (Linux) and local Windows alike.
Credentials are read from st.secrets["email"]:

  [email]
  smtp_host     = "smtp.gmail.com"
  smtp_port     = 587
  smtp_user     = "your-sender@gmail.com"
  smtp_password = "your-app-password"
  sender_name   = "HPCL SOD MIS"      # optional display name
"""

import smtplib
import ssl
from datetime import date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import streamlit as st

# ── Constants ─────────────────────────────────────────────────────────────────

SENDER_EMAIL = "shoaibrehman@hpcl.in"   # display / Reply-To address

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


# ── SMTP credential helper ─────────────────────────────────────────────────────

def _smtp_cfg() -> dict | None:
    """Return SMTP config dict from st.secrets, or None if not configured.

    Secrets keys (under [email]):
      smtp_host     – SMTP server (e.g. smtp.office365.com)
      smtp_port     – default 587
      smtp_user     – login username (usually the sending email address)
      smtp_password – SMTP/app password
      from_email    – optional From address; defaults to smtp_user
      sender_name   – display name; default "HPCL SOD MIS"
    """
    try:
        cfg = st.secrets.get("email", {})
        if cfg.get("smtp_host") and cfg.get("smtp_user") and cfg.get("smtp_password"):
            user = str(cfg["smtp_user"])
            return {
                "host":       str(cfg["smtp_host"]),
                "port":       int(cfg.get("smtp_port", 587)),
                "user":       user,
                "password":   str(cfg["smtp_password"]),
                "from_email": str(cfg.get("from_email", user)),
                "name":       str(cfg.get("sender_name", "HPCL SOD MIS")),
            }
    except Exception:
        pass
    return None


_O365_TEMPLATE = (
    '[email]\n'
    'smtp_host     = "smtp.office365.com"\n'
    'smtp_port     = 587\n'
    'smtp_user     = "shoaibrehman@hpcl.in"\n'
    'smtp_password = "your-hpcl-password"\n'
    'from_email    = "shoaibrehman@hpcl.in"\n'
    'sender_name   = "HPCL SOD MIS"'
)


def email_configured() -> tuple:
    """Return (True, '') if SMTP credentials are present, else (False, reason)."""
    cfg = _smtp_cfg()
    if cfg:
        return True, ""
    return False, _O365_TEMPLATE


# kept for backward-compat with any existing callers
def outlook_available() -> tuple:
    return email_configured()


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
    """Send one reminder email for a zone via SMTP."""
    contacts = ZONE_EMAIL_MAP.get(zone_name)
    if not contacts:
        return {"ok": False, "msg": f"No email contacts configured for: {zone_name}"}

    cfg = _smtp_cfg()
    if not cfg:
        return {"ok": False, "msg": "Email not configured. Add [email] section to Streamlit secrets."}

    try:
        today   = date.today()
        overdue = today > due_date

        to_list  = [e.strip() for e in contacts["to"].split(";") if e.strip()]
        cc_list  = [e.strip() for e in contacts["cc"].split(";") if e.strip()]
        bcc_list = list(BCC_EMAILS)
        all_rcpt = to_list + cc_list + bcc_list

        subject = (
            f"OVERDUE — MIS Submission Pending | {zone_name} | {month_year}"
            if overdue else
            f"Reminder — MIS Submission Pending | {zone_name} | {month_year}"
        )

        msg = MIMEMultipart("alternative")
        msg["Subject"]  = subject
        msg["From"]     = f"{cfg['name']} <{cfg['from_email']}>"
        msg["To"]       = "; ".join(to_list)
        msg["CC"]       = "; ".join(cc_list)
        msg["Reply-To"] = cfg["from_email"]

        msg.attach(MIMEText(
            _build_email_html(zone_name, month_year, pending_locs, due_date),
            "html",
            "utf-8",
        ))

        ctx = ssl.create_default_context()
        with smtplib.SMTP(cfg["host"], cfg["port"]) as server:
            server.ehlo()
            server.starttls(context=ctx)
            server.login(cfg["user"], cfg["password"])
            server.sendmail(cfg["from_email"], all_rcpt, msg.as_string())

        return {"ok": True}

    except Exception as exc:
        return {"ok": False, "msg": str(exc)}


# ── All-zones batch sender ────────────────────────────────────────────────────

def send_all_reminders(month_year: str,
                       all_location_rows: list,
                       due_date: date) -> dict:
    """Send reminder emails for every zone that has non-submitted locations."""
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
    msg = (
        f"Reminder emails sent to {sent_count} zone(s) via SMTP."
        if all_ok else
        f"Sent: {sent_count}. Failed: {failed_count}. Errors: " + "; ".join(errors)
    )
    return {
        "ok": all_ok,
        "sent": sent_count,
        "failed": failed_count,
        "errors": errors,
        "msg": msg,
    }
