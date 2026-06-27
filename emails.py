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


# ── Location credential email map ─────────────────────────────────────────────
# Source: SOD Location In-charges.msg (Outlook contact group, extracted 2026-06)
# Keys = Plant Code (str), Values = Location In-charge email

LOCATION_EMAIL_MAP = {
    "1128": "mtra.tmlic@hpcl.in",
    "1146": "meer.irdic@hpcl.in",
    "1155": "roork.irdic@hpcl.in",
    "1157": "brly.irdic@hpcl.in",
    "1164": "bhatinda.irdic@hpcl.in",
    "1180": "sgrur.irdic@hpcl.in",        # Sangrur Depot
    "1183": "nlgrh.irdic@hpcl.in",        # Nalgarh Depot
    "1187": "vinodkdhamija@hpcl.in",      # HMEL White Oil Terminal
    "1216": "baha.tmlic@hpcl.in",
    "1221": "del.tmlic@hpcl.in",
    "1222": "lalk.top@hpcl.in",           # LALKUAN-IOC
    "1224": "dineshbahadur@hpcl.in",      # Barmer Terminal
    "1233": "bhtp.irdic@hpcl.in",
    "1242": "jaip.tmlic@hpcl.in",
    "1254": "jal.irdic@hpcl.in",
    "1256": "jammu.irdic@hpcl.in",
    "1259": "srin.irdic@hpcl.in",
    "1265": "leh.irdic@hpcl.in",
    "1278": "jodh.irdic@hpcl.in",
    "1281": "ajmer.tmlic@hpcl.in",
    "1292": "lko.irdic@hpcl.in",
    "1305": "kanp.tmlic@hpcl.in",
    "1308": "strgnj.irdic@hpcl.in",
    "1313": "btlpr.irdic@hpcl.in",
    "1319": "mgsr.irdic@hpcl.in",
    "1334": "rwri.irdic@hpcl.in",
    "1341": "hisr.tmlic@hpcl.in",
    "1385": "plnpr.tmlic@hpcl.in",
    "1396": "manm.irdic@hpcl.in",
    "1397": "manm.irdic@hpcl.in",
    "1410": "vado.irdic@hpcl.in",
    "1412": "hazira.irdic@hpcl.in",
    "1424": "jaba.irdic@hpcl.in",
    "1435": "gwal.irdic@hpcl.in",
    "1436": "sagar.irdic@hpcl.in",
    "1442": "serajul.haque@hpcl.in",       # BAKANIA TOP-RIL
    "1449": "vadi.top@hpcl.in",
    "1457": "indo.irdic@hpcl.in",
    "1462": "blr.irdic@hpcl.in",           # BANGROD TOP-IOC
    "1485": "akol.irdic@hpcl.in",
    "1498": "ward.irdic@hpcl.in;wardha.top@hpcl.in",  # WARDHA TOP-NEL (two IDs)
    "1504": "miraj.irdic@hpcl.in",
    "1509": "slpr.irdic@hpcl.in",
    "1527": "rpr.irdic@hpcl.in",
    "1528": "rpr.irdic@hpcl.in",
    "1541": "vadi.top@hpcl.in",            # HPCL JAMNAGAR (same as Vadinar TOP)
    "1546": "mund.tmlic@hpcl.in",
    "1551": "vasco.tmlic@hpcl.in",
    "1552": "vasco.tmlic@hpcl.in",
    "1554": "mum.tmlic@hpcl.in",
    "1583": "kandla.tmlic@hpcl.in",
    "1584": "loni.tmlic@hpcl.in",
    "1585": "mahul.tmlic@hpcl.in",
    "1588": "vashi.tmlic@hpcl.in",
    "1629": "prdp.tmlic@hpcl.in",
    "1630": "prdp.tmlic@hpcl.in",
    "1636": "blsr.irdic@hpcl.in",
    "1644": "kolk.tmlic@hpcl.in",
    "1649": "hald.tmlic@hpcl.in",
    "1650": "hald.tmlic@hpcl.in",
    "1652": "mour.top@hpcl.in",
    "1655": "drgp.irdic@hpcl.in",          # RAJBUNDH IOC TOP
    "1656": "drgp.irdic@hpcl.in",
    "1672": "bngn.top@hpcl.in",
    "1676": "digb.top@hpcl.in",            # DIGBOI TOP-IOC
    "1677": "guwa.irdic@hpcl.in",
    "1687": "mald.top@hpcl.in",            # MALDA TOP-IOC
    "1689": "numa.top@hpcl.in",
    "1691": "silg.top@hpcl.in",            # RANGPO TOP-IOC
    "1693": "silg.top@hpcl.in",
    "1698": "rahulchangmai@hpcl.in",       # Dimapur TOP-IOC
    "1700": "silg.top@hpcl.in",            # RANGAPANI TOP-NRL
    "1708": "rahulchangmai@hpcl.in",       # Dimapur Depot
    "1711": "brni.irdic@hpcl.in",
    "1712": "brni.irdic@hpcl.in",
    "1723": "patna.irdic@hpcl.in",
    "1742": "bokr.irdic@hpcl.in",
    "1743": "bokr.irdic@hpcl.in",          # JASIDIH TOP-IOC
    "1775": "blr.tmlic@hpcl.in",
    "1777": "hassn.tmlic@hpcl.in",
    "1797": "gulb.irdic@hpcl.in",
    "1800": "hubl.irdic@hpcl.in",
    "1831": "coch.tmlic@hpcl.in",
    "1898": "mang.tmlic@hpcl.in",          # KASARGOD TOP-ONGC
    "1845": "irum.tmlic@hpcl.in",
    "1856": "coim.irdic@hpcl.in",
    "1871": "madu.irdic@hpcl.in",
    "1879": "ten.irdic@hpcl.in",
    "1892": "kozh.irdic@hpcl.in",
    "1895": "mang.tmlic@hpcl.in",
    "1915": "rmgdm.irdic@hpcl.in",
    "1919": "secu.tmlic@hpcl.in",
    "1937": "sury.tmlic@hpcl.in",
    "1940": "kada.irdic@hpcl.in",
    "1953": "rjmdy.tmlic@hpcl.in",
    "1973": "cass.tmlic@hpcl.in",
    "1979": "vija.tmlic@hpcl.in",
    "1991": "chen.tmlic@hpcl.in",
    "1992": "vskp.tmlic@hpcl.in",
    "1999": "dhmpri.tmlic@hpcl.in",
    "3129": "vinodkdhamija@hpcl.in",      # HMEL Bitumen Terminal
    "3562": "vashi.bo.tmlic@hpcl.in",
    "3693": "kaknd.tmlic@hpcl.in",
    "3718": "vskp.bo.tmlic@hpcl.in",
    "3833": "coch.tmlic@hpcl.in",
    # Skip: 1899 Ferokee TOP IOC — no email; 1882 VNKOTI TOP-IOC — not functional;
    #       1445 Borkheri TOP BPC — not functional; 1885 MRPL TOP — non-HPCL
}


# ── Zone credential email map ──────────────────────────────────────────────────
# Maps zone name → dict with "to" and "cc" for sending zone account credentials.
# "to" = personal email of the Zone OD (from Zonal OD.msg + ZONE_EMAIL_MAP)
# Credential emails for zone accounts (BLRMIS, BHOMIS, etc.) go to these recipients.

ZONE_CREDENTIAL_MAP = {k: v for k, v in ZONE_EMAIL_MAP.items()}


def build_credentials_email_html(
    loc_name: str,
    loc_code: str,
    password: str,
    app_url: str = "https://hpcl-sod-mis.streamlit.app",
) -> str:
    return (
        "<!DOCTYPE html><html><head><meta charset='UTF-8'></head>"
        "<body style='font-family:Arial,sans-serif;margin:0;padding:0;background:#f4f6fb;'>"
        "<div style='max-width:600px;margin:30px auto;background:white;border-radius:14px;"
        "box-shadow:0 2px 12px rgba(0,0,0,0.10);overflow:hidden;'>"
        "<div style='background:#002B8F;padding:22px 30px;'>"
        "<div style='color:white;font-size:20px;font-weight:700;letter-spacing:0.5px;'>"
        "HPCL SOD &mdash; MIS Portal</div>"
        "<div style='color:#a8bfe8;font-size:12px;margin-top:4px;'>"
        "Supply, Operations &amp; Distribution &nbsp;&middot;&nbsp; Login Credentials</div>"
        "</div>"
        "<div style='padding:28px 30px;'>"
        f"<p style='font-size:15px;color:#333;margin-top:0;'>Dear <strong>{loc_name} Team</strong>,</p>"
        "<p style='font-size:14px;color:#444;line-height:1.8;'>Your login credentials for the "
        "<strong>HPCL SOD MIS Portal</strong> are shared below. "
        "Please log in and change your password after the first login.</p>"
        "<table style='width:100%;border-collapse:collapse;margin:18px 0;'>"
        "<tr style='background:#002B8F;'>"
        "<th style='padding:10px 16px;color:white;text-align:left;font-size:13px;width:40%;'>Detail</th>"
        "<th style='padding:10px 16px;color:white;text-align:left;font-size:13px;'>Value</th></tr>"
        "<tr style='background:#f5f7ff;'>"
        "<td style='padding:10px 16px;font-size:14px;font-weight:600;color:#002B8F;'>Portal URL</td>"
        f"<td style='padding:10px 16px;font-size:14px;'><a href='{app_url}' style='color:#0033A0;'>{app_url}</a></td></tr>"
        "<tr style='background:#ffffff;'>"
        "<td style='padding:10px 16px;font-size:14px;font-weight:600;color:#002B8F;'>User ID</td>"
        f"<td style='padding:10px 16px;font-size:14px;font-weight:700;letter-spacing:1px;'>{loc_code}</td></tr>"
        "<tr style='background:#f5f7ff;'>"
        "<td style='padding:10px 16px;font-size:14px;font-weight:600;color:#002B8F;'>Password</td>"
        f"<td style='padding:10px 16px;font-size:14px;font-weight:700;letter-spacing:1px;'>{password}</td></tr>"
        "</table>"
        "<div style='background:#fff8e1;border-left:4px solid #f59e0b;border-radius:6px;"
        "padding:12px 16px;margin:16px 0;font-size:13px;color:#7c5c00;'>"
        "<strong>&#9888; Important:</strong> Please change your password after the first login "
        "using the <em>Change Password</em> option in the sidebar. Keep your credentials confidential.</div>"
        "<p style='font-size:14px;color:#333;line-height:1.8;margin-top:16px;'>"
        "The MIS data submission for your location is due by the <strong>15th of every month</strong>"
        " (for the previous month). For any issues, raise a support ticket from within the portal"
        f" or contact <a href='mailto:{SENDER_EMAIL}' style='color:#0033A0;'>{SENDER_EMAIL}</a>.</p>"
        "<p style='font-size:12px;color:#aaa;border-top:1px solid #eee;"
        "padding-top:16px;margin-top:24px;line-height:1.7;'>"
        "This is an auto-generated message from the HPCL SOD MIS Portal.<br>"
        f"Sent by: {SENDER_EMAIL} &nbsp;&middot;&nbsp; Supply, Operations &amp; Distribution HQO<br>"
        "Please do not reply directly to this email.</p>"
        "</div></div></body></html>"
    )


def send_credential_email(
    to_email: str,
    loc_name: str,
    loc_code: str,
    password: str,
    app_url: str = "https://hpcl-sod-mis.streamlit.app",
    cc_email: str = "",
    test_mode: bool = False,
    test_email: str = "",
) -> dict:
    """Send login credentials to a location via Outlook COM."""
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError:
        return {"ok": False, "msg": "pywin32 not installed. Run: pip install pywin32"}

    try:
        html_body  = build_credentials_email_html(loc_name, loc_code, password, app_url)
        pfx        = "[TEST] " if test_mode else ""
        subject    = f"{pfx}HPCL SOD MIS Portal — Login Credentials | {loc_name}"
        actual_to  = (test_email or SENDER_EMAIL) if test_mode else to_email
        actual_cc  = "" if test_mode else cc_email
        actual_bcc = "" if test_mode else SENDER_EMAIL  # BCC sender for audit trail

        pythoncom.CoInitialize()
        try:
            outlook   = win32.Dispatch("Outlook.Application")
            mail_item = outlook.CreateItem(0)
            mail_item.To       = actual_to
            if actual_cc:
                mail_item.CC   = actual_cc
            if actual_bcc:
                mail_item.BCC  = actual_bcc
            mail_item.Subject  = subject
            mail_item.HTMLBody = html_body
            try:
                mail_item.Send()
                return {"ok": True, "mode": "sent", "to": actual_to}
            except Exception as send_exc:
                try:
                    mail_item.Save()
                    return {"ok": True, "mode": "draft",
                            "msg": f"Saved as Outlook draft (send error: {send_exc})",
                            "to": actual_to}
                except Exception as draft_exc:
                    return {"ok": False, "msg": f"Send and draft both failed: {draft_exc}"}
        finally:
            pythoncom.CoUninitialize()
    except Exception as exc:
        return {"ok": False, "msg": f"Outlook error: {exc}"}
