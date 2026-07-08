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

_MONTH_ORDER = {m: i for i, m in enumerate(
    ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
)}


def month_sort_key(my: str):
    """Chronological key for 'Apr-2026' style strings -- plain sorted() on these
    is alphabetical (Apr, Jun, May), not chronological (Apr, May, Jun)."""
    parts = my.split("-")
    mon = _MONTH_ORDER.get(parts[0], 0)
    yr  = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 0
    return (yr, mon)


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

def _send_outlook(to: str, subject: str, html_body: str,
                  cc: str = "", bcc: str = "",
                  inline_images: list | None = None) -> dict:
    """Send email via Outlook COM — same pattern as Auto Reco 2.0 mailer.py.

    Outlook must be installed and signed in. COM objects are explicitly deleted
    and gc.collect() called before CoUninitialize to avoid Sent-Items race.

    inline_images: optional [(cid, png_bytes), ...] -- embedded as attachments
    with PR_ATTACH_CONTENT_ID set, so html_body can reference them via
    <img src="cid:...">. Outlook (Word rendering engine) doesn't reliably
    display base64 data-URI images, so this attach-by-CID approach is used
    instead. Temp files are written and cleaned up around the send.
    """
    import gc
    try:
        import pythoncom
        import win32com.client as win32
    except Exception as exc:
        return {"ok": False, "msg": f"pywin32 not installed: {exc}"}

    com_initialized = False
    outlook     = None
    mail_item   = None
    temp_paths  = []
    PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001E"
    try:
        pythoncom.CoInitialize()
        com_initialized = True

        outlook   = win32.Dispatch("Outlook.Application")
        mail_item = outlook.CreateItem(0)
        mail_item.To      = to
        mail_item.Subject = subject
        if cc.strip():
            mail_item.CC  = cc
        if bcc.strip():
            mail_item.BCC = bcc

        for cid, png_bytes in (inline_images or []):
            import tempfile, os
            fd, path = tempfile.mkstemp(suffix=".png", prefix=f"emb_{cid}_")
            with os.fdopen(fd, "wb") as f:
                f.write(png_bytes)
            temp_paths.append(path)
            attachment = mail_item.Attachments.Add(path)
            attachment.PropertyAccessor.SetProperty(PR_ATTACH_CONTENT_ID, cid)

        # HTMLBody must be set AFTER attachments are added, or Outlook can
        # drop the cid: reference for inline images added afterward.
        mail_item.HTMLBody = html_body

        try:
            mail_item.Send()
            return {"ok": True, "mode": "sent"}
        except Exception as send_exc:
            try:
                mail_item.Save()
                return {"ok": True, "mode": "draft",
                        "msg": f"Saved as Outlook draft: {send_exc}"}
            except Exception as draft_exc:
                return {"ok": False, "msg": f"Send and draft both failed: {draft_exc}"}

    except Exception as exc:
        return {"ok": False, "msg": f"Outlook error: {exc}"}

    finally:
        try:
            del mail_item
        except Exception:
            pass
        try:
            del outlook
        except Exception:
            pass
        gc.collect()
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
        import os as _os
        for p in temp_paths:
            try:
                _os.remove(p)
            except Exception:
                pass


def email_configured() -> tuple:
    """Return (True, '') if Outlook COM is reachable, else (False, reason)."""
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError:
        return False, "pywin32 not installed — run: pip install pywin32"
    try:
        import gc
        pythoncom.CoInitialize()
        ol = None
        try:
            ol = win32.Dispatch("Outlook.Application")
            return True, ""
        finally:
            try:
                del ol
            except Exception:
                pass
            gc.collect()
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
    except Exception as exc:
        return False, str(exc)


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
    c = _get_zone_map().get(zone_name, {})
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
    """Send one reminder email for a zone.

    TO  = Zone OD + all pending location in-charges (that have email configured).
    CC  = Zone IC/OND.
    BCC = none.
    """
    contacts = _get_zone_map().get(zone_name)
    if not contacts:
        return {"ok": False, "msg": f"No email contacts configured for: {zone_name}"}

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
        to_str, cc_str = test_email or SENDER_EMAIL, ""
    else:
        # Build TO: Zone OD + pending location emails
        loc_email_map = _get_loc_map()
        loc_emails = [
            loc_email_map[str(loc.get("userId", ""))]
            for loc in pending_locs
            if loc_email_map.get(str(loc.get("userId", "")))
        ]
        zone_od = contacts.get("to", "")
        to_str  = "; ".join(filter(None, [zone_od] + loc_emails))
        cc_str  = contacts.get("cc", "")

    return _send_outlook(to_str, subject, html_body, cc_str, "")


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
                f"[TEST] Zone email {mode} to {test_email or SENDER_EMAIL} "
                f"(sample: {first_zone}, {len(first_locs)} location(s)). "
                f"In production would send to {len(zones_to_send)} zone(s) "
                f"with location emails in TO."
            )
        else:
            failed_count = 1
            errors = [res.get("msg", "")]
            msg = f"[TEST] Send failed: {errors[0]}"
        return {"ok": res["ok"], "sent": sent_count, "failed": failed_count,
                "errors": errors, "msg": msg}

    for zone_name, locs in zones_to_send:
        if zone_name not in _get_zone_map():
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
        f"Zone reminder emails sent to {sent_count} zone(s) "
        f"(Zone OD + location in-charges in TO, Zone IC in CC)."
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


# ── Consolidated HQO summary email ────────────────────────────────────────────

def build_consolidated_html(month_year: str,
                            all_location_rows: list,
                            due_date: date) -> str:
    """Build one HTML email listing all pending locations across all zones."""
    today   = date.today()
    overdue = today > due_date
    due_str = due_date.strftime("%d %b %Y")

    pending_by_zone: dict = {}
    for loc in all_location_rows:
        if loc.get("status") != "SUBMITTED":
            zone = loc.get("zone", "Unknown")
            pending_by_zone.setdefault(zone, []).append(loc)

    total_pending = sum(len(v) for v in pending_by_zone.values())

    zones_html = ""
    for zone_name in sorted(pending_by_zone.keys()):
        locs = pending_by_zone[zone_name]
        rows_html = ""
        for loc in locs:
            uid    = loc.get("userId", "")
            name   = loc.get("locName", "")
            status = loc.get("status", "NOT_STARTED").replace("_", " ").title()
            pct    = int(float(loc.get("completion_pct", 0)))
            bg     = "#fff8f8" if overdue else "#ffffff"
            rows_html += (
                f"<tr style='background:{bg};'>"
                f"<td style='padding:7px 11px;border-bottom:1px solid #e8ecf4;font-size:12px;'>{uid}</td>"
                f"<td style='padding:7px 11px;border-bottom:1px solid #e8ecf4;font-size:12px;'>{name}</td>"
                f"<td style='padding:7px 11px;border-bottom:1px solid #e8ecf4;font-size:12px;'>{status}</td>"
                f"<td style='padding:7px 11px;border-bottom:1px solid #e8ecf4;font-size:12px;"
                f"text-align:center;'>{pct}%</td>"
                f"</tr>"
            )
        zones_html += (
            f"<div style='margin-bottom:22px;'>"
            f"<div style='background:#003580;color:white;padding:8px 14px;"
            f"font-size:13px;font-weight:700;border-radius:6px 6px 0 0;'>"
            f"{zone_name} — {len(locs)} location(s) pending</div>"
            f"<table style='width:100%;border-collapse:collapse;'>"
            f"<thead><tr style='background:#dde4f0;'>"
            f"<th style='padding:7px 11px;text-align:left;font-size:12px;color:#333;'>Code</th>"
            f"<th style='padding:7px 11px;text-align:left;font-size:12px;color:#333;'>Location</th>"
            f"<th style='padding:7px 11px;text-align:left;font-size:12px;color:#333;'>Status</th>"
            f"<th style='padding:7px 11px;text-align:center;font-size:12px;color:#333;'>Done%</th>"
            f"</tr></thead>"
            f"<tbody>{rows_html}</tbody>"
            f"</table></div>"
        )

    overdue_banner = ""
    if overdue:
        overdue_banner = (
            f"<div style='background:#fdecea;border:2px solid #e53935;border-radius:8px;"
            f"padding:12px 18px;margin:16px 0;color:#b71c1c;font-weight:700;font-size:14px;'>"
            f"&#9888;&nbsp; OVERDUE — Deadline was {due_str}. "
            f"Pending across {len(pending_by_zone)} zone(s).</div>"
        )

    return f"""<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;margin:0;padding:0;background:#f4f6fb;">
  <div style="max-width:760px;margin:30px auto;background:white;border-radius:14px;
              box-shadow:0 2px 12px rgba(0,0,0,0.10);overflow:hidden;">
    <div style="background:#002B8F;padding:22px 30px;">
      <div style="color:white;font-size:20px;font-weight:700;letter-spacing:0.5px;">
        HPCL SOD &mdash; MIS Portal</div>
      <div style="color:#a8bfe8;font-size:12px;margin-top:4px;">
        Supply, Operations &amp; Distribution &nbsp;&middot;&nbsp; Consolidated Pending Report
      </div>
    </div>
    <div style="padding:28px 30px;">
      <p style="font-size:15px;color:#333;margin-top:0;">Dear Team,</p>
      <p style="font-size:14px;color:#444;line-height:1.8;">
        Consolidated MIS submission status for <strong>{month_year}</strong>:
        <strong>{total_pending} location(s)</strong> across
        <strong>{len(pending_by_zone)} zone(s)</strong> have not yet submitted.
      </p>
      {overdue_banner}
      {zones_html}
      <p style="font-size:14px;color:#333;margin-top:16px;">
        Submission deadline: <strong>{due_str}</strong>
      </p>
    </div>
  </div>
</body>
</html>"""


def send_consolidated_reminder(month_year: str,
                               all_location_rows: list,
                               due_date: date,
                               test_mode: bool = False,
                               test_email: str = "") -> dict:
    """Send one consolidated pending-report email.

    TO  = shoaibrehman@hpcl.in
    CC  = SOD.OPNS.HQO, bhsgk, shubham.tayal  (former BCC recipients)
    BCC = none
    """
    today   = date.today()
    overdue = today > due_date
    pfx     = "[TEST] " if test_mode else ""

    pending = [r for r in all_location_rows if r.get("status") != "SUBMITTED"]
    zones   = sorted({r.get("zone", "") for r in pending})
    subject = (
        f"{pfx}OVERDUE — Consolidated MIS Pending Report | {month_year} | "
        f"{len(pending)} location(s) across {len(zones)} zone(s)"
        if overdue else
        f"{pfx}Consolidated MIS Pending Report | {month_year} | "
        f"{len(pending)} location(s) across {len(zones)} zone(s)"
    )
    html_body = build_consolidated_html(month_year, all_location_rows, due_date)

    if test_mode:
        to_str, cc_str = test_email or SENDER_EMAIL, ""
    else:
        to_str = SENDER_EMAIL
        cc_str = "; ".join(BCC_EMAILS)

    res = _send_outlook(to_str, subject, html_body, cc_str, "")
    if "msg" not in res:
        res["msg"] = (
            f"Consolidated report {'sent' if res.get('mode') == 'sent' else res.get('mode','sent')} "
            f"to {to_str}" + (f" (CC: {cc_str})" if cc_str else "") + "."
        )
    return res


# ── Multi-month email builders & senders ──────────────────────────────────────

_STATUS_COLORS = {
    # Semantic 4-color scheme (confirmed): Green = no action required,
    # Orange = with checker for review, Red = needs maker action (in
    # progress / rejected), Grey = not started.
    #
    # Rendered as plain colored TEXT (no fill), so these are deliberately
    # darker/richer than typical "fill" shades -- a color that reads fine
    # as white-on-solid-background often fails contrast as plain text on
    # white (e.g. light grey #9E9E9E is ~2.8:1 on white, well under the
    # ~4.5:1 needed for legible text) -- each shade below targets >=4.5:1.
    "SUBMITTED":       "#2E7D32",   # dark green  -- no action required
    "LOCKED":          "#2E7D32",   # dark green  -- no action required
    "PENDING_REVIEW":  "#B75900",   # dark amber/orange -- awaiting checker review
    "IN_PROGRESS":     "#C62828",   # dark red    -- needs maker action to submit
    "REJECTED":        "#C62828",   # dark red    -- needs maker action to resubmit
    "NOT_STARTED":     "#5F6368",   # slate grey  -- not even started (darker than
                                    # a typical light grey so it stays legible on white)
}


def _status_badge(status: str) -> str:
    color = _STATUS_COLORS.get(status, "#5F6368")
    label = status.replace("_", " ").title()
    return (f"<span style='color:{color};font-size:13px;font-weight:700;"
            f"white-space:nowrap;'>{label}</span>")


def _pivot_table_html(months_data: dict, month_list: list, include_zone_col: bool) -> str:
    """One row per location, one column per month -- status pivoted across months.

    months_data : {month_year: [ALL location rows for that month]} -- UNFILTERED,
                  so a location's status for months it already submitted is still
                  visible next to months it's still pending.
    month_list  : ordered list of month_year strings to show as columns.
    Only includes locations pending in at least one month in month_list --
    this is a pending-reminder email, not a full compliance matrix.
    """
    by_loc: dict   = {}   # userId -> {month_year: row}
    loc_meta: dict = {}   # userId -> (zone, locName)
    for my in month_list:
        for row in months_data.get(my, []):
            uid = row.get("userId", "")
            by_loc.setdefault(uid, {})[my] = row
            loc_meta[uid] = (row.get("zone", ""), row.get("locName", uid))

    pending_uids = [
        uid for uid, months in by_loc.items()
        if any(months.get(my, {}).get("status") != "SUBMITTED" for my in month_list)
    ]
    if not pending_uids:
        return "<p style='font-size:13px;color:#666;'>No pending locations.</p>"

    pending_uids.sort(key=lambda u: (loc_meta[u][0], loc_meta[u][1]))

    month_hdrs = "".join(
        f"<th style='padding:8px 10px;text-align:center;font-size:11px;color:#333;'>{my}</th>"
        for my in month_list
    )
    zone_hdr = (
        "<th style='padding:8px 12px;text-align:left;font-size:11px;color:#333;'>Zone</th>"
        if include_zone_col else ""
    )
    rows_html = ""
    for i, uid in enumerate(pending_uids):
        zone, name = loc_meta[uid]
        bg = "#fafbfe" if i % 2 else "#ffffff"
        zone_cell = (
            f"<td style='padding:8px 12px;border-bottom:1px solid #e8ecf4;font-size:12px;'>{zone}</td>"
            if include_zone_col else ""
        )
        cells = "".join(
            f"<td style='padding:8px 10px;border-bottom:1px solid #e8ecf4;text-align:center;'>"
            + (_status_badge(by_loc[uid][my]["status"]) if my in by_loc[uid] else "&mdash;")
            + "</td>"
            for my in month_list
        )
        rows_html += (
            f"<tr style='background:{bg};'>"
            f"{zone_cell}"
            f"<td style='padding:8px 12px;border-bottom:1px solid #e8ecf4;font-size:12px;font-weight:600;'>{name}</td>"
            f"{cells}"
            f"</tr>"
        )

    return (
        f"<table style='width:100%;border-collapse:collapse;'>"
        f"<thead><tr style='background:#e8ecf4;'>"
        f"{zone_hdr}"
        f"<th style='padding:8px 12px;text-align:left;font-size:11px;color:#333;'>Location</th>"
        f"{month_hdrs}"
        f"</tr></thead>"
        f"<tbody>{rows_html}</tbody>"
        f"</table>"
    )


# ── Zone performance chart (embedded in consolidated email) ───────────────────

# Vivid Office/PowerPoint theme accent colors (not pastel/washed out).
_CHART_COLORS  = ["#4472C4", "#ED7D31", "#70AD47", "#FFC000", "#5B9BD5", "#A5A5A5"]
_CHART_BORDER  = "#33475B"   # one consistent dark border for all bars/swatches


def build_zone_performance_chart(months_rows: dict, month_list: list):
    """Render a grouped bar chart (zone x month submission compliance %) as PNG
    bytes, using Pillow -- no matplotlib/kaleido dependency (neither is
    installed on this machine, and pip install is blocked by network SSL
    restrictions here). Hand-drawn bars/gridlines/legend/labels.

    Rendered at 2x pixel density (SCALE) and displayed at 1x via explicit
    width/height on the <img> tag -- the standard "retina image" technique.
    Without this, Outlook/Windows display scaling can make small chart text
    look visibly blurry.

    months_rows : {month_year: [ALL loc rows for that month]}
    month_list  : ordered month_year strings to chart as grouped bars.
    Returns (png_bytes, display_width, display_height); png_bytes is b"" if
    there is nothing to chart.
    """
    from PIL import Image, ImageDraw, ImageFont
    import io as _io

    zone_month_pct: dict = {}
    for my in month_list:
        by_zone: dict = {}
        for r in months_rows.get(my, []):
            by_zone.setdefault(r.get("zone", "Unknown"), []).append(r)
        for z, locs in by_zone.items():
            total = len(locs)
            done  = sum(1 for l in locs if l.get("status") == "SUBMITTED")
            zone_month_pct.setdefault(z, {})[my] = (done / total * 100) if total else 0.0

    zones = sorted(zone_month_pct.keys())
    if not zones:
        return b"", 0, 0

    SCALE = 2   # render at 2x, display at 1x -- crisp on any screen/DPI

    n_months  = len(month_list)
    W         = 1000 * SCALE
    margin_l  = 190 * SCALE
    margin_r  = 50 * SCALE
    margin_t  = 60 * SCALE
    bar_h     = 7 * SCALE     # thickness of one bar
    intra_gap = 3 * SCALE     # gap between bars WITHIN the same zone's cluster
    zone_gap  = 14 * SCALE    # extra whitespace BETWEEN different zones' clusters
    group_h   = n_months * bar_h + (n_months - 1) * intra_gap
    row_h     = group_h + zone_gap
    H         = margin_t + row_h * len(zones) + 50 * SCALE
    plot_w    = W - margin_l - margin_r

    try:
        font       = ImageFont.truetype(r"C:\Windows\Fonts\segoeui.ttf", 13 * SCALE)
        font_small = ImageFont.truetype(r"C:\Windows\Fonts\segoeui.ttf", 11 * SCALE)
        font_title = ImageFont.truetype(r"C:\Windows\Fonts\segoeuib.ttf", 16 * SCALE)
    except Exception:
        font = font_small = font_title = ImageFont.load_default()

    img  = Image.new("RGB", (W, H), "#fcfcfb")
    draw = ImageDraw.Draw(img)

    draw.text((margin_l, 16 * SCALE), "Zone Submission Compliance % — by Month",
              font=font_title, fill="#0b0b0b")

    # Legend (one swatch + label per month, fixed color order)
    lx, ly = margin_l, 42 * SCALE
    sw = 12 * SCALE
    for i, my in enumerate(month_list):
        color = _CHART_COLORS[i % len(_CHART_COLORS)]
        draw.rounded_rectangle([lx, ly, lx + sw, ly + sw], radius=2 * SCALE,
                               fill=color, outline=_CHART_BORDER, width=SCALE)
        draw.text((lx + sw + 6 * SCALE, ly - 1 * SCALE), my, font=font_small, fill="#52514e")
        lx += sw + 6 * SCALE + int(draw.textlength(my, font=font_small)) + 22 * SCALE

    plot_top = margin_t
    plot_bot = margin_t + row_h * (len(zones) - 1) + group_h

    # Gridlines + x-axis %
    for pct in (0, 25, 50, 75, 100):
        x = margin_l + int(plot_w * pct / 100)
        draw.line([(x, plot_top - 4 * SCALE), (x, plot_bot)], fill="#e1e0d9", width=SCALE)
        draw.text((x - 10 * SCALE, plot_bot + 6 * SCALE), f"{pct}%", font=font_small, fill="#898781")

    # Grouped bars, one row per zone
    for zi, z in enumerate(zones):
        row_top = plot_top + zi * row_h
        label = z if len(z) <= 26 else z[:24] + "…"
        draw.text((10 * SCALE, row_top + group_h / 2 - 7 * SCALE), label, font=font, fill="#0b0b0b")
        for mi, my in enumerate(month_list):
            pct   = zone_month_pct[z].get(my, 0.0)
            color = _CHART_COLORS[mi % len(_CHART_COLORS)]
            y0 = row_top + mi * (bar_h + intra_gap)
            y1 = y0 + bar_h
            x0 = margin_l
            x1 = margin_l + max(int(plot_w * pct / 100), 2 * SCALE)
            draw.rounded_rectangle([x0, y0, x1, y1], radius=2 * SCALE,
                                   fill=color, outline=_CHART_BORDER, width=SCALE)

    buf = _io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue(), W // SCALE, H // SCALE


def build_multimonth_zone_html(zone_name: str, months_rows: dict,
                               due_dates: dict, custom_intro: str = "") -> str:
    """HTML for one zone covering multiple months (TO = Zone OD + location in-charges).

    months_rows must be UNFILTERED (all locations, all statuses) per month, so
    the pivot table can show a location's Submitted status for months it's
    done alongside Pending for months it isn't -- see _pivot_table_html.
    """
    month_list  = sorted(months_rows.keys(), key=month_sort_key)
    total_pend  = sum(1 for rows in months_rows.values()
                      for r in rows if r.get("status") != "SUBMITTED")
    my_list     = ", ".join(month_list)
    intro   = (
        custom_intro.strip()
        or (
            f"This is a combined reminder for MIS data submission for "
            f"<strong>{my_list}</strong>. The table below shows each location's "
            f"status across these months. Please ensure all pending submissions "
            f"are completed at the earliest."
        )
    )
    return f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;margin:0;padding:0;background:#f4f6fb;">
<div style="max-width:760px;margin:30px auto;background:white;border-radius:14px;
            box-shadow:0 2px 12px rgba(0,0,0,0.10);overflow:hidden;">
  <div style="background:#002B8F;padding:22px 30px;">
    <div style="color:white;font-size:20px;font-weight:700;">HPCL SOD &mdash; MIS Portal</div>
    <div style="color:#a8bfe8;font-size:12px;margin-top:4px;">
      Supply, Operations &amp; Distribution &middot; MIS Submission Reminder</div>
  </div>
  <div style="padding:28px 30px;">
    <p style="font-size:15px;color:#333;margin-top:0;">
      Dear <strong>{zone_name} Team</strong>,</p>
    <p style="font-size:14px;color:#444;line-height:1.8;">{intro}</p>
    <p style="font-size:13px;color:#666;margin-bottom:18px;">
      Total: <strong>{total_pend} location-month(s)</strong> pending across
      <strong>{len(month_list)} month(s)</strong>
    </p>
    {_pivot_table_html(months_rows, month_list, include_zone_col=False)}
  </div>
</div>
</body></html>"""


def build_multimonth_consolidated_html(months_rows: dict, due_dates: dict,
                                       chart_cid: str = "",
                                       chart_w: int = 0, chart_h: int = 0) -> str:
    """HTML consolidated email — all zones, all selected months.

    months_rows must be UNFILTERED (all locations, all statuses) per month.
    chart_cid: if set, embeds <img src="cid:{chart_cid}"> for the attached
    zone-performance bar chart (see build_zone_performance_chart).
    chart_w/chart_h: explicit DISPLAY size (not the underlying pixel size,
    which is 2x for crispness) -- Outlook needs a width/height attribute,
    not just CSS, to size a cid: image correctly.
    """
    month_list = sorted(months_rows.keys(), key=month_sort_key)
    total_pend = sum(1 for rows in months_rows.values()
                     for r in rows if r.get("status") != "SUBMITTED")
    my_list    = ", ".join(month_list)
    chart_html = (
        f'<div style="margin:20px 0;text-align:center;">'
        f'<img src="cid:{chart_cid}" alt="Zone performance chart" '
        f'width="{chart_w}" height="{chart_h}" '
        f'style="max-width:100%;height:auto;border-radius:8px;border:1px solid #e8ecf4;"></div>'
        if chart_cid else ""
    )
    return f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;margin:0;padding:0;background:#f4f6fb;">
<div style="max-width:800px;margin:30px auto;background:white;border-radius:14px;
            box-shadow:0 2px 12px rgba(0,0,0,0.10);overflow:hidden;">
  <div style="background:#002B8F;padding:22px 30px;">
    <div style="color:white;font-size:20px;font-weight:700;">HPCL SOD &mdash; MIS Portal</div>
    <div style="color:#a8bfe8;font-size:12px;margin-top:4px;">
      Supply, Operations &amp; Distribution &middot; Consolidated Pending Report</div>
  </div>
  <div style="padding:28px 30px;">
    <p style="font-size:15px;color:#333;margin-top:0;">Dear Team,</p>
    <p style="font-size:14px;color:#444;line-height:1.8;">
      Consolidated MIS pending status for <strong>{my_list}</strong>:
      <strong>{total_pend} location-month(s)</strong> across
      <strong>{len(month_list)} month(s)</strong> not yet submitted.
    </p>
    {chart_html}
    {_pivot_table_html(months_rows, month_list, include_zone_col=True)}
  </div>
</div>
</body></html>"""


def send_all_multimonth_reminders(months_data: dict, due_dates: dict,
                                  custom_intro: str = "",
                                  test_mode: bool = False,
                                  test_email: str = "") -> dict:
    """Send one zone email per zone covering all selected months.

    months_data : {month_year: [all_loc_rows]}
    due_dates   : {month_year: date}
    Zone email  : TO = Zone OD + unique location emails across months; CC = Zone IC
    """
    ok, err = email_configured()
    if not ok:
        return {"ok": False, "sent": 0, "failed": 0, "errors": [], "msg": err}

    # Build {zone: {month: [ALL_locs]}} -- unfiltered, so the pivot table can
    # show a location's Submitted status for months it's done alongside
    # Pending for months it isn't. zones_with_pending decides who gets an
    # email at all; it's tracked separately since zone_months itself now
    # includes locations that have already submitted everything.
    zone_months: dict = {}
    zones_with_pending: set = set()
    for my, rows in months_data.items():
        for loc in rows:
            z = loc.get("zone", "Unknown")
            zone_months.setdefault(z, {}).setdefault(my, []).append(loc)
            if loc.get("status") != "SUBMITTED":
                zones_with_pending.add(z)

    zones_with_email = sorted(z for z in zones_with_pending if z in _get_zone_map())
    if not zones_with_email:
        return {
            "ok": True, "sent": 0, "failed": 0, "errors": [],
            "msg": "All locations submitted across all selected months.",
        }

    today      = date.today()
    my_str     = ", ".join(sorted(months_data.keys(), key=month_sort_key))
    any_overdue = any(today > d for d in due_dates.values())

    if test_mode:
        first_zone  = zones_with_email[0]
        m_rows_zone = zone_months[first_zone]
        contacts    = _get_zone_map().get(first_zone, {})
        pfx         = "[TEST] "
        subject     = (
            f"{pfx}OVERDUE — MIS Submission Pending | {first_zone} | {my_str}"
            if any_overdue else
            f"{pfx}Reminder — MIS Submission Pending | {first_zone} | {my_str}"
        )
        html_body = build_multimonth_zone_html(
            first_zone, m_rows_zone, due_dates, custom_intro
        )
        res = _send_outlook(test_email or SENDER_EMAIL, subject, html_body, "", "")
        mode = res.get("mode", "sent")
        if res["ok"]:
            return {
                "ok": True, "sent": 1, "failed": 0, "errors": [],
                "msg": (
                    f"[TEST] Zone email {mode} to {test_email or SENDER_EMAIL} "
                    f"(sample: {first_zone}, months: {my_str}). "
                    f"In production would send to {len(zones_with_email)} zone(s)."
                ),
            }
        return {"ok": False, "sent": 0, "failed": 1,
                "errors": [res.get("msg", "")], "msg": f"[TEST] {res.get('msg', '')}"}

    loc_email_map = _get_loc_map()
    sent, failed, errors = 0, 0, []
    for zone_name in zones_with_email:
        m_rows_zone = zone_months[zone_name]
        contacts    = _get_zone_map().get(zone_name, {})
        any_ov_z    = any(today > due_dates.get(my, date(9999,1,1)) for my in m_rows_zone)
        subject     = (
            f"OVERDUE — MIS Submission Pending | {zone_name} | {my_str}"
            if any_ov_z else
            f"Reminder — MIS Submission Pending | {zone_name} | {my_str}"
        )
        html_body   = build_multimonth_zone_html(
            zone_name, m_rows_zone, due_dates, custom_intro
        )
        # TO = Zone OD + unique location emails across all months, but only
        # for locations that are actually pending in >=1 month -- m_rows_zone
        # is unfiltered (for the pivot table's display), so this must filter.
        loc_emails = list({
            loc_email_map[str(loc.get("userId", ""))]
            for locs in m_rows_zone.values()
            for loc in locs
            if loc.get("status") != "SUBMITTED"
            and loc_email_map.get(str(loc.get("userId", "")))
        })
        zone_od = contacts.get("to", "")
        to_str  = "; ".join(filter(None, [zone_od] + loc_emails))
        cc_str  = contacts.get("cc", "")
        res = _send_outlook(to_str, subject, html_body, cc_str, "")
        if res["ok"]:
            sent += 1
        else:
            failed += 1
            errors.append(f"{zone_name}: {res.get('msg', 'unknown')}")

    all_ok = (failed == 0)
    msg = (
        f"Zone emails sent to {sent} zone(s) covering months: {my_str}."
        if all_ok else
        f"Sent: {sent}. Failed: {failed}. " + "; ".join(errors)
    )
    return {"ok": all_ok, "sent": sent, "failed": failed, "errors": errors, "msg": msg}


def send_multimonth_consolidated_reminder(months_data: dict, due_dates: dict,
                                          test_mode: bool = False,
                                          test_email: str = "") -> dict:
    """Send one consolidated report covering all selected months.

    TO  = shoaibrehman@hpcl.in
    CC  = SOD.OPNS.HQO, bhsgk, shubham.tayal
    BCC = none
    """
    today       = date.today()
    my_str      = ", ".join(sorted(months_data.keys(), key=month_sort_key))
    total       = sum(
        1 for rows in months_data.values()
        for r in rows if r.get("status") != "SUBMITTED"
    )
    any_overdue = any(today > d for d in due_dates.values())
    pfx         = "[TEST] " if test_mode else ""
    subject     = (
        f"{pfx}OVERDUE — Consolidated MIS Pending | {my_str} | {total} pending"
        if any_overdue else
        f"{pfx}Consolidated MIS Pending Report | {my_str} | {total} pending"
    )

    chart_png, chart_w, chart_h = build_zone_performance_chart(
        months_data, sorted(months_data.keys(), key=month_sort_key)
    )
    chart_cid = "zone_perf_chart" if chart_png else ""
    html_body = build_multimonth_consolidated_html(
        months_data, due_dates, chart_cid=chart_cid, chart_w=chart_w, chart_h=chart_h
    )
    inline_images = [(chart_cid, chart_png)] if chart_png else []

    if test_mode:
        to_str, cc_str = test_email or SENDER_EMAIL, ""
    else:
        to_str = SENDER_EMAIL
        cc_str = "; ".join(BCC_EMAILS)

    res = _send_outlook(to_str, subject, html_body, cc_str, "", inline_images=inline_images)
    if "msg" not in res:
        res["msg"] = (
            f"Consolidated report sent covering {len(months_data)} month(s): {my_str}."
        )
    return res


# ── Location pending reminder ──────────────────────────────────────────────────

def build_location_pending_html(
    loc_name: str,
    loc_code: str,
    month_year: str,
    status: str,
    completion_pct: float,
    due_date: date,
) -> str:
    """Build HTML for an individual location's pending submission reminder email."""
    today   = date.today()
    overdue = today > due_date
    due_str = due_date.strftime("%d %b %Y")
    status_label = status.replace("_", " ").title()
    pct = int(float(completion_pct))

    overdue_banner = ""
    if overdue:
        overdue_banner = (
            f"<div style='background:#fdecea;border:2px solid #e53935;border-radius:8px;"
            f"padding:12px 18px;margin:16px 0;color:#b71c1c;font-weight:700;font-size:14px;'>"
            f"&#9888;&nbsp; OVERDUE — Submission deadline was {due_str}. "
            f"Immediate action is required.</div>"
        )

    return (
        "<!DOCTYPE html><html><head><meta charset='UTF-8'></head>"
        "<body style='font-family:Arial,sans-serif;margin:0;padding:0;background:#f4f6fb;'>"
        "<div style='max-width:620px;margin:30px auto;background:white;border-radius:14px;"
        "box-shadow:0 2px 12px rgba(0,0,0,0.10);overflow:hidden;'>"
        "<div style='background:#002B8F;padding:22px 30px;'>"
        "<div style='color:white;font-size:20px;font-weight:700;letter-spacing:0.5px;'>"
        "HPCL SOD &mdash; MIS Portal</div>"
        "<div style='color:#a8bfe8;font-size:12px;margin-top:4px;'>"
        "Supply, Operations &amp; Distribution &nbsp;&middot;&nbsp; MIS Submission Reminder</div>"
        "</div>"
        "<div style='padding:28px 30px;'>"
        f"<p style='font-size:15px;color:#333;margin-top:0;'>Dear <strong>{loc_name} Team</strong>,</p>"
        f"<p style='font-size:14px;color:#444;line-height:1.8;'>"
        f"This is a reminder that your MIS data submission for <strong>{month_year}</strong> "
        f"is <strong>pending</strong>. Please log in and complete your submission at the earliest.</p>"
        f"{overdue_banner}"
        "<table style='width:100%;border-collapse:collapse;margin:18px 0;'>"
        "<tr style='background:#002B8F;'>"
        "<th style='padding:10px 16px;color:white;text-align:left;font-size:13px;width:45%;'>Detail</th>"
        "<th style='padding:10px 16px;color:white;text-align:left;font-size:13px;'>Value</th></tr>"
        "<tr style='background:#f5f7ff;'>"
        f"<td style='padding:10px 16px;font-size:14px;font-weight:600;color:#002B8F;'>Location</td>"
        f"<td style='padding:10px 16px;font-size:14px;'>{loc_name} ({loc_code})</td></tr>"
        "<tr style='background:#ffffff;'>"
        f"<td style='padding:10px 16px;font-size:14px;font-weight:600;color:#002B8F;'>Month</td>"
        f"<td style='padding:10px 16px;font-size:14px;'>{month_year}</td></tr>"
        "<tr style='background:#f5f7ff;'>"
        f"<td style='padding:10px 16px;font-size:14px;font-weight:600;color:#002B8F;'>Status</td>"
        f"<td style='padding:10px 16px;font-size:14px;font-weight:700;color:#b71c1c;'>{status_label}</td></tr>"
        "<tr style='background:#ffffff;'>"
        f"<td style='padding:10px 16px;font-size:14px;font-weight:600;color:#002B8F;'>Completion</td>"
        f"<td style='padding:10px 16px;font-size:14px;'>{pct}%</td></tr>"
        "<tr style='background:#fdecea;'>"
        f"<td style='padding:10px 16px;font-size:14px;font-weight:600;color:#002B8F;'>Submission Deadline</td>"
        f"<td style='padding:10px 16px;font-size:14px;font-weight:700;color:#b71c1c;'>{due_str}</td></tr>"
        "</table>"
        f"<p style='font-size:14px;color:#333;margin-top:16px;line-height:1.8;'>"
        f"For assistance contact "
        f"<a href='mailto:{SENDER_EMAIL}' style='color:#0033A0;'>{SENDER_EMAIL}</a>.</p>"
        "</div></div></body></html>"
    )


def send_location_pending_reminder(
    to_email: str,
    loc_name: str,
    loc_code: str,
    month_year: str,
    status: str,
    completion_pct: float,
    due_date: date,
    test_mode: bool = False,
    test_email: str = "",
) -> dict:
    """Send a pending MIS submission reminder email to one location in-charge."""
    today   = date.today()
    overdue = today > due_date
    pfx     = "[TEST] " if test_mode else ""
    subject = (
        f"{pfx}OVERDUE — MIS Submission Pending | {loc_name} | {month_year}"
        if overdue else
        f"{pfx}Reminder — MIS Submission Pending | {loc_name} | {month_year}"
    )
    html_body  = build_location_pending_html(
        loc_name, loc_code, month_year, status, completion_pct, due_date
    )
    actual_to  = (test_email or SENDER_EMAIL) if test_mode else to_email
    actual_bcc = "" if test_mode else SENDER_EMAIL
    return _send_outlook(actual_to, subject, html_body, "", actual_bcc)


def send_all_location_reminders(
    month_year: str,
    all_location_rows: list,
    due_date: date,
    test_mode: bool = False,
    test_email: str = "",
) -> dict:
    """Send individual pending-submission reminders to all non-submitted locations."""
    ok, err = email_configured()
    if not ok:
        return {"ok": False, "sent": 0, "skipped": 0, "failed": 0, "msg": err}

    loc_email_map = _get_loc_map()
    pending = [r for r in all_location_rows if r.get("status") != "SUBMITTED"]
    if not pending:
        return {
            "ok": True, "sent": 0, "skipped": 0, "failed": 0,
            "msg": "All locations submitted. No location reminders sent.",
        }

    sent, skipped, failed, errors = 0, 0, 0, []

    if test_mode:
        r = pending[0]
        loc_code = str(r.get("userId", ""))
        to_email = loc_email_map.get(loc_code)
        if not to_email:
            return {
                "ok": False, "sent": 0, "skipped": 1, "failed": 0,
                "msg": f"[TEST] First pending location ({r.get('locName', '')}) has no email configured.",
            }
        res = send_location_pending_reminder(
            to_email, r.get("locName", ""), loc_code,
            month_year, r.get("status", "NOT_STARTED"),
            float(r.get("completion_pct", 0)), due_date,
            test_mode=True, test_email=test_email or SENDER_EMAIL,
        )
        n_with_email = len([x for x in pending if loc_email_map.get(str(x.get("userId", "")))])
        mode = res.get("mode", "sent")
        if res["ok"]:
            return {
                "ok": True, "sent": 1, "skipped": 0, "failed": 0,
                "msg": (
                    f"[TEST] Location email {mode} to {test_email or SENDER_EMAIL} "
                    f"(sample: {r.get('locName', '')}). "
                    f"In production would send to {n_with_email} location(s), "
                    f"skip {len(pending)-n_with_email} (no email)."
                ),
            }
        return {"ok": False, "sent": 0, "skipped": 0, "failed": 1,
                "msg": f"[TEST] {res.get('msg', '')}"}

    for r in pending:
        loc_code = str(r.get("userId", ""))
        to_email = loc_email_map.get(loc_code)
        if not to_email:
            skipped += 1
            continue
        res = send_location_pending_reminder(
            to_email, r.get("locName", ""), loc_code,
            month_year, r.get("status", "NOT_STARTED"),
            float(r.get("completion_pct", 0)), due_date,
        )
        if res["ok"]:
            sent += 1
        else:
            failed += 1
            errors.append(f"{r.get('locName', loc_code)}: {res.get('msg', '')}")

    all_ok = (failed == 0)
    parts = [f"Sent to {sent} location(s)"]
    if skipped:
        parts.append(f"{skipped} skipped (no email configured)")
    if failed:
        parts.append(f"{failed} failed")
    msg = ". ".join(parts) + "."
    if errors:
        msg += " Errors: " + "; ".join(errors[:3])
    return {
        "ok": all_ok, "sent": sent, "skipped": skipped,
        "failed": failed, "errors": errors, "msg": msg,
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
    "1691": "silg.top@hpcl.in",            # RANGPO TOP-IOC (East Zone)
    "1693": "silg.top@hpcl.in",            # SILIGURI location
    "1698": "rahulchangmai@hpcl.in",       # Dimapur TOP-IOC (non-operational; excluded via Settings)
    "1700": "silg.top@hpcl.in",            # RANGAPANI TOP-NRL
    "1708": "rahulchangmai@hpcl.in",       # Dimapur Depot (operational)
    "1711": "brni.irdic@hpcl.in",
    "1712": "brni.irdic@hpcl.in",
    "1723": "patna.irdic@hpcl.in",
    "1742": "bokr.irdic@hpcl.in",
    "1743": "bokr.irdic@hpcl.in",          # JASIDIH TOP-IOC
    "1747": "bokr.irdic@hpcl.in",
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
    "1420": "hazira.irdic@hpcl.in",
    "1885": "mang.tmlic@hpcl.in",
    # Skip: 1899 Ferokee TOP IOC — no email; 1882 VNKOTI TOP-IOC — not functional;
    #       1445 Borkheri TOP BPC — not functional
}


# ── Zone credential email map ──────────────────────────────────────────────────
# Maps zone name → dict with "to" and "cc" for sending zone account credentials.
# "to" = personal email of the Zone OD (from Zonal OD.msg + ZONE_EMAIL_MAP)
# Credential emails for zone accounts (BLRMIS, BHOMIS, etc.) go to these recipients.

ZONE_CREDENTIAL_MAP = {k: v for k, v in ZONE_EMAIL_MAP.items()}


# ── Dynamic email maps (sheet-backed with hardcoded fallback) ─────────────────

def _get_loc_map() -> dict:
    """Location email map — sheet first, hardcoded fallback."""
    try:
        import sheets as _sh
        loc, _ = _sh.get_email_master_maps()
        if loc:
            return loc
    except Exception:
        pass
    return LOCATION_EMAIL_MAP


def _get_zone_map() -> dict:
    """Zone email map — sheet first, hardcoded fallback."""
    try:
        import sheets as _sh
        _, zone = _sh.get_email_master_maps()
        if zone:
            return zone
    except Exception:
        pass
    return ZONE_EMAIL_MAP


def get_location_email_map() -> dict:
    """Public accessor for the location→email dict (dynamic)."""
    return _get_loc_map()


def get_zone_email_map() -> dict:
    """Public accessor for the zone→{to,cc} dict (dynamic)."""
    return _get_zone_map()


def build_credentials_email_html(
    loc_name: str,
    loc_code: str,
    password: str,
    app_url: str = "https://sodemis.streamlit.app/",
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
        "The MIS data submission for your location is due by the <strong>5th of every month</strong>"
        " (for the previous month). For any issues, raise a support ticket from within the portal"
        f" or contact <a href='mailto:{SENDER_EMAIL}' style='color:#0033A0;'>{SENDER_EMAIL}</a>.</p>"
        "</div></div></body></html>"
    )


def send_credential_email(
    to_email: str,
    loc_name: str,
    loc_code: str,
    password: str,
    app_url: str = "https://sodemis.streamlit.app/",
    cc_email: str = "",
    test_mode: bool = False,
    test_email: str = "",
) -> dict:
    """Send login credentials to a location via PowerShell + Outlook."""
    html_body  = build_credentials_email_html(loc_name, loc_code, password, app_url)
    pfx        = "[TEST] " if test_mode else ""
    subject    = f"{pfx}HPCL SOD MIS Portal — Login Credentials | {loc_name}"
    actual_to  = (test_email or SENDER_EMAIL) if test_mode else to_email
    actual_cc  = "" if test_mode else cc_email
    actual_bcc = "" if test_mode else SENDER_EMAIL
    return _send_outlook(actual_to, subject, html_body, actual_cc, actual_bcc)


# ── Forgot Password / Help Request Email ──────────────────────────────────────
ADMIN_EMAIL = "shoaibrehman@hpcl.in"


# ── HelpDesk response email ────────────────────────────────────────────────────

def _build_helpdesk_response_html(
    location_code: str,
    issue_type: str,
    issue_desc: str,
    admin_response: str,
    status: str,
) -> str:
    status_color = "#1b5e20" if status == "Resolved" else "#e65100"
    status_bg    = "#e8f5e9" if status == "Resolved" else "#fff3e0"
    return (
        "<!DOCTYPE html><html><head><meta charset='UTF-8'></head>"
        "<body style='font-family:Arial,sans-serif;margin:0;padding:0;background:#f4f6fb;'>"
        "<div style='max-width:620px;margin:30px auto;background:white;border-radius:14px;"
        "box-shadow:0 2px 12px rgba(0,0,0,0.10);overflow:hidden;'>"
        "<div style='background:#002B8F;padding:22px 30px;'>"
        "<div style='color:white;font-size:20px;font-weight:700;letter-spacing:0.5px;'>"
        "HPCL SOD &mdash; MIS Portal</div>"
        "<div style='color:#a8bfe8;font-size:12px;margin-top:4px;'>"
        "Supply, Operations &amp; Distribution &nbsp;&middot;&nbsp; Help Desk Response</div>"
        "</div>"
        "<div style='padding:28px 30px;'>"
        f"<p style='font-size:15px;color:#333;margin-top:0;'>Dear <strong>Location {location_code}</strong>,</p>"
        "<p style='font-size:14px;color:#444;line-height:1.8;'>"
        "Your support request has been reviewed. Please find the details and Admin response below.</p>"
        "<table style='width:100%;border-collapse:collapse;margin:18px 0;border-radius:8px;overflow:hidden;'>"
        "<tr style='background:#002B8F;'>"
        "<th style='padding:10px 16px;color:white;text-align:left;font-size:13px;width:36%;'>Field</th>"
        "<th style='padding:10px 16px;color:white;text-align:left;font-size:13px;'>Detail</th></tr>"
        "<tr style='background:#f5f7ff;'>"
        "<td style='padding:10px 16px;font-size:13px;font-weight:600;color:#002B8F;'>Issue Type</td>"
        f"<td style='padding:10px 16px;font-size:13px;'>{issue_type}</td></tr>"
        "<tr style='background:#ffffff;'>"
        "<td style='padding:10px 16px;font-size:13px;font-weight:600;color:#002B8F;'>Your Request</td>"
        f"<td style='padding:10px 16px;font-size:13px;'>{issue_desc}</td></tr>"
        f"<tr style='background:{status_bg};'>"
        "<td style='padding:10px 16px;font-size:13px;font-weight:600;color:#002B8F;'>Status</td>"
        f"<td style='padding:10px 16px;font-size:13px;font-weight:700;color:{status_color};'>{status}</td></tr>"
        "<tr style='background:#f5f7ff;'>"
        "<td style='padding:10px 16px;font-size:13px;font-weight:600;color:#002B8F;'>Admin Response</td>"
        f"<td style='padding:10px 16px;font-size:14px;color:#222;line-height:1.7;'>{admin_response}</td></tr>"
        "</table>"
        "<p style='font-size:14px;color:#333;margin-top:16px;line-height:1.8;'>"
        "If you have further questions, please raise a new ticket from within the portal "
        f"or contact <a href='mailto:{ADMIN_EMAIL}' style='color:#0033A0;'>{ADMIN_EMAIL}</a>.</p>"
        "</div></div></body></html>"
    )


def send_helpdesk_response_email(
    to_email: str,
    location_code: str,
    issue_type: str,
    issue_desc: str,
    admin_response: str,
    status: str = "Resolved",
) -> dict:
    """Email Admin's HelpDesk response to the location in-charge."""
    subject = (
        f"HPCL SOD MIS — Your {issue_type} is {status} | Location {location_code}"
    )
    html = _build_helpdesk_response_html(
        location_code, issue_type, issue_desc, admin_response, status
    )
    return _send_outlook(to_email, subject, html)


def send_forgot_password_email(user_id: str, issue_text: str) -> dict:
    """Send a password reset / help request email to Admin via Outlook."""
    subject = f"SOD e-MIS: Password Reset Request — User ID {user_id}"
    html_body = (
        "<html><body style='font-family:Arial,sans-serif;font-size:14px;color:#222;'>"
        "<div style='max-width:520px;margin:20px auto;border:1px solid #ddd;"
        "border-radius:8px;padding:24px;background:#f9f9f9;'>"
        "<div style='background:#001F5E;color:#fff;padding:12px 18px;"
        "border-radius:6px 6px 0 0;margin:-24px -24px 20px;'>"
        "<b>HPCL SOD e-MIS &mdash; Password Reset Request</b></div>"
        f"<p><b>User ID:</b> {user_id}</p>"
        f"<p><b>Message from user:</b><br>"
        f"<span style='background:#fff;border:1px solid #eee;display:block;"
        f"padding:10px;border-radius:4px;margin-top:4px;'>{issue_text}</span></p>"
        "<hr style='margin:20px 0;border:none;border-top:1px solid #ddd;'>"
        "<p style='font-size:12px;color:#888;'>Sent from SOD e-MIS login page.</p>"
        "</div></body></html>"
    )
    return _send_outlook(ADMIN_EMAIL, subject, html_body)
