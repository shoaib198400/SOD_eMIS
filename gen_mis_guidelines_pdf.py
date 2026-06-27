"""Generate assets/MIS_Guidelines.pdf  — run once: python gen_mis_guidelines_pdf.py"""
import os
from fpdf import FPDF
from fpdf.enums import XPos, YPos

# ── Colours ──────────────────────────────────────────────────────────────────
BLUE  = (0,  31,  94)
RED   = (204, 0,   0)
GOLD  = (198, 166, 74)
LBLUE = (236, 241, 250)
LGREY = (245, 247, 251)
MGREY = (100, 112, 133)
DARK  = (30,  30,  45)
AMBER = (146, 64,  14)
GREEN = (21,  87,  36)

# ── Section content ───────────────────────────────────────────────────────────
SECTIONS = [
    (1, "Operations", [
        "MS (MT) — Motor Spirit throughput for the month",
        "HSD (MT) — High Speed Diesel throughput",
        "Total Throughput (MT) including all other products",
        "Throughput Target (MT) — as per plan",
        "MEB (Rs Lakhs) — Monthly Expenditure Budget and % vs Budget",
        "OPEX (Rs/MT) and OPEX Target (Rs/MT)",
        "Solar Plant: Installed Capacity (KW), Generation (KWH), Grid Export (KWH)",
        "Electricity: KWH from SEB / DG / Solar, Contract Demand (KVA), Power Factor",
        "SEC (KWH/MT) — auto-calculated from total KWH divided by total throughput",
        "Stock Loss product-wise (KL and % vs Target)",
        "AIM Holds (Nos.), Day-End Report signed (Y/N), Auto-Reconciliation (%)",
        "Manpower: Management, Non-Mgmt, Contract Engineers, GATs, TATs, Security, Housekeeping",
        "E-diary observations recorded (integer, minimum 1) - count of E-diary observations for the month",
    ]),
    (2, "Finance & Planning (F&P)", [
        "VRU Available & Functioning (Yes / No / NA)",
        "No. of Approved NPCB Projects and their cumulative value (Rs Lakhs)",
        "CAPEX done (Lakhs) vs AOP Target (Lakhs)",
        "CAPITALIZATION done (Lakhs) vs AOP Target (Lakhs)",
        "Scrap Value at Location (Rs Lakhs) and Physical Scrap Disposal this month (Rs Lakhs)",
    ]),
    (3, "Supply & Distribution (S&D)", [
        "Rakes Unloaded / Loaded — count for the month",
        "Pending Railway Claims uploaded in system (Yes / No / NA)",
        "MDP Quantity Target & Actual (KL) — HPCL and OMC separately",
        "Railway Claim details: enter each claim as a separate row in the 'Railway Claims' tab",
        "IRR Details: each IRR is entered as a row in the 'IRR Details' tab — IRR date,",
        "  closure date, no. of IRRs, etc.",
    ]),
    (4, "Biofuel", [
        "EBP — Ethanol Blending Percentage (%) for the month",
        "Unblended MS Retail Sales (KL)",
        "Biodiesel Receipt Quantity (KL)",
        "Biodiesel Blending Percentage (%)",
    ]),
    (5, "Maintenance & Inspection (M&I)", [
        "--- S5 MAIN FORM ---",
        "M&I Index (0-100 scale)",
        "PM Percentage (%) — Preventive Maintenance completed vs scheduled",
        "Breakdown Equipment Details (text description, optional)",
        "Tank Cleaned this month: Tank No., Product, Capacity, Cleaning Date",
        "Tank Cleaning Extensions: eFN Number and extension date",
        "",
        "--- S5A M&I MIS (10-SUBSECTION DETAILED MODULE) ---",
        "Access via the 'S5A M&I MIS' button in the left sidebar under Section 5.",
        "Each of the 10 subsections must either have data saved OR be marked Not Applicable.",
        "All dates must be entered in DD/MM/YYYY format only.",
        "",
        "S5A-1  Tank Outage: Tank no. (from Tank Master), description, planned start &",
        "  end dates, actual start & end dates, reason for outage, current status",
        "S5A-2  Major Repair: Tank no., repair description, ETC date, current status",
        "S5A-3  VRU (Vapour Recovery Unit): Operating? (Y/N), date not operating,",
        "  ETC date for restoration, VRU installed capacity (MT/day)",
        "S5A-4  M&I Audit 2025-26: Audit date, no. of recommendations, pending count,",
        "  external audit score",
        "S5A-5  M&I Audit 2026-27: Audit carried out (Y/N), audit date, pending count,",
        "  score (fill once audit is conducted; NA if not yet due)",
        "S5A-6  Technical Audit: Audit date, auditing agency, recommendations, pending",
        "S5A-7  Equipment Breakdown: Equipment name, failure type, breakdown start date,",
        "  proposed restoration date, actual end date, current status",
        "S5A-8  Internal Pipeline: Pipeline name / segment, last UT date, last hydrotest",
        "  date, last DCVG date, last LRUT date",
        "S5A-9  External Pipeline: Same fields as internal pipeline for external segments",
        "S5A-10 Tank Status: Per tank — cleaning completed & due date, extension taken",
        "  (Y/N), extension eFN no., inspection date & due date, painting date & due date,",
        "  tank current status",
        "",
        "NOT APPLICABLE rule: If a subsection does not apply to your location (e.g., no VRU",
        "  installed, no external pipeline, no tank outage this month), tick the amber-highlighted",
        "  'Not Applicable' checkbox visible inside that tab. This marks the tab as complete",
        "  (green tick). Skipping without ticking NA will block submission.",
    ]),
    (6, "Health, Safety & Environment (HSE)", [
        "HSE Index vs Target (%)",
        "Water Consumed in Month (KL); SWC (KL/MT) — auto-calculated",
        "TT Accidents, Fatalities, Near Miss / Unsafe Acts / Unsafe Conditions",
        "Permits Issued, MOCs Created and Approved",
        "Pending MDSA & OISD Observations (%), OISD Obs > 1 Year (count)",
        "SOPs re-validated within 3 years (Yes / No / NA), Green Belt Area (SQM)",
        "CCTV: Installed (count), Non-Functional (count), % Breakdown (auto-calc), VA Exception %",
        "Awards Received (optional text field)",
    ]),
    (7, "Operational Efficiency", [
        "No. of Bays at Location and Functional Bays",
        "OLA %, COLA %, AUTO DC %, CAT-A %, LPM",
        "Operating Hours, Indents within/beyond hours, % beyond hours (auto-calc)",
        "TFMS vs SAP Performance (%), First 2 Hours TT Released per Bay",
        "Online Density %, PLC IOs Forced, Manual Bay Allocations, Sick TTs",
        "Cancelled TTs, Local Loading (KL), Cycle Time R2 to R3 (Min)",
        "TAS Open Audit Observations %, BCU vs MFM matching, Unauthorized Flow",
        "User Access as per Circular (Yes/No/NA), TT Dip Check (%)",
    ]),
    (8, "EM Lock", [
        "EM Locks Spare Inventory (count)",
        "EM Lock Spare Key Inventory (count)",
        "Dormant EM Locks (count) and Dormant Keys (count)",
        "AIM Holds (Nos.) — locks with active holds in the month",
    ]),
    (9, "Transportation", [
        "Transportation compliance metrics cross-referencing S7 TAS data",
        "TT-related loading metrics and cycle times",
        "Logistics compliance indicators as per zone guidelines",
    ]),
    (10, "Others", [
        "CAPEX & Capitalization summary (cross-referenced from S2)",
        "MDP performance summary (cross-referenced from S3)",
        "IRR Details — each IRR entered as a separate row in the 'IRR Details' Google Sheet tab;",
        "  fields include IRR date, closure date, no. of IRRs, description (DD/MM/YYYY format)",
        "Legal Cases — each case entered as a row in the 'Legal Cases' tab: case no., court,",
        "  nature of case, current status, last hearing date, next hearing date (DD/MM/YYYY)",
        "Railway Claims — pending and settled claim amounts entered in 'Railway Claims' tab",
        "Additional compliance or operational remarks",
    ]),
]

# ── Maker-Checker process ─────────────────────────────────────────────────────
PROCESS = [
    ("Step 1", "Maker fills MIS data",
     "Log in as Maker. Use the left sidebar to navigate to each of the 10 MIS sections.\n"
     "Fill all required fields (*) and click 'Save Draft' after each section.\n"
     "For Section 5, also click 'S5A M&I MIS' in the sidebar to fill all 10 M&I subsections.\n"
     "Alternatively: download the Excel template, fill it offline, and upload to auto-populate."),
    ("Step 2", "Maker submits for review",
     "Once the Dashboard shows 100% completion across all 10 sections and M&I subsections,\n"
     "the 'Submit for Review' button becomes active. Click it to submit.\n"
     "Status changes to PENDING REVIEW and the Checker is notified automatically."),
    ("Step 3", "Checker reviews data",
     "Checker logs in and clicks 'Review Data' on the Dashboard.\n"
     "All filled fields are shown in read-only view for verification.\n"
     "Cross-check values against physical records, TAS/SAP reports, and zone data.\n"
     "Zone Officers and HQO Admin can also view data and generate MIS Reports."),
    ("Step 4", "Checker approves or rejects",
     "Approve & Lock: MIS is locked. Status changes to SUBMITTED.\n"
     "Reject: Enter correction notes. MIS returns to Maker with REJECTED status.\n"
     "Maker must correct the flagged fields and re-submit (back to Step 2)."),
]

# ── Key rules ─────────────────────────────────────────────────────────────────
KEY_RULES = [
    ("Submission Deadline",
     "MIS must be submitted by the 5th of every month for the preceding month.\n"
     "e.g. April 2026 MIS must be submitted by 5-May-2026.\n"
     "Locations that miss the deadline are flagged OVERDUE in the portal."),
    ("100% Completion Required",
     "All 10 sections must be fully complete AND all 10 M&I subsections (S5A) must\n"
     "be either filled or marked 'Not Applicable'. The Submit button remains disabled\n"
     "until every required field is filled and all sections show a green tick."),
    ("M&I Not Applicable Rule",
     "In the S5A M&I MIS module, if a subsection does not apply to your location\n"
     "(e.g., no VRU installed, no external pipeline, no tank outage this month),\n"
     "tick the amber 'Not Applicable' checkbox inside that tab.\n"
     "This marks the subsection as complete. Simply leaving it blank will block submission."),
    ("Date Format — DD/MM/YYYY",
     "All date fields in the portal use DD/MM/YYYY format.\n"
     "In the Excel template, enter dates as text in DD/MM/YYYY format (e.g. 25/06/2025).\n"
     "For unknown dates in detail sheets, enter NA. Do not leave date fields blank."),
    ("Auto-Saved Drafts",
     "Data is saved as a draft. You can leave and return at any time\n"
     "before submission. Progress is preserved across sessions and logins."),
    ("Excel Template",
     "Download the MIS Excel Template from the Dashboard. It contains:\n"
     "  - 'MIS Data' sheet: main 135+ field row (row 4 is the entry row)\n"
     "  - 'Railway Claims', 'IRR Details', 'Legal Cases': one row per entry\n"
     "  - 'S5A-1 Tank Outage' through 'S5A-10 Tank Status': M&I subsection sheets\n"
     "Fill and upload to auto-populate all fields in one step."),
    ("Locked MIS",
     "Once Checker approves and locks the MIS, no edits are allowed\n"
     "without an explicit unlock request approved by the Zone Officer."),
    ("Role Restrictions",
     "Maker: fill & submit MIS for their location only.\n"
     "Checker: review & approve/reject for their location.\n"
     "Zone Officer: view all locations in zone, generate MIS Reports, unlock MIS.\n"
     "HQO Admin: full access across all zones and locations."),
    ("Modification After Rejection",
     "If Checker rejects, the MIS returns to IN PROGRESS status.\n"
     "Maker corrects the noted fields and re-submits for Checker review."),
    ("Support Ticket",
     "Use the 'Raise a Support Ticket' button in the sidebar for portal issues.\n"
     "Select the issue type, describe the problem, and submit.\n"
     "Your Zone Officer will be notified. Or email: shoaibrehman@hpcl.in"),
]

# ── Excel template guide steps ────────────────────────────────────────────────
EXCEL_STEPS = [
    ("Step 1", "Download the Template",
     "Log in to the portal and open your Dashboard.\n"
     "In the right panel, click 'Download MIS Template'.\n"
     "The Excel file downloads pre-filled with your location code, name, and month."),
    ("Step 2", "Fill the MIS Data sheet",
     "Open the Excel file. Go to the 'MIS Data' sheet — Row 4 is the data entry row.\n"
     "Fill each column corresponding to an MIS field.\n"
     "Do NOT modify column headers (rows 1-3) or any structural elements.\n"
     "Hints are shown in row 2; field names in row 1."),
    ("Step 3", "Fill Railway Claims / IRR Details / Legal Cases sheets",
     "Switch to each detail sheet and add one row per record:\n"
     "  - Railway Claims: claim type, amount, hearing dates, status\n"
     "  - IRR Details: IRR date, closure date, count, description\n"
     "  - Legal Cases: case no., court, nature, status, hearing dates\n"
     "Use DD/MM/YYYY for all date columns. Enter NA if a date is unknown."),
    ("Step 4", "Fill S5A-1 through S5A-10 sheets (M&I MIS)",
     "Each of the 10 S5A sheets corresponds to an M&I subsection:\n"
     "  S5A-1 Tank Outage | S5A-2 Major Repair | S5A-3 VRU\n"
     "  S5A-4 Audit 25-26 | S5A-5 Audit 26-27 | S5A-6 Tech. Audit\n"
     "  S5A-7 Equip. Breakdown | S5A-8 Int. Pipeline\n"
     "  S5A-9 Ext. Pipeline | S5A-10 Tank Status\n"
     "Add one row per entry. Use DD/MM/YYYY for all date columns.\n"
     "If a subsection does not apply, the sheet will already show 'Not Applicable'\n"
     "(if you ticked NA in the portal before downloading) — leave it as is."),
    ("Step 5", "Upload and verify",
     "Return to the portal Dashboard and click 'Upload Filled Template'.\n"
     "Select your saved Excel file. The portal validates and auto-populates all fields\n"
     "including all S5A M&I subsection data.\n"
     "Review the section completion grid; manually fill any fields still showing incomplete."),
    ("Step 6", "Save & Submit",
     "Once the Dashboard shows 100% across all 10 sections and all M&I tabs show green,\n"
     "click 'Submit for Review'. The Checker will be notified automatically."),
]


# ── PDF class ─────────────────────────────────────────────────────────────────
class PDF(FPDF):
    def header(self):
        self.set_fill_color(*BLUE)
        self.rect(0, 0, 210, 14, 'F')
        self.set_font('Helvetica', 'B', 10)
        self.set_text_color(255, 255, 255)
        self.set_xy(10, 3)
        self.cell(130, 8, 'HPCL SOD e-MIS  |  MIS Filling Guidelines',
                  new_x=XPos.RIGHT, new_y=YPos.TOP)
        self.set_font('Helvetica', '', 8)
        self.set_text_color(180, 200, 235)
        self.cell(60, 8, 'Supply, Operations & Distribution', align='R',
                  new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    def footer(self):
        self.set_y(-12)
        self.set_fill_color(*BLUE)
        self.rect(0, self.get_y(), 210, 12, 'F')
        self.set_font('Helvetica', '', 7.5)
        self.set_text_color(180, 200, 235)
        self.set_x(10)
        self.cell(0, 12,
                  f'Page {self.page_no()}  |  Hindustan Petroleum Corporation Limited  |  CONFIDENTIAL',
                  new_x=XPos.LMARGIN, new_y=YPos.TOP)

    def ch(self, title, color=None):
        self.set_fill_color(*(color or BLUE))
        self.set_text_color(255, 255, 255)
        self.set_font('Helvetica', 'B', 11)
        self.cell(0, 8, _s(f'  {title}'), fill=True,
                  new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln(2)
        self.set_text_color(*DARK)

    def sec(self, num, name):
        self.set_fill_color(*LBLUE)
        self.set_draw_color(*BLUE)
        self.set_line_width(0.3)
        self.set_font('Helvetica', 'B', 10)
        self.set_text_color(*BLUE)
        self.cell(0, 7, _s(f'  S{num}  {name}'), fill=True, border='LB',
                  new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.set_text_color(*DARK)
        self.ln(1)

    def bullet(self, text):
        if not text:
            self.ln(1)
            return
        # Section dividers inside bullet list (lines starting with ---)
        if text.startswith('---'):
            self.ln(1)
            self.set_x(14)
            self.set_font('Helvetica', 'B', 8)
            self.set_text_color(*BLUE)
            label = text.strip('-').strip()
            self.cell(0, 5, _s(f'  {label}'), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            self.set_text_color(*DARK)
            return
        # Sub-items indented with two spaces
        indent = 14
        marker = '-'
        if text.startswith('  '):
            indent = 20
            marker = ' '
        self.set_x(indent)
        self.set_font('Helvetica', '', 8.5)
        self.set_text_color(*DARK)
        self.cell(5, 5, marker, new_x=XPos.RIGHT, new_y=YPos.TOP)
        self.multi_cell(0, 5, _s(text.lstrip()), new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    def step(self, label, title, body):
        self.set_fill_color(*BLUE)
        self.set_text_color(255, 255, 255)
        self.set_font('Helvetica', 'B', 9)
        self.cell(22, 7, _s(f'  {label}'), fill=True,
                  new_x=XPos.RIGHT, new_y=YPos.TOP)
        self.set_fill_color(*LBLUE)
        self.set_text_color(*BLUE)
        self.cell(0, 7, _s(f'  {title}'), fill=True,
                  new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.set_x(22)
        self.set_font('Helvetica', '', 8.5)
        self.set_text_color(*DARK)
        self.multi_cell(0, 5, _s(body), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln(2)

    def kv(self, key, val, alt=False):
        bg = LGREY if alt else (255, 255, 255)
        self.set_fill_color(*bg)
        self.set_font('Helvetica', 'B', 8.5)
        self.set_text_color(*BLUE)
        self.cell(58, 6, _s(f'  {key}'), fill=True,
                  new_x=XPos.RIGHT, new_y=YPos.TOP)
        self.set_font('Helvetica', '', 8.5)
        self.set_text_color(*DARK)
        self.multi_cell(0, 6, _s(val), fill=True,
                        new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln(0.5)

    def info_box(self, text, color=None):
        """Highlighted info / tip box."""
        fg = color or AMBER
        self.set_fill_color(255, 248, 225)
        self.set_draw_color(*fg)
        self.set_line_width(0.5)
        self.set_font('Helvetica', 'I', 8.5)
        self.set_text_color(*fg)
        self.set_x(10)
        self.multi_cell(0, 5.5, _s(text), border='L', fill=True,
                        new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln(2)


# ── Build ─────────────────────────────────────────────────────────────────────
def _s(text: str) -> str:
    """Normalize common Unicode to ASCII equivalents, then strip remaining non-Latin-1."""
    text = (text
            .replace('—', ' - ')   # em dash
            .replace('–', '-')     # en dash
            .replace('‘', "'").replace('’', "'")
            .replace('“', '"').replace('”', '"')
            .replace('…', '...'))  # ellipsis
    return text.encode('latin-1', errors='replace').decode('latin-1')


def build(out: str):
    pdf = PDF()
    pdf.set_auto_page_break(auto=True, margin=16)
    pdf.set_margins(10, 16, 10)

    # ── Cover ──────────────────────────────────────────────────────────────
    pdf.add_page()
    pdf.set_fill_color(*BLUE)
    pdf.rect(0, 16, 210, 56, 'F')
    pdf.set_xy(12, 24)
    pdf.set_font('Helvetica', 'B', 20)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(0, 11, 'HPCL SOD e-MIS', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(12)
    pdf.set_font('Helvetica', 'B', 14)
    pdf.set_text_color(*GOLD)
    pdf.cell(0, 9, 'MIS Data Filling Guidelines', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(12)
    pdf.set_font('Helvetica', '', 9.5)
    pdf.set_text_color(180, 200, 235)
    pdf.cell(0, 7,
             'Supply, Operations & Distribution  |  Hindustan Petroleum Corporation Limited',
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    tiles = [
        ('Submission Deadline', '5th of every month'),
        ('Total MIS Sections',  '10 sections'),
        ('Total Fields',        '135+ fields'),
        ('M&I Subsections',     '10 S5A subsections'),
    ]
    for i, (k, v) in enumerate(tiles):
        pdf.set_xy(10 + (i % 2) * 96, 84 + (i // 2) * 14)
        pdf.set_fill_color(*LBLUE)
        pdf.set_draw_color(*BLUE)
        pdf.set_line_width(0.3)
        pdf.set_font('Helvetica', 'B', 9)
        pdf.set_text_color(*BLUE)
        pdf.cell(94, 12, f'  {k}:  {v}', fill=True, border=1,
                 new_x=XPos.RIGHT, new_y=YPos.TOP)
    pdf.set_xy(10, 115)
    pdf.set_font('Helvetica', 'I', 8)
    pdf.set_text_color(*MGREY)
    pdf.cell(0, 6, 'Document auto-generated by HPCL SOD e-MIS Portal',
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # ── Process flow ───────────────────────────────────────────────────────
    pdf.add_page()
    pdf.ch('1.  SUBMISSION DEADLINE & PROCESS FLOW')
    pdf.set_font('Helvetica', 'B', 9.5)
    pdf.set_text_color(*AMBER)
    pdf.cell(0, 6, '  SUBMISSION DEADLINE', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font('Helvetica', '', 9)
    pdf.set_text_color(*DARK)
    pdf.set_x(10)
    pdf.multi_cell(0, 5, _s(
        'MIS data for the preceding month must be submitted by the 5th of every month.\n'
        'Example: April 2026 MIS must be submitted by 5-May-2026.\n'
        'Locations that miss the deadline will be flagged OVERDUE in the portal.'),
        new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(4)
    pdf.set_font('Helvetica', 'B', 9.5)
    pdf.set_text_color(*BLUE)
    pdf.cell(0, 6, '  MAKER-CHECKER PROCESS FLOW', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(2)
    for lbl, title, body in PROCESS:
        pdf.step(lbl, title, body)

    # ── Key rules ──────────────────────────────────────────────────────────
    pdf.ln(3)
    pdf.ch('2.  KEY RULES & IMPORTANT NOTES', RED)
    for i, (k, v) in enumerate(KEY_RULES):
        pdf.kv(k, v, alt=(i % 2 == 0))

    # ── Section guide ──────────────────────────────────────────────────────
    pdf.add_page()
    pdf.ch('3.  SECTION-BY-SECTION FILLING GUIDE')
    pdf.set_font('Helvetica', 'I', 8.5)
    pdf.set_text_color(*MGREY)
    pdf.multi_cell(0, 5, _s(
        'All 10 sections must be 100% complete before submission. '
        'Required fields are marked with * in the portal. '
        'Fields marked [auto-calc] are computed automatically - do not override. '
        'Section 5 also has a dedicated M&I MIS module (S5A) with 10 subsections.'),
        new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(3)
    for num, name, fields in SECTIONS:
        if pdf.get_y() > 240:
            pdf.add_page()
        pdf.sec(num, name)
        for f in fields:
            pdf.bullet(f)
        pdf.ln(2)

    # ── M&I MIS subsection detail page ────────────────────────────────────
    pdf.add_page()
    pdf.ch('4.  M&I MIS MODULE — S5A DETAILED GUIDE')

    pdf.info_box(
        'The M&I MIS module (S5A) is separate from the main S5 form. '
        'Access it using the "S5A  M&I MIS" button in the left sidebar under Section 5. '
        'It has 10 subsections — all must be complete (green tick) before the '
        'M&I Report can be generated.',
        color=BLUE,
    )

    mi_tabs = [
        ("S5A-1", "Tank Outage",
         "One row per tank currently under outage or planned outage.\n"
         "Fields: Tank No. (from Tank Master), Other Tank Description, Planned Start Date,\n"
         "Planned End Date, Actual Start Date, Actual End Date, Outage Reason, Current Status.\n"
         "If no tank is under outage this month, tick 'Not Applicable'."),
        ("S5A-2", "Major Repair",
         "One row per major repair in progress or completed this month.\n"
         "Fields: Tank No., Repair Description, ETC Date, Current Status.\n"
         "Tick 'Not Applicable' if no major repair this month."),
        ("S5A-3", "VRU (Vapour Recovery Unit)",
         "Fields: VRU Operating (Yes/No), Date Not Operating, ETC Date for Restoration,\n"
         "VRU Installed Capacity (MT/day), Additional Remarks.\n"
         "Tick 'Not Applicable - VRU not installed at this location' if no VRU exists."),
        ("S5A-4", "M&I Audit 2025-26",
         "Enter once at the beginning of the year; update pending count monthly.\n"
         "Fields: Audit Date, No. of Recommendations, Pending Recommendations, External Score.\n"
         "Tick 'Not Applicable' if audit has not been conducted at this location."),
        ("S5A-5", "M&I Audit 2026-27",
         "Audit may not have been carried out yet — update when done.\n"
         "Fields: Audit Carried Out (Yes/No), Audit Date, Pending Count, Score.\n"
         "Tick 'Not Applicable' if not applicable for this location."),
        ("S5A-6", "Technical Audit",
         "Fields: Audit Date, Auditing Agency, No. of Recommendations, Pending Count.\n"
         "Tick 'Not Applicable - No technical audit this month' if none conducted."),
        ("S5A-7", "Equipment Breakdown",
         "One row per breakdown event. Fields: Equipment Name, Failure Type,\n"
         "Breakdown Start Date, Proposed Restoration Date, Actual End Date, Status.\n"
         "Tick 'Not Applicable' if no equipment breakdown this month."),
        ("S5A-8", "Internal Pipeline",
         "Fields: Pipeline Name/Segment, Last UT Date, Last Hydrotest Date,\n"
         "Last DCVG Survey Date, Last LRUT Date.\n"
         "Tick 'Not Applicable - No internal pipeline at this location' if none exists."),
        ("S5A-9", "External Pipeline",
         "Same fields as S5A-8 for external pipeline segments.\n"
         "Tick 'Not Applicable - No external pipeline at this location' if none exists."),
        ("S5A-10", "Tank Status",
         "One row per tank in the Tank Master for this location.\n"
         "Fields: Tank No., Cleaning Completed Date, Cleaning Due Date, Extension Taken (Y/N),\n"
         "Extension eFN No., Inspection Date, Inspection Due Date, Painting Date,\n"
         "Painting Due Date, Tank Status, Other Status remarks.\n"
         "All tanks from Tank Master must have a row. Tick 'Not Applicable' only if\n"
         "this location has no tanks at all."),
    ]

    for tab_id, tab_name, desc in mi_tabs:
        if pdf.get_y() > 245:
            pdf.add_page()
        pdf.set_fill_color(*LBLUE)
        pdf.set_draw_color(*BLUE)
        pdf.set_line_width(0.3)
        pdf.set_font('Helvetica', 'B', 9.5)
        pdf.set_text_color(*BLUE)
        pdf.cell(0, 7, _s(f'  {tab_id}  {tab_name}'), fill=True, border='LB',
                 new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_text_color(*DARK)
        pdf.ln(1)
        pdf.set_x(14)
        pdf.set_font('Helvetica', '', 8.5)
        pdf.multi_cell(0, 5, _s(desc), new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.ln(2)

    pdf.info_box(
        'DATE FORMAT: All dates in S5A must be entered as DD/MM/YYYY '
        '(e.g. 25/06/2025). The portal enforces this format. '
        'In the Excel template, enter dates as text in DD/MM/YYYY. '
        'For unknown dates, enter "NA" — do not leave the cell blank.',
        color=AMBER,
    )

    # ── Excel template guide ───────────────────────────────────────────────
    pdf.add_page()
    pdf.ch('5.  USING THE EXCEL TEMPLATE')

    pdf.info_box(
        'The Excel template has 14 sheets: MIS Data (main fields), '
        'Railway Claims, IRR Details, Legal Cases, and S5A-1 through S5A-10 (M&I subsections). '
        'Fill all applicable sheets before uploading.',
        color=BLUE,
    )

    for lbl, title, body in EXCEL_STEPS:
        pdf.step(lbl, title, body)

    # ── Contact & support ──────────────────────────────────────────────────
    pdf.ln(2)
    pdf.ch('6.  CONTACT & SUPPORT')
    pdf.set_font('Helvetica', '', 9)
    pdf.set_text_color(*DARK)
    pdf.multi_cell(0, 5.5, _s(
        'RAISING A SUPPORT TICKET:\n'
        '  Use the "Raise a Support Ticket" button in the left sidebar of the portal.\n'
        '  Select the issue type (Login Issue, Data Entry, Upload Error, etc.),\n'
        '  describe the problem, and click Submit. Your Zone Officer will be notified.\n\n'
        'FOR LOGIN ISSUES OR DATA UNLOCK:\n'
        '  Contact your Zone Officer directly - Zone users have unlock access in the portal.\n'
        '  Or email the HQO SOD Admin: shoaibrehman@hpcl.in\n\n'
        'FOR TECHNICAL ISSUES (errors, slow loading, blank screens):\n'
        '  Report to IT Helpdesk with a screenshot of the error message.\n\n'
        'This document is auto-generated by the HPCL SOD e-MIS Portal. '
        'For the latest version, download again from the MIS Guidelines link in the Dashboard.'),
        new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.output(out)
    sz = os.path.getsize(out)
    print(f'PDF saved: {out}  ({sz:,} bytes, {pdf.page} pages)')


if __name__ == '__main__':
    assets = os.path.join(os.path.dirname(__file__), 'assets')
    os.makedirs(assets, exist_ok=True)
    build(os.path.join(assets, 'MIS_Guidelines.pdf'))
