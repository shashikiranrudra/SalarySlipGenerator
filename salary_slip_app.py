"""
Salary Slip Generator - Professional Desktop Application
Generates salary slips from Excel sheets matching format
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
import hashlib
import json
import os
import sys
import shutil
from pathlib import Path
import threading
from datetime import datetime, date

# ── PDF generation
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                Paragraph, Spacer, HRFlowable, Image as RLImage)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfbase import pdfmetrics

# ── Excel reading
import openpyxl
import pandas as pd

# ── Image handling
from PIL import Image, ImageTk

# ────────────────────────────────────────────────
#  CONFIG / PERSISTENCE
# ────────────────────────────────────────────────
APP_DIR = Path.home() / ".salary_slip_app"
CONFIG_FILE = APP_DIR / "config.json"
APP_DIR.mkdir(exist_ok=True)

# ────────────────────────────────────────────────
#  TRIAL / EXPIRY SETTINGS
# ────────────────────────────────────────────────
TRIAL_DAYS = 90
TRIAL_FILE  = APP_DIR / ".install_date"    # hidden file stores install timestamp


def _get_or_create_install_date() -> date:
    """Return the installation date. Creates the marker file on first run."""
    if TRIAL_FILE.exists():
        try:
            raw = TRIAL_FILE.read_text().strip()
            return datetime.strptime(raw, "%Y-%m-%d").date()
        except Exception:
            pass
    # First run – record today
    today = date.today()
    TRIAL_FILE.write_text(today.strftime("%Y-%m-%d"))
    return today


def check_trial() -> dict:
    """
    Returns a dict:
      { "expired": bool, "days_used": int, "days_left": int,
        "install_date": str, "expiry_date": str }
    """
    install_date  = _get_or_create_install_date()
    expiry_date   = date.fromordinal(install_date.toordinal() + TRIAL_DAYS)
    today         = date.today()
    days_used     = (today - install_date).days
    days_left     = max(0, TRIAL_DAYS - days_used)
    return {
        "expired":      today >= expiry_date,
        "days_used":    days_used,
        "days_left":    days_left,
        "install_date": install_date.strftime("%d-%b-%Y"),
        "expiry_date":  expiry_date.strftime("%d-%b-%Y"),
    }


DEFAULT_CONFIG = {
    "password_hash": hashlib.sha256("admin123".encode()).hexdigest(),
    "company_name": "",
    "company_address": "",
    "company_logo": "",
    "output_folder": str(Path.home() / "SalarySlips"),
    "accent_color": "#1a3a6b",
    "column_map": {
        "employee_name": "Employee Name",
        "employee_code": "Employee Code",
        "department": "Department",
        "designation": "Designation",
        "location": "Location",
        "pf_no": "PF No",
        "uan_no": "UAN No",
        "esic_no": "ESIC No",
        "pan_no": "PAN No",
        "aadhar_no": "Aadhar No",
        "bank_name": "Bank Name",
        "bank_ac": "Bank A/c No",
        "monthly_salary": "Monthly Salary",
        "basic": "Basic",
        "conveyance": "Conveyance",
        "hra": "House Rent Allowance",
        "post_allowance": "Post Allowance",
        "other_allowance": "Other Allowance",
        "profession_tax": "Profession Tax",
        "pf": "Provident Fund",
        "esic_deduction": "ESIC Deduction",
        "other_deduction": "Other Deduction",
        "month_days": "Month Days",
        "weekly_off": "Weekly Off",
        "holidays": "Holidays",
        "present_days": "Present Days",
        "pay_days": "Pay Days",
        "leave_without_pay": "Leave Without Pay",
        "month": "Month",
        "year": "Year",
    }
}


def load_config():
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE) as f:
                cfg = json.load(f)
            # merge missing keys
            for k, v in DEFAULT_CONFIG.items():
                if k not in cfg:
                    cfg[k] = v
            return cfg
        except Exception:
            pass
    return DEFAULT_CONFIG.copy()


def save_config(cfg):
    with open(CONFIG_FILE, "w") as f:
        json.dump(cfg, f, indent=2)


def hash_password(pw):
    return hashlib.sha256(pw.encode()).hexdigest()


def num2words_indian(n):
    """Convert integer to Indian English words."""
    ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine',
            'Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen',
            'Seventeen', 'Eighteen', 'Nineteen']
    tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety']

    def two_digits(n):
        if n < 20:
            return ones[n]
        return tens[n // 10] + (' ' + ones[n % 10] if n % 10 else '')

    def three_digits(n):
        if n >= 100:
            return ones[n // 100] + ' Hundred' + (' ' + two_digits(n % 100) if n % 100 else '')
        return two_digits(n)

    if n == 0:
        return 'Zero'
    result = ''
    if n >= 10000000:
        result += three_digits(n // 10000000) + ' Crore '
        n %= 10000000
    if n >= 100000:
        result += three_digits(n // 100000) + ' Lakh '
        n %= 100000
    if n >= 1000:
        result += three_digits(n // 1000) + ' Thousand '
        n %= 1000
    if n > 0:
        result += three_digits(n)
    return result.strip()


def amount_to_words(amount):
    try:
        amount = round(float(amount))
        rupees = int(amount)
        return f"Rupees {num2words_indian(rupees)} Only"
    except Exception:
        return ""


# ────────────────────────────────────────────────
#  PDF GENERATOR
# ────────────────────────────────────────────────
def generate_salary_slip_pdf(data: dict, config: dict, output_path: str):
    """Generate one salary slip PDF matching the format."""
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=12*mm, bottomMargin=12*mm
    )

    accent = config.get("accent_color", "#1a3a6b")
    # convert hex to reportlab color
    def hex_to_rl(h):
        h = h.lstrip("#")
        r, g, b = int(h[0:2],16)/255, int(h[2:4],16)/255, int(h[4:6],16)/255
        return colors.Color(r, g, b)

    ACCENT = hex_to_rl(accent)
    LIGHT_GRAY = colors.Color(0.94, 0.94, 0.94)
    WHITE = colors.white
    BLACK = colors.black
    DARK_GRAY = colors.Color(0.3, 0.3, 0.3)

    story = []
    styles = getSampleStyleSheet()

    # ── HEADER ──────────────────────────────────
    company_name = config.get("company_name", "COMPANY NAME")
    company_addr = config.get("company_address", "Company Address")
    logo_path = config.get("company_logo", "")

    header_data = []
    logo_cell = ""
    if logo_path and os.path.exists(logo_path):
        try:
            logo_cell = RLImage(logo_path, width=24*mm, height=20*mm)
        except Exception:
            logo_cell = ""

    title_style = ParagraphStyle("title", fontSize=14, fontName="Helvetica-Bold",
                                  alignment=TA_CENTER, textColor=ACCENT)
    addr_style  = ParagraphStyle("addr",  fontSize=8,  fontName="Helvetica",
                                  alignment=TA_CENTER, textColor=DARK_GRAY, leading=11)
    slip_style  = ParagraphStyle("slip",  fontSize=9,  fontName="Helvetica-Bold",
                                  alignment=TA_CENTER, textColor=WHITE)

    month_year = data.get("month_year", "")
    header_inner = [
        Paragraph(company_name, title_style),
	Spacer(1, 3*mm),
        Paragraph(company_addr, addr_style),
        Paragraph(f"PAY SLIP : {month_year}", ParagraphStyle("slip2", fontSize=9,
                  fontName="Helvetica-Bold", alignment=TA_CENTER, textColor=ACCENT)),
    ]

    if logo_cell:
        header_table = Table([[logo_cell, header_inner]], colWidths=[22*mm, None])
        header_table.setStyle(TableStyle([
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("ALIGN", (1,0), (1,0), "CENTER"),
            ("LINEBELOW", (0,0), (-1,-1), 0.5, ACCENT),
        ]))
    else:
        header_table = Table([[header_inner]], colWidths=[None])
        header_table.setStyle(TableStyle([
            ("LINEBELOW", (0,0), (-1,-1), 1, ACCENT),
        ]))
    story.append(header_table)
    story.append(Spacer(1, 3*mm))

    # ── EMPLOYEE INFO ────────────────────────────
    lbl = ParagraphStyle("lbl", fontSize=8, fontName="Helvetica-Bold", textColor=DARK_GRAY)
    val = ParagraphStyle("val", fontSize=8, fontName="Helvetica", textColor=BLACK)

    def lv(label, value):
        return [Paragraph(label, lbl), Paragraph(str(value) if value else "-", val)]

    left_col = [
        lv("Employee Name", data.get("employee_name", "")),
        lv("Employee Code", data.get("employee_code", "")),
        lv("Department",    data.get("department", "")),
        lv("Designation",   data.get("designation", "")),
        lv("Location",      data.get("location", "")),
    ]
    right_col = [
        lv("PF No.",        data.get("pf_no", "")),
        lv("UAN No.",       data.get("uan_no", "")),
        lv("ESIC No.",      data.get("esic_no", "")),
        lv("PAN No.",       data.get("pan_no", "")),
        lv("Aadhar No.",    data.get("aadhar_no", "")),
    ]
    right_col2 = [
        lv("Monthly Salary Rs.", data.get("monthly_salary", "")),
        lv("Bank Name",        str(data.get("bank_name",""))),
        lv("Bank A/c No.",     str(data.get("bank_ac",""))),
        ["", ""],
        ["", ""],
    ]

    def build_kv_rows(col):
        rows = []
        for item in col:
            rows.append(item)
        return rows

    # Merge into 3-section table
    emp_rows = []
    for i in range(5):
        lc = left_col[i]
        rc = right_col[i]
        rc2 = right_col2[i]
        emp_rows.append([lc[0], lc[1], rc[0], rc[1], rc2[0], rc2[1]])

    emp_table = Table(emp_rows, colWidths=[28*mm, 42*mm, 22*mm, 38*mm, 28*mm, 22*mm])
    emp_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), LIGHT_GRAY),
        ("GRID", (0,0), (-1,-1), 0.3, colors.Color(0.8,0.8,0.8)),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0,0), (-1,-1), [WHITE, LIGHT_GRAY]),
        ("LEFTPADDING", (0,0), (-1,-1), 3),
        ("RIGHTPADDING", (0,0), (-1,-1), 3),
        ("TOPPADDING", (0,0), (-1,-1), 2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
        ("LINEAFTER", (1,0), (1,-1), 0.5, colors.Color(0.7,0.7,0.7)),
        ("LINEAFTER", (3,0), (3,-1), 0.5, colors.Color(0.7,0.7,0.7)),
    ]))
    story.append(emp_table)
    story.append(Spacer(1, 3*mm))

    # ── EARNINGS & DEDUCTIONS HEADER ────────────
    hdr_style = ParagraphStyle("hdr", fontSize=9, fontName="Helvetica-Bold",
                                alignment=TA_CENTER, textColor=WHITE)
    hdr_row = [
        Paragraph("Earnings", hdr_style), Paragraph("Amount", hdr_style),
        Paragraph("Deductions", hdr_style), Paragraph("Amount", hdr_style),
    ]
    ed_header = Table([hdr_row], colWidths=[65*mm, 25*mm, 65*mm, 25*mm])
    ed_header.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), ACCENT),
        ("GRID", (0,0), (-1,-1), 0.3, WHITE),
        ("TOPPADDING", (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
    ]))
    story.append(ed_header)

    # ── EARNINGS & DEDUCTIONS ROWS ───────────────
    earn_items = [
        ("Basic",                 data.get("basic", 0)),
        ("Conveyance",            data.get("conveyance", 0)),
        ("House Rent Allowance",  data.get("hra", 0)),
        ("Post Allowance",        data.get("post_allowance", 0)),
    ]
    if data.get("other_allowance"):
        earn_items.append(("Other Allowance", data.get("other_allowance", 0)))

    dedn_items = [
        ("Profession Tax",  data.get("profession_tax", 0)),
        ("Provident Fund",  data.get("pf", 0)),
    ]
    if data.get("esic_deduction"):
        dedn_items.append(("ESIC",           data.get("esic_deduction", 0)))
    if data.get("other_deduction"):
        dedn_items.append(("Other Deduction", data.get("other_deduction", 0)))

    rows_count = max(len(earn_items), len(dedn_items))
    while len(earn_items) < rows_count:
        earn_items.append(("", ""))
    while len(dedn_items) < rows_count:
        dedn_items.append(("", ""))

    item_lbl = ParagraphStyle("il", fontSize=8, fontName="Helvetica", textColor=BLACK)
    item_amt = ParagraphStyle("ia", fontSize=8, fontName="Helvetica",
                               alignment=TA_RIGHT, textColor=BLACK)

    def fmt_amt(v):
        try:
            f = float(v)
            return f"{f:,.2f}" if f else ""
        except Exception:
            return str(v) if v else ""

    ed_rows = []
    for i in range(rows_count):
        el, ea = earn_items[i]
        dl, da = dedn_items[i]
        row = [
            Paragraph(el, item_lbl),
            Paragraph(fmt_amt(ea), item_amt),
            Paragraph(dl, item_lbl),
            Paragraph(fmt_amt(da), item_amt),
        ]
        ed_rows.append(row)

    ed_table = Table(ed_rows, colWidths=[65*mm, 25*mm, 65*mm, 25*mm])
    ed_table.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.3, colors.Color(0.85,0.85,0.85)),
        ("ROWBACKGROUNDS", (0,0), (-1,-1), [WHITE, LIGHT_GRAY]),
        ("LEFTPADDING", (0,0), (-1,-1), 4),
        ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("TOPPADDING", (0,0), (-1,-1), 2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
        ("LINEAFTER", (1,0), (1,-1), 0.5, ACCENT),
    ]))
    story.append(ed_table)

    # ── TOTALS ROW ───────────────────────────────
    total_earn = sum(float(v) for _, v in earn_items if v and str(v).replace('.','').isdigit() or (isinstance(v, (int, float)) and v))
    total_dedn = sum(float(v) for _, v in dedn_items if v and str(v).replace('.','').isdigit() or (isinstance(v, (int, float)) and v))

    try:
        total_earn = sum(float(v) for _, v in earn_items[:] if v not in ("", None))
    except Exception:
        total_earn = 0
    try:
        total_dedn = sum(float(v) for _, v in dedn_items[:] if v not in ("", None))
    except Exception:
        total_dedn = 0

    net_salary = total_earn - total_dedn

    tot_style = ParagraphStyle("tot", fontSize=9, fontName="Helvetica-Bold", textColor=WHITE)
    tot_amt   = ParagraphStyle("tota", fontSize=9, fontName="Helvetica-Bold",
                                alignment=TA_RIGHT, textColor=WHITE)

    totals_row = Table([[
        Paragraph("Earnings Total Rs. :", tot_style),
        Paragraph(fmt_amt(total_earn), tot_amt),
        Paragraph("Deduction Total Rs. :", tot_style),
        Paragraph(fmt_amt(total_dedn), tot_amt),
    ]], colWidths=[65*mm, 25*mm, 65*mm, 25*mm])
    totals_row.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), ACCENT),
        ("GRID", (0,0), (-1,-1), 0.3, WHITE),
        ("TOPPADDING", (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ("LEFTPADDING", (0,0), (-1,-1), 4),
        ("RIGHTPADDING", (0,0), (-1,-1), 4),
    ]))
    story.append(totals_row)

    # Net Salary row
    net_row = Table([[
        Paragraph("Net Salary Rs. :", ParagraphStyle("ns", fontSize=9,
                   fontName="Helvetica-Bold", alignment=TA_RIGHT, textColor=ACCENT)),
        Paragraph(fmt_amt(net_salary), ParagraphStyle("nsa", fontSize=10,
                   fontName="Helvetica-Bold", alignment=TA_RIGHT, textColor=ACCENT)),
    ]], colWidths=[140*mm, 40*mm])
    net_row.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), LIGHT_GRAY),
        ("LINEBELOW", (0,0), (-1,-1), 1, ACCENT),
        ("TOPPADDING", (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ("RIGHTPADDING", (0,0), (-1,-1), 4),
    ]))
    story.append(net_row)

    # Net in words
    words_row = Table([[
        Paragraph(f"Net Salary Rs In Words :- {amount_to_words(net_salary)}",
                  ParagraphStyle("wrd", fontSize=8.5, fontName="Helvetica-Bold",
                                  textColor=BLACK)),
    ]], colWidths=[None])
    words_row.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), LIGHT_GRAY),
        ("LINEBELOW", (0,0), (-1,-1), 0.5, ACCENT),
        ("TOPPADDING", (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ("LEFTPADDING", (0,0), (-1,-1), 4),
    ]))
    story.append(words_row)
    story.append(Spacer(1, 2*mm))

    # ── ATTENDANCE SUMMARY ───────────────────────
    att_hdr = ParagraphStyle("ah", fontSize=8, fontName="Helvetica-Bold",
                              alignment=TA_CENTER, textColor=WHITE)
    att_val = ParagraphStyle("av", fontSize=8, fontName="Helvetica",
                              alignment=TA_CENTER, textColor=BLACK)

    att_labels = ["PL Op.Bal", "Leave this month", "Total Lv Taken", "Encashed",
                  "Cl.Bal", "Leave wo.pay", "CO Op.Bal", "C Off", "Total C Off",
                  "Encashed", "Cl.Bal"]
    att_data = ["", "", "0.00", "0.00", "", fmt_amt(data.get("leave_without_pay",0)),
                "", "0.00", "0.00", "0.00", ""]

    att_hdr_row = [Paragraph(l, att_hdr) for l in att_labels]
    att_val_row = [Paragraph(v, att_val) for v in att_data]

    att_table = Table([att_hdr_row, att_val_row])
    att_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), ACCENT),
        ("BACKGROUND", (0,1), (-1,1), LIGHT_GRAY),
        ("GRID", (0,0), (-1,-1), 0.3, colors.Color(0.8,0.8,0.8)),
 	("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("TOPPADDING", (0,0), (-1,-1), 2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
    ]))
    story.append(att_table)
    story.append(Spacer(1, 2*mm))

    # ── DAYS SUMMARY ────────────────────────────
    days_labels = ["Month Day", "Weekly off", "Holiday", "Present Days", "Pay Days"]
    days_vals   = [
        str(data.get("month_days", "")),
        str(data.get("weekly_off", "")),
        str(data.get("holidays", "")),
        str(data.get("present_days", "")),
        str(data.get("pay_days", "")),
    ]
    d_hdr = [Paragraph(l, att_hdr) for l in days_labels]
    d_val = [Paragraph(v, att_val) for v in days_vals]

    days_table = Table([d_hdr, d_val])
    days_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), ACCENT),
        ("BACKGROUND", (0,1), (-1,1), LIGHT_GRAY),
        ("GRID", (0,0), (-1,-1), 0.3, colors.Color(0.8,0.8,0.8)),
        ("TOPPADDING", (0,0), (-1,-1), 2),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
    ]))
    story.append(days_table)

    doc.build(story)


# ────────────────────────────────────────────────
#  EXCEL READER
# ────────────────────────────────────────────────
def read_excel_salaries(path: str, col_map: dict) -> list[dict]:
    """Read salary data from Excel using the column mapping."""
    df = pd.read_excel(path, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    # Reverse map: display_name -> internal_key
    rev = {v.strip(): k for k, v in col_map.items()}

    records = []
    for _, row in df.iterrows():
        rec = {}
        for col in df.columns:
            internal = rev.get(col)
            if internal:
                val = row.get(col, "")
                rec[internal] = "" if (val is None or str(val).lower() in ("nan","none","")) else str(val).strip()
        # month_year
        month = rec.get("month", "")
        year  = rec.get("year", "")
        rec["month_year"] = f"{month}-{year}".strip("-")
        records.append(rec)
    return records


# ────────────────────────────────────────────────
#  MAIN APPLICATION
# ────────────────────────────────────────────────
class SalarySlipApp:
    def __init__(self):
        self.config = load_config()
        self.authenticated = False
        self.records = []
        self.excel_path = ""

        self.root = tk.Tk()
        self.root.title("Salary Slip Generator")
        self.root.geometry("1100x720")
        self.root.minsize(900, 600)
        self._set_icon()

        self._apply_theme()
        trial = check_trial()
        if trial["expired"]:
            self._show_expired(trial)
        else:
            self._show_trial_banner(trial)
            self._show_login()
        self.root.mainloop()

    def _set_icon(self):
        try:
            self.root.iconbitmap(default='')
        except Exception:
            pass

    def _apply_theme(self):
        accent = self.config.get("accent_color", "#1a3a6b")
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TFrame", background="#f4f6fa")
        style.configure("TLabel", background="#f4f6fa", font=("Segoe UI", 10))
        style.configure("Header.TLabel", background="#f4f6fa",
                        font=("Segoe UI", 18, "bold"), foreground=accent)
        style.configure("Sub.TLabel", background="#f4f6fa",
                        font=("Segoe UI", 9), foreground="#666")
        style.configure("TButton", font=("Segoe UI", 10), padding=(12, 6))
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"),
                        background=accent, foreground="white")
        style.map("Accent.TButton",
                  background=[("active", "#2a5298"), ("pressed", "#0f2a55")])
        style.configure("Danger.TButton", font=("Segoe UI", 10, "bold"),
                        background="#c0392b", foreground="white")
        style.configure("TEntry", font=("Segoe UI", 10), padding=6)
        style.configure("TNotebook", background="#f4f6fa")
        style.configure("TNotebook.Tab", font=("Segoe UI", 10), padding=(16,6))
        style.configure("Treeview", font=("Segoe UI", 9), rowheight=26)
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"),
                        background=accent, foreground="white")
        style.map("Treeview.Heading", background=[("active", accent)])
        style.configure("Card.TFrame", background="white",
                        relief="flat", borderwidth=1)

    # ── TRIAL BANNER (shown while valid) ──────────
    def _show_trial_banner(self, trial: dict):
        """Show a small non-blocking banner on the root window about days remaining."""
        days_left = trial["days_left"]
        expiry    = trial["expiry_date"]
        accent    = self.config.get("accent_color", "#1a3a6b")

        if days_left <= 10:
            bg_color = "#c0392b"   # urgent red
            msg = f"⚠️  Trial expires in {days_left} day(s)  ({expiry})  —  Please contact your administrator."
        elif days_left <= 30:
            bg_color = "#e67e22"   # warning orange
            msg = f"⚠️  Trial expires in {days_left} day(s)  ({expiry})"
        else:
            bg_color = "#27ae60"   # healthy green
            msg = f"✅  Trial active — {days_left} day(s) remaining  (expires {expiry})"

        banner = tk.Frame(self.root, bg=bg_color, height=28)
        banner.pack(side="bottom", fill="x")
        tk.Label(banner, text=msg, bg=bg_color, fg="white",
                 font=("Segoe UI", 9, "bold")).pack(pady=4)

    # ── EXPIRED LOCK SCREEN ────────────────────
    def _show_expired(self, trial: dict):
        """Replace login with a permanent lock screen — no bypass."""
        self.root.configure(bg="#1a1a2e")
        self.root.title("Salary Slip Generator – EXPIRED")
        accent = self.config.get("accent_color", "#1a3a6b")

        outer = tk.Frame(self.root, bg="#1a1a2e")
        outer.place(relx=0.5, rely=0.5, anchor="center")

        # Lock icon
        tk.Label(outer, text="🔒", bg="#1a1a2e",
                 font=("Segoe UI", 64)).pack(pady=(0, 8))

        tk.Label(outer, text="Trial Period Expired",
                 bg="#1a1a2e", fg="#e74c3c",
                 font=("Segoe UI", 22, "bold")).pack()

        tk.Label(outer,
                 text=(
                     f"Installed on : {trial['install_date']}\n"
                     f"Expired on   : {trial['expiry_date']}\n\n"
                     "Your 90-day trial has ended.\n"
                     "Please contact your administrator to renew."
                 ),
                 bg="#1a1a2e", fg="#ecf0f1",
                 font=("Segoe UI", 11), justify="center",
                 pady=10).pack()

        tk.Frame(outer, bg="#e74c3c", height=2, width=340).pack(pady=12)

        tk.Label(outer,
                 text="Salary Slip Generator  ·  By Shashikiran Rudrawar",
                 bg="#1a1a2e", fg="#666",
                 font=("Segoe UI", 8)).pack()

        # Disable close-and-reopen trick: closing just quits entirely
        self.root.protocol("WM_DELETE_WINDOW", self.root.destroy)

    # ── LOGIN ──────────────────────────────────
    def _show_login(self):
        self.root.configure(bg="#f4f6fa")
        login_win = tk.Toplevel(self.root)
        login_win.title("Login – Salary Slip Generator")
        login_win.geometry("420x480")
        login_win.resizable(False, False)
        login_win.grab_set()
        login_win.configure(bg="#f4f6fa")
        login_win.protocol("WM_DELETE_WINDOW", self.root.destroy)

        # Center
        login_win.update_idletasks()
        x = (login_win.winfo_screenwidth() - 420) // 2
        y = (login_win.winfo_screenheight() - 480) // 2
        login_win.geometry(f"+{x}+{y}")

        accent = self.config.get("accent_color", "#1a3a6b")

        # Top banner
        banner = tk.Frame(login_win, bg=accent, height=120)
        banner.pack(fill="x")
        tk.Label(banner, text="💼", bg=accent, font=("Segoe UI", 36)).pack(pady=(20,4))
        tk.Label(banner, text="Salary Slip Generator",
                 bg=accent, fg="white", font=("Segoe UI", 14, "bold")).pack()

        body = tk.Frame(login_win, bg="#f4f6fa")
        body.pack(fill="both", expand=True, padx=40, pady=30)

        tk.Label(body, text="Enter Password", bg="#f4f6fa",
                 font=("Segoe UI", 11, "bold"), fg="#333").pack(anchor="w", pady=(0,6))

        self._pw_var = tk.StringVar()
        pw_entry = ttk.Entry(body, textvariable=self._pw_var, show="●", width=30,
                             font=("Segoe UI", 12))
        pw_entry.pack(fill="x", ipady=6)
        pw_entry.focus()

        self._login_msg = tk.Label(body, text="", bg="#f4f6fa",
                                    fg="#c0392b", font=("Segoe UI", 9))
        self._login_msg.pack(pady=4)

        def attempt_login(event=None):
            pw = self._pw_var.get()
            if hash_password(pw) == self.config["password_hash"]:
                login_win.destroy()
                self.authenticated = True
                self._build_main_ui()
            else:
                self._login_msg.config(text="❌  Incorrect password. Try again.")
                self._pw_var.set("")

        tk.Button(body, text="  Login  ", bg=accent, fg="white",
                  font=("Segoe UI", 11, "bold"), relief="flat",
                  cursor="hand2", command=attempt_login,
                  padx=20, pady=8).pack(pady=16, fill="x")

        tk.Label(body, text="By Shashikiran Rudrawar",
                 bg="#f4f6fa", fg="#999", font=("Segoe UI", 8)).pack()

        pw_entry.bind("<Return>", attempt_login)

    # ── MAIN UI ───────────────────────────────
    def _build_main_ui(self):
        accent = self.config.get("accent_color", "#1a3a6b")

        # Sidebar
        sidebar = tk.Frame(self.root, bg=accent, width=200)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        tk.Label(sidebar, text="💼", bg=accent, fg="white",
                 font=("Segoe UI", 28)).pack(pady=(24,4))
        tk.Label(sidebar, text="Salary Slip\nGenerator\n \n By S.Rudrawar",
                 bg=accent, fg="white", font=("Segoe UI", 11, "bold"),
                 justify="center").pack()

        tk.Frame(sidebar, bg="white", height=1).pack(fill="x", padx=20, pady=16)

        self._active_section = tk.StringVar(value="generate")
        self._section_btns = {}
        nav_items = [
            ("generate", "📊  Generate Slips"),
            ("company",  "🏢  Company Info"),
            ("columns",  "📋  Column Mapping"),
            ("settings", "⚙️   Settings"),
        ]
        for key, label in nav_items:
            btn = tk.Button(sidebar, text=label, bg=accent, fg="white",
                            font=("Segoe UI", 10), relief="flat", anchor="w",
                            padx=16, pady=10, cursor="hand2",
                            activebackground="#2a5298", activeforeground="white",
                            command=lambda k=key: self._switch_section(k))
            btn.pack(fill="x")
            self._section_btns[key] = btn

        # Main content
        self._main_frame = tk.Frame(self.root, bg="#f4f6fa")
        self._main_frame.pack(side="left", fill="both", expand=True)

        self._sections = {}
        self._build_generate_section()
        self._build_company_section()
        self._build_columns_section()
        self._build_settings_section()

        self._switch_section("generate")

    def _switch_section(self, key):
        accent = self.config.get("accent_color", "#1a3a6b")
        for k, btn in self._section_btns.items():
            btn.config(bg=accent if k != key else "#2a5298")
        for k, frame in self._sections.items():
            frame.pack_forget()
        self._sections[key].pack(fill="both", expand=True)
        self._active_section.set(key)

    # ── GENERATE SECTION ──────────────────────
    def _build_generate_section(self):
        frame = tk.Frame(self._main_frame, bg="#f4f6fa")
        self._sections["generate"] = frame

        # Top bar
        topbar = tk.Frame(frame, bg="#f4f6fa")
        topbar.pack(fill="x", padx=24, pady=(20,8))
        tk.Label(topbar, text="Generate Salary Slips", bg="#f4f6fa",
                 font=("Segoe UI", 18, "bold"),
                 fg=self.config.get("accent_color","#1a3a6b")).pack(side="left")

        # Excel loader card
        card = tk.Frame(frame, bg="white", relief="groove", bd=1)
        card.pack(fill="x", padx=24, pady=6)

        inner = tk.Frame(card, bg="white")
        inner.pack(fill="x", padx=16, pady=14)

        tk.Label(inner, text="Excel Salary Sheet", bg="white",
                 font=("Segoe UI", 10, "bold"), fg="#333").grid(row=0, column=0, sticky="w")

        self._excel_var = tk.StringVar(value="No file selected")
        tk.Label(inner, textvariable=self._excel_var, bg="white",
                 fg="#666", font=("Segoe UI", 9), wraplength=400).grid(row=1, column=0, sticky="w", pady=2)

        btn_frame = tk.Frame(inner, bg="white")
        btn_frame.grid(row=0, column=1, rowspan=2, sticky="e", padx=(20,0))

        accent = self.config.get("accent_color","#1a3a6b")
        tk.Button(btn_frame, text="📂  Browse Excel", bg=accent, fg="white",
                  font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2",
                  padx=14, pady=6, command=self._browse_excel).pack(side="left", padx=4)
        tk.Button(btn_frame, text="🔄  Load / Refresh", bg="#27ae60", fg="white",
                  font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2",
                  padx=14, pady=6, command=self._load_excel).pack(side="left", padx=4)

        inner.columnconfigure(0, weight=1)

        # Filter bar
        filter_frame = tk.Frame(frame, bg="#f4f6fa")
        filter_frame.pack(fill="x", padx=24, pady=4)
        tk.Label(filter_frame, text="🔍 Search:", bg="#f4f6fa",
                 font=("Segoe UI", 10)).pack(side="left")
        self._search_var = tk.StringVar()
        self._search_var.trace("w", self._filter_records)
        ttk.Entry(filter_frame, textvariable=self._search_var, width=30).pack(side="left", padx=8)

        # Table
        tree_frame = tk.Frame(frame, bg="#f4f6fa")
        tree_frame.pack(fill="both", expand=True, padx=24, pady=4)

        cols = ("sel", "name", "code", "dept", "month", "net")
        self._tree = ttk.Treeview(tree_frame, columns=cols, show="headings",
                                   selectmode="extended")
        headers = {
            "sel":   ("☑", 40),
            "name":  ("Employee Name", 200),
            "code":  ("Emp Code", 80),
            "dept":  ("Department", 140),
            "month": ("Month-Year", 110),
            "net":   ("Net Salary ₹", 110),
        }
        for c, (h, w) in headers.items():
            self._tree.heading(c, text=h)
            self._tree.column(c, width=w, anchor="center" if c in ("sel","code","net","month") else "w")

        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        self._tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self._tree.tag_configure("selected", background="#dbe9ff")
        self._tree.bind("<space>", self._toggle_selection)
        self._tree.bind("<Double-1>", self._toggle_selection)

        # Bottom action bar
        bot = tk.Frame(frame, bg="white", relief="groove", bd=1)
        bot.pack(fill="x", padx=24, pady=8)
        inner_bot = tk.Frame(bot, bg="white")
        inner_bot.pack(fill="x", padx=16, pady=10)

        self._sel_count_var = tk.StringVar(value="0 records selected")
        tk.Label(inner_bot, textvariable=self._sel_count_var, bg="white",
                 font=("Segoe UI", 9), fg="#555").pack(side="left")

        tk.Button(inner_bot, text="Select All", bg="#ecf0f1", fg="#333",
                  font=("Segoe UI", 9), relief="flat", cursor="hand2",
                  padx=10, pady=4, command=self._select_all).pack(side="left", padx=8)
        tk.Button(inner_bot, text="Clear", bg="#ecf0f1", fg="#333",
                  font=("Segoe UI", 9), relief="flat", cursor="hand2",
                  padx=10, pady=4, command=self._clear_selection).pack(side="left", padx=2)

        self._out_var = tk.StringVar(value=self.config.get("output_folder", ""))
        tk.Label(inner_bot, text="Output:", bg="white",
                 font=("Segoe UI", 9)).pack(side="left", padx=(16,4))
        tk.Label(inner_bot, textvariable=self._out_var, bg="white",
                 fg="#666", font=("Segoe UI", 8), width=30, anchor="w").pack(side="left")
        tk.Button(inner_bot, text="📁", bg="#ecf0f1", fg="#333",
                  font=("Segoe UI", 10), relief="flat", cursor="hand2",
                  command=self._choose_output).pack(side="left", padx=4)

        tk.Button(inner_bot, text="🖨️  Generate Selected PDFs", bg=accent, fg="white",
                  font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2",
                  padx=16, pady=6, command=self._generate_pdfs).pack(side="right", padx=4)
        tk.Button(inner_bot, text="⚡ Generate All", bg="#27ae60", fg="white",
                  font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2",
                  padx=16, pady=6, command=self._generate_all).pack(side="right", padx=4)

        # Progress
        self._progress_var = tk.DoubleVar()
        self._progress_bar = ttk.Progressbar(frame, variable=self._progress_var,
                                              maximum=100)
        self._status_var = tk.StringVar(value="Ready")
        tk.Label(frame, textvariable=self._status_var, bg="#f4f6fa",
                 font=("Segoe UI", 8), fg="#555").pack(padx=24, anchor="w")

        self._filtered_records = []
        self._selected_iids = set()

    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="Select Excel Salary Sheet",
            filetypes=[("Excel files", "*.xlsx *.xls *.xlsm"), ("All", "*.*")]
        )
        if path:
            self.excel_path = path
            self._excel_var.set(path)
            self._load_excel()

    def _load_excel(self):
        if not self.excel_path:
            messagebox.showwarning("No File", "Please select an Excel file first.")
            return
        try:
            self.records = read_excel_salaries(self.excel_path, self.config["column_map"])
            self._filtered_records = self.records[:]
            self._populate_tree()
            self._status_var.set(f"✅ Loaded {len(self.records)} records from Excel.")
        except Exception as e:
            messagebox.showerror("Load Error", f"Could not read Excel:\n{e}")

    def _filter_records(self, *_):
        q = self._search_var.get().lower()
        self._filtered_records = [
            r for r in self.records
            if q in r.get("employee_name","").lower()
            or q in r.get("employee_code","").lower()
            or q in r.get("department","").lower()
            or q in r.get("month_year","").lower()
        ]
        self._populate_tree()

    def _populate_tree(self):
        self._tree.delete(*self._tree.get_children())
        self._selected_iids.clear()
        for i, r in enumerate(self._filtered_records):
            net = self._calc_net(r)
            iid = self._tree.insert("", "end", iid=str(i), values=(
                "☐",
                r.get("employee_name",""),
                r.get("employee_code",""),
                r.get("department",""),
                r.get("month_year",""),
                f"₹ {net:,.2f}" if net else "",
            ))
        self._update_sel_count()

    def _calc_net(self, r):
        try:
            earn = sum(float(r.get(k,0) or 0) for k in
                       ("basic","conveyance","hra","post_allowance","other_allowance"))
            dedn = sum(float(r.get(k,0) or 0) for k in
                       ("profession_tax","pf","esic_deduction","other_deduction"))
            return earn - dedn
        except Exception:
            return 0

    def _toggle_selection(self, event=None):
        for iid in self._tree.selection():
            if iid in self._selected_iids:
                self._selected_iids.discard(iid)
                self._tree.set(iid, "sel", "☐")
                self._tree.item(iid, tags=())
            else:
                self._selected_iids.add(iid)
                self._tree.set(iid, "sel", "☑")
                self._tree.item(iid, tags=("selected",))
        self._update_sel_count()

    def _select_all(self):
        for iid in self._tree.get_children():
            self._selected_iids.add(iid)
            self._tree.set(iid, "sel", "☑")
            self._tree.item(iid, tags=("selected",))
        self._update_sel_count()

    def _clear_selection(self):
        for iid in self._selected_iids:
            self._tree.set(iid, "sel", "☐")
            self._tree.item(iid, tags=())
        self._selected_iids.clear()
        self._update_sel_count()

    def _update_sel_count(self):
        self._sel_count_var.set(f"{len(self._selected_iids)} of "
                                 f"{len(self._filtered_records)} records selected")

    def _choose_output(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.config["output_folder"] = folder
            self._out_var.set(folder)
            save_config(self.config)

    def _generate_pdfs(self):
        selected = [self._filtered_records[int(iid)] for iid in self._selected_iids]
        if not selected:
            messagebox.showinfo("Nothing Selected",
                                "Please select at least one record (double-click or Space).")
            return
        self._run_generation(selected)

    def _generate_all(self):
        if not self._filtered_records:
            messagebox.showinfo("No Records", "Load an Excel file first.")
            return
        self._run_generation(self._filtered_records)

    def _run_generation(self, records):
        out_folder = self.config.get("output_folder", str(Path.home() / "SalarySlips"))
        Path(out_folder).mkdir(parents=True, exist_ok=True)

        self._progress_bar.pack(fill="x", padx=24, pady=2)
        self._progress_var.set(0)
        self._status_var.set("Generating PDFs…")

        def worker():
            total = len(records)
            success, fail = 0, 0
            for idx, rec in enumerate(records):
                try:
                    name = rec.get("employee_name", f"emp_{idx}").replace(" ", "_")
                    code = rec.get("employee_code", "")
                    month = rec.get("month_year", "").replace(" ","_").replace("/","-")
                    fname = f"{code}_{name}_{month}.pdf".replace(" ","_")
                    out_path = os.path.join(out_folder, fname)
                    generate_salary_slip_pdf(rec, self.config, out_path)
                    success += 1
                except Exception as e:
                    fail += 1
                    print(f"Error generating {rec.get('employee_name')}: {e}")
                pct = (idx+1)/total*100
                self.root.after(0, self._progress_var.set, pct)
                self.root.after(0, self._status_var.set,
                                f"Processing {idx+1}/{total}…")

            def done():
                self._progress_bar.pack_forget()
                msg = f"✅ Generated {success} PDF(s)"
                if fail:
                    msg += f"  ⚠️ {fail} failed"
                self._status_var.set(msg)
                if messagebox.askyesno("Done!", f"{msg}\n\nOpen output folder?"):
                    import subprocess, platform
                    if platform.system() == "Windows":
                        os.startfile(out_folder)
                    elif platform.system() == "Darwin":
                        subprocess.Popen(["open", out_folder])
                    else:
                        subprocess.Popen(["xdg-open", out_folder])

            self.root.after(0, done)

        threading.Thread(target=worker, daemon=True).start()

    # ── COMPANY SECTION ───────────────────────
    def _build_company_section(self):
        frame = tk.Frame(self._main_frame, bg="#f4f6fa")
        self._sections["company"] = frame

        tk.Label(frame, text="Company Information", bg="#f4f6fa",
                 font=("Segoe UI", 18, "bold"),
                 fg=self.config.get("accent_color","#1a3a6b")).pack(anchor="w", padx=24, pady=(20,8))

        card = tk.Frame(frame, bg="white", relief="groove", bd=1)
        card.pack(fill="both", expand=True, padx=24, pady=8)

        inner = tk.Frame(card, bg="white")
        inner.pack(fill="both", expand=True, padx=24, pady=20)

        def row_lbl(text):
            tk.Label(inner, text=text, bg="white", font=("Segoe UI", 10, "bold"),
                     fg="#333", anchor="w").pack(fill="x", pady=(10,2))

        # Company Name
        row_lbl("Company Name")
        self._co_name = tk.StringVar(value=self.config.get("company_name",""))
        ttk.Entry(inner, textvariable=self._co_name, font=("Segoe UI", 11),
                  width=60).pack(fill="x", ipady=4)

        # Address
        row_lbl("Company Address")
        self._co_addr = tk.Text(inner, height=3, font=("Segoe UI", 10),
                                 wrap="word", relief="groove", bd=1)
        self._co_addr.insert("1.0", self.config.get("company_address",""))
        self._co_addr.pack(fill="x")

        # Logo
        row_lbl("Company Logo (PNG/JPG)")
        logo_frame = tk.Frame(inner, bg="white")
        logo_frame.pack(fill="x", pady=4)

        self._logo_var = tk.StringVar(value=self.config.get("company_logo",""))
        tk.Label(logo_frame, textvariable=self._logo_var, bg="white",
                 fg="#666", font=("Segoe UI", 9), width=50, anchor="w").pack(side="left")

        self._logo_preview = tk.Label(logo_frame, bg="white")
        self._logo_preview.pack(side="left", padx=10)

        def pick_logo():
            path = filedialog.askopenfilename(
                title="Select Logo Image",
                filetypes=[("Images", "*.png *.jpg *.jpeg *.gif *.bmp"),("All","*.*")]
            )
            if path:
                self._logo_var.set(path)
                self._update_logo_preview(path)

        def clear_logo():
            self._logo_var.set("")
            self._logo_preview.config(image="", text="")

        tk.Button(logo_frame, text="📁 Browse", bg="#ecf0f1", fg="#333",
                  font=("Segoe UI", 9), relief="flat", cursor="hand2",
                  padx=10, pady=4, command=pick_logo).pack(side="left", padx=4)
        tk.Button(logo_frame, text="✕ Clear", bg="#ecf0f1", fg="#c0392b",
                  font=("Segoe UI", 9), relief="flat", cursor="hand2",
                  padx=10, pady=4, command=clear_logo).pack(side="left")

        if self.config.get("company_logo"):
            self._update_logo_preview(self.config["company_logo"])

        # Accent color
        row_lbl("Accent Color (PDF Header & Borders)")
        col_frame = tk.Frame(inner, bg="white")
        col_frame.pack(fill="x", pady=4)

        self._color_preview = tk.Label(col_frame, text="       ",
                                        bg=self.config.get("accent_color","#1a3a6b"),
                                        relief="groove", bd=2)
        self._color_preview.pack(side="left", padx=(0,8))
        self._accent_var = tk.StringVar(value=self.config.get("accent_color","#1a3a6b"))
        tk.Label(col_frame, textvariable=self._accent_var, bg="white",
                 font=("Segoe UI", 10)).pack(side="left")

        def pick_color():
            c = colorchooser.askcolor(color=self._accent_var.get(),
                                       title="Choose Accent Color")
            if c and c[1]:
                self._accent_var.set(c[1])
                self._color_preview.config(bg=c[1])

        tk.Button(col_frame, text="🎨 Pick Color", bg="#ecf0f1", fg="#333",
                  font=("Segoe UI", 9), relief="flat", cursor="hand2",
                  padx=10, pady=4, command=pick_color).pack(side="left", padx=8)

        # Save button
        accent = self.config.get("accent_color","#1a3a6b")
        tk.Button(inner, text="💾  Save Company Info", bg=accent, fg="white",
                  font=("Segoe UI", 11, "bold"), relief="flat", cursor="hand2",
                  padx=20, pady=8, command=self._save_company).pack(pady=20)

    def _update_logo_preview(self, path):
        try:
            img = Image.open(path)
            img.thumbnail((60, 60))
            photo = ImageTk.PhotoImage(img)
            self._logo_preview.config(image=photo, text="")
            self._logo_preview._photo = photo
        except Exception:
            pass

    def _save_company(self):
        self.config["company_name"] = self._co_name.get().strip()
        self.config["company_address"] = self._co_addr.get("1.0","end").strip()
        self.config["company_logo"] = self._logo_var.get().strip()
        self.config["accent_color"] = self._accent_var.get().strip()
        save_config(self.config)
        messagebox.showinfo("Saved", "Company information saved successfully!")
        self._apply_theme()

    # ── COLUMN MAPPING SECTION ────────────────
    def _build_columns_section(self):
        frame = tk.Frame(self._main_frame, bg="#f4f6fa")
        self._sections["columns"] = frame

        tk.Label(frame, text="Column Mapping", bg="#f4f6fa",
                 font=("Segoe UI", 18, "bold"),
                 fg=self.config.get("accent_color","#1a3a6b")).pack(anchor="w", padx=24, pady=(20,4))
        tk.Label(frame, text="Map the internal salary fields to your Excel column headers.",
                 bg="#f4f6fa", font=("Segoe UI", 9), fg="#666").pack(anchor="w", padx=24, pady=(0,8))

        card = tk.Frame(frame, bg="white", relief="groove", bd=1)
        card.pack(fill="both", expand=True, padx=24, pady=8)

        canvas = tk.Canvas(card, bg="white", highlightthickness=0)
        vsb = ttk.Scrollbar(card, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        inner = tk.Frame(canvas, bg="white")
        canvas.create_window((0,0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")))

        col_map = self.config.get("column_map", DEFAULT_CONFIG["column_map"])
        self._col_vars = {}

        labels_map = {
            "employee_name": "Employee Name",
            "employee_code": "Employee Code",
            "department": "Department",
            "designation": "Designation",
            "location": "Location",
            "pf_no": "PF Number",
            "uan_no": "UAN Number",
            "esic_no": "ESIC Number",
            "pan_no": "PAN Number",
            "aadhar_no": "Aadhar Number",
            "bank_name": "Bank Name",
            "bank_ac": "Bank Account No",
            "monthly_salary": "Monthly Salary (CTC)",
            "month": "Month (e.g. MARCH)",
            "year": "Year (e.g. 2022)",
            "basic": "Basic Salary",
            "conveyance": "Conveyance",
            "hra": "House Rent Allowance",
            "post_allowance": "Post Allowance",
            "other_allowance": "Other Allowance",
            "profession_tax": "Profession Tax",
            "pf": "Provident Fund",
            "esic_deduction": "ESIC Deduction",
            "other_deduction": "Other Deduction",
            "month_days": "Month Days",
            "weekly_off": "Weekly Off Days",
            "holidays": "Holidays",
            "present_days": "Present Days",
            "pay_days": "Pay Days",
            "leave_without_pay": "Leave Without Pay",
        }

        # Group
        groups = [
            ("👤 Employee Details", ["employee_name","employee_code","department","designation","location"]),
            ("🏦 IDs & Bank", ["pf_no","uan_no","esic_no","pan_no","aadhar_no","bank_name","bank_ac","monthly_salary","month","year"]),
            ("💰 Earnings", ["basic","conveyance","hra","post_allowance","other_allowance"]),
            ("➖ Deductions", ["profession_tax","pf","esic_deduction","other_deduction"]),
            ("📅 Attendance", ["month_days","weekly_off","holidays","present_days","pay_days","leave_without_pay"]),
        ]

        for grp_title, keys in groups:
            tk.Label(inner, text=grp_title, bg="white",
                     font=("Segoe UI", 10, "bold"), fg="#333",
                     pady=6).grid(row=len(self._col_vars)*2, column=0,
                                  columnspan=4, sticky="w", padx=16, pady=(12,4))
            for key in keys:
                r = len(self._col_vars)*2 + 1
                lbl = labels_map.get(key, key)
                tk.Label(inner, text=lbl, bg="white",
                         font=("Segoe UI", 9), fg="#555", anchor="w",
                         width=26).grid(row=r, column=0, padx=(24,8), pady=2, sticky="w")
                tk.Label(inner, text="→", bg="white", fg="#aaa").grid(row=r, column=1, padx=4)
                var = tk.StringVar(value=col_map.get(key, ""))
                self._col_vars[key] = var
                ttk.Entry(inner, textvariable=var, width=28,
                          font=("Segoe UI", 9)).grid(row=r, column=2, padx=4, pady=2, sticky="w")
                tk.Label(inner, text="← Excel column header name", bg="white",
                         fg="#aaa", font=("Segoe UI", 8)).grid(row=r, column=3, padx=8, sticky="w")

        accent = self.config.get("accent_color","#1a3a6b")
        btn_frame = tk.Frame(card, bg="white")
        btn_frame.pack(fill="x", padx=16, pady=12)
        tk.Button(btn_frame, text="💾  Save Column Mapping", bg=accent, fg="white",
                  font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2",
                  padx=16, pady=6, command=self._save_columns).pack(side="left")
        tk.Button(btn_frame, text="↺ Reset to Defaults", bg="#ecf0f1", fg="#333",
                  font=("Segoe UI", 10), relief="flat", cursor="hand2",
                  padx=16, pady=6, command=self._reset_columns).pack(side="left", padx=10)

    def _save_columns(self):
        new_map = {k: v.get().strip() for k, v in self._col_vars.items()}
        self.config["column_map"] = new_map
        save_config(self.config)
        messagebox.showinfo("Saved", "Column mapping saved!")

    def _reset_columns(self):
        if messagebox.askyesno("Reset", "Reset all column mappings to defaults?"):
            for k, v in self._col_vars.items():
                v.set(DEFAULT_CONFIG["column_map"].get(k, ""))

    # ── SETTINGS SECTION ──────────────────────
    def _build_settings_section(self):
        frame = tk.Frame(self._main_frame, bg="#f4f6fa")
        self._sections["settings"] = frame

        tk.Label(frame, text="Settings", bg="#f4f6fa",
                 font=("Segoe UI", 18, "bold"),
                 fg=self.config.get("accent_color","#1a3a6b")).pack(anchor="w", padx=24, pady=(20,8))

        card = tk.Frame(frame, bg="white", relief="groove", bd=1)
        card.pack(fill="x", padx=24, pady=8)
        inner = tk.Frame(card, bg="white")
        inner.pack(fill="x", padx=24, pady=20)

        # Change password
        tk.Label(inner, text="🔐  Change Password", bg="white",
                 font=("Segoe UI", 12, "bold"), fg="#333").pack(anchor="w", pady=(0,10))

        def lbl_entry(parent, text, show=None):
            tk.Label(parent, text=text, bg="white",
                     font=("Segoe UI", 10), fg="#555", anchor="w").pack(fill="x", pady=(6,2))
            var = tk.StringVar()
            e = ttk.Entry(parent, textvariable=var, font=("Segoe UI", 10),
                          show=show or "")
            e.pack(fill="x", ipady=4)
            return var

        pw_frame = tk.Frame(inner, bg="white")
        pw_frame.pack(fill="x")

        self._old_pw = lbl_entry(pw_frame, "Current Password", show="●")
        self._new_pw = lbl_entry(pw_frame, "New Password", show="●")
        self._cnf_pw = lbl_entry(pw_frame, "Confirm New Password", show="●")

        accent = self.config.get("accent_color","#1a3a6b")

        def change_password():
            old = self._old_pw.get()
            new = self._new_pw.get()
            cnf = self._cnf_pw.get()
            if hash_password(old) != self.config["password_hash"]:
                messagebox.showerror("Error", "Current password is incorrect.")
                return
            if len(new) < 6:
                messagebox.showerror("Error", "New password must be at least 6 characters.")
                return
            if new != cnf:
                messagebox.showerror("Error", "Passwords do not match.")
                return
            self.config["password_hash"] = hash_password(new)
            save_config(self.config)
            self._old_pw.set("")
            self._new_pw.set("")
            self._cnf_pw.set("")
            messagebox.showinfo("Success", "Password changed successfully!")

        tk.Button(pw_frame, text="🔒  Change Password", bg=accent, fg="white",
                  font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2",
                  padx=16, pady=6, command=change_password).pack(pady=12, anchor="w")

        tk.Frame(inner, bg="#eee", height=1).pack(fill="x", pady=12)

        # Output folder
        tk.Label(inner, text="📁  Default Output Folder", bg="white",
                 font=("Segoe UI", 12, "bold"), fg="#333").pack(anchor="w", pady=(0,8))

        of_frame = tk.Frame(inner, bg="white")
        of_frame.pack(fill="x")
        self._out_settings_var = tk.StringVar(value=self.config.get("output_folder",""))
        tk.Label(of_frame, textvariable=self._out_settings_var, bg="white",
                 fg="#666", font=("Segoe UI", 9), anchor="w").pack(side="left", fill="x", expand=True)

        def pick_out():
            folder = filedialog.askdirectory(title="Select Default Output Folder")
            if folder:
                self.config["output_folder"] = folder
                self._out_settings_var.set(folder)
                self._out_var.set(folder)
                save_config(self.config)

        tk.Button(of_frame, text="📁 Browse", bg="#ecf0f1", fg="#333",
                  font=("Segoe UI", 9), relief="flat", cursor="hand2",
                  padx=10, pady=4, command=pick_out).pack(side="right")

        tk.Frame(inner, bg="#eee", height=1).pack(fill="x", pady=12)

        # About
        tk.Label(inner, text="ℹ️  About", bg="white",
                 font=("Segoe UI", 12, "bold"), fg="#333").pack(anchor="w", pady=(0,6))
        tk.Label(inner, text="Salary Slip Generator v1.0 By Shashikiran Rudrawar\n",
                 bg="white", font=("Segoe UI", 11), fg="#666",
                 justify="left").pack(anchor="w")


# ────────────────────────────────────────────────
#  ENTRY POINT
# ────────────────────────────────────────────────
if __name__ == "__main__":
    SalarySlipApp()
