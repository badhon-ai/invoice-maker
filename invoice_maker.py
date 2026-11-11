"""
Invoice Maker - simple desktop GUI (Tkinter) that reproduces the S&P invoice layout.

Install once:   pip install python-docx reportlab
Run:            python invoice_maker.py

- Each work entry has Date, Start, Break Start, Break End, Finish, Position, Depot, Hourly Rate, Total
- Hours and pay are calculated automatically from the shift and break times; overnight shifts work too
- Your own details (name, address, ABN, bank) are saved in invoice_settings.json
- Invoice number auto-increments for each new invoice
- Export to Word (.docx) and/or PDF
"""

import json
import os
import tkinter as tk
from datetime import datetime, timedelta
from tkinter import ttk, messagebox, filedialog

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    HAVE_PDF = True
except ImportError:
    HAVE_PDF = False

SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "invoice_settings.json")

DEFAULTS = {
    "name": "",
    "address": "",
    "abn": "",
    "bank": "",
    "bsb": "",
    "acc": "",
    "bill_to": "",
    "position": "",
    "depot": "",
    "start": "6:00 AM",
    "finish": "2:00 PM",
    "break_min": 30,
    "break_start": "10:00 AM",
    "break_end": "10:30 AM",
    "hourly_rate": "30.00",
    "no_break": False,
    "next_invoice": 1,
    "gst_mode": "Included",
}

DATE_FMT = "%d/%m/%Y"
TIME_FORMATS = ("%H:%M", "%I:%M %p", "%I:%M%p", "%I %p", "%I%p", "%H%M")


# ----------------------------------------------------------------- helpers
def load_settings():
    data = dict(DEFAULTS)
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                data.update(json.load(f))
        except (OSError, json.JSONDecodeError):
            pass
    return data


def save_settings(data):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)


def money(x):
    return f"${x:,.2f}"


def parse_date(s):
    s = s.strip()
    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def parse_time(s):
    """Accepts 6:00 AM, 6am, 06:00, 18:00, 1430 ... Returns datetime or None."""
    s = s.strip().upper().replace(".", "")
    if not s:
        return None
    for fmt in TIME_FORMATS:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def fmt_time(dt):
    return dt.strftime("%I:%M %p").lstrip("0")


def break_duration(start, end):
    """Return break duration in minutes; supports breaks crossing midnight."""
    if not start.strip() and not end.strip():
        return 0
    s, e = parse_time(start), parse_time(end)
    if not s or not e:
        return None
    mins = (e - s).total_seconds() / 60
    if mins < 0:
        mins += 24 * 60
    return int(mins)


def break_text(start, end):
    minutes = break_duration(start, end)
    if minutes is None:
        return "Invalid break"
    return "No break" if minutes == 0 else f"{minutes} min"


def hours_worked(start, finish, break_min):
    """Hours between start and finish minus break; handles overnight shifts."""
    s, f = parse_time(start), parse_time(finish)
    if not s or not f or break_min is None:
        return None
    mins = (f - s).total_seconds() / 60
    if mins <= 0:
        mins += 24 * 60
    mins -= break_min
    return mins / 60 if mins >= 0 else None


def fmt_hours(h):
    return "-" if h is None else f"{h:.2f} h"


def compute_totals(rows, gst_mode):
    """Rows contain date, shift/break times, position, depot, hourly rate and calculated amount."""
    subtotal = sum(r[8] for r in rows)
    if gst_mode == "Added (+10%)":
        gst = round(subtotal * 0.10, 2)
        return subtotal, gst, subtotal + gst, money(gst)
    if gst_mode == "None":
        return subtotal, 0.0, subtotal, "N/A"
    return subtotal, round(subtotal / 11, 2), subtotal, "Included"


HEADERS = ["From-To", "Position", "Depot Name:", "Total"]


def row_cells(r):
    d, start, break_start, break_end, finish, pos, dep, rate, amt = r
    return [d, pos, dep, money(amt)]


# ------------------------------------------------------------- Word export
def _set_cell(cell, text, bold=False, align=None, size=10):
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    if align:
        p.alignment = align


def export_docx(path, inv):
    doc = Document()
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Cm(1.7)
    sec.top_margin = Cm(1.6)
    sec.bottom_margin = Cm(1.6)
    doc.styles["Normal"].font.name = "Aptos"
    doc.styles["Normal"].font.size = Pt(10)

    # Small reusable helpers for a clean, modern invoice layout.
    def shade(cell, fill):
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:fill"), fill)
        tcPr.append(shd)

    def set_cell_margins(cell, top=100, start=120, bottom=100, end=120):
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcMar = tcPr.first_child_found_in("w:tcMar")
        if tcMar is None:
            tcMar = OxmlElement("w:tcMar")
            tcPr.append(tcMar)
        for m, v in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
            node = tcMar.find(qn("w:" + m))
            if node is None:
                node = OxmlElement("w:" + m)
                tcMar.append(node)
            node.set(qn("w:w"), str(v))
            node.set(qn("w:type"), "dxa")

    # Invoice title and number
    ptitle = doc.add_paragraph()
    ptitle.paragraph_format.space_after = Pt(2)
    r = ptitle.add_run("INVOICE")
    r.bold = True
    r.font.name = "Aptos Display"
    r.font.size = Pt(26)
    r.font.color.rgb = __import__("docx").shared.RGBColor(31, 55, 85)
    pnum = doc.add_paragraph()
    pnum.paragraph_format.space_after = Pt(14)
    rn = pnum.add_run(f"Invoice No. {inv['number']}")
    rn.bold = True
    rn.font.size = Pt(11)

    # Bill-to and invoice period information
    info = doc.add_table(rows=1, cols=2)
    info.autofit = False
    info.columns[0].width = Cm(8.2)
    info.columns[1].width = Cm(8.2)
    for cell in info.rows[0].cells:
        set_cell_margins(cell, 160, 180, 160, 180)
        shade(cell, "F1F5F9")
    left, right = info.rows[0].cells
    left.text = ""
    pp = left.paragraphs[0]
    rr = pp.add_run("BILL TO\n")
    rr.bold = True
    rr.font.size = Pt(9)
    rr.font.color.rgb = __import__("docx").shared.RGBColor(71, 85, 105)
    rr2 = pp.add_run(inv["bill_to"])
    rr2.bold = True
    rr2.font.size = Pt(12)
    right.text = ""
    pp = right.paragraphs[0]
    rr = pp.add_run("PERIOD\n")
    rr.bold = True
    rr.font.size = Pt(9)
    rr.font.color.rgb = __import__("docx").shared.RGBColor(71, 85, 105)
    rr2 = pp.add_run(f"Week Ending {inv['week_ending']}")
    rr2.bold = True
    rr2.font.size = Pt(12)
    doc.add_paragraph().paragraph_format.space_after = Pt(3)

    # Work entries: only the four requested columns.
    rows = inv["rows"]
    table = doc.add_table(rows=1, cols=len(HEADERS))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    widths = [Cm(3.2), Cm(4.4), Cm(4.4), Cm(3.8)]
    for i, width in enumerate(widths):
        table.columns[i].width = width
    for i, h in enumerate(HEADERS):
        cell = table.rows[0].cells[i]
        _set_cell(cell, h.upper(), bold=True, size=9)
        shade(cell, "1F3755")
        set_cell_margins(cell, 140, 120, 140, 120)
        for run in cell.paragraphs[0].runs:
            run.font.color.rgb = __import__("docx").shared.RGBColor(255, 255, 255)
    for idx, row in enumerate(rows):
        cells = table.add_row().cells
        for i, val in enumerate(row_cells(row)):
            _set_cell(cells[i], val, bold=(i == 3), align=WD_ALIGN_PARAGRAPH.RIGHT if i == 3 else None, size=10)
            set_cell_margins(cells[i])
            if idx % 2 == 1:
                shade(cells[i], "F8FAFC")
    for _ in range(max(0, 3 - len(rows))):
        cells = table.add_row().cells
        for cell in cells:
            set_cell_margins(cell, 180, 120, 180, 120)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)

    _, _, total, gst_label = compute_totals(rows, inv["gst_mode"])
    summary = doc.add_table(rows=2, cols=2)
    summary.alignment = WD_TABLE_ALIGNMENT.RIGHT
    summary.autofit = False
    summary.columns[0].width = Cm(4.5)
    summary.columns[1].width = Cm(4.0)
    for row in summary.rows:
        for cell in row.cells:
            set_cell_margins(cell, 110, 150, 110, 150)
    _set_cell(summary.rows[0].cells[0], "GST", bold=True, size=10)
    _set_cell(summary.rows[0].cells[1], gst_label, align=WD_ALIGN_PARAGRAPH.RIGHT, size=10)
    _set_cell(summary.rows[1].cells[0], "TOTAL", bold=True, size=11)
    _set_cell(summary.rows[1].cells[1], money(total), bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT, size=12)
    for cell in summary.rows[1].cells:
        shade(cell, "E8EEF5")

    pthank = doc.add_paragraph()
    pthank.paragraph_format.space_before = Pt(18)
    pthank.paragraph_format.space_after = Pt(12)
    rt = pthank.add_run("Thank you for your business!")
    rt.bold = True
    rt.font.size = Pt(11)
    rt.font.color.rgb = __import__("docx").shared.RGBColor(31, 55, 85)

    # Payment and supplier details in a neat footer section.
    footer = doc.add_table(rows=1, cols=2)
    footer.autofit = False
    footer.columns[0].width = Cm(8.2)
    footer.columns[1].width = Cm(8.2)
    for cell in footer.rows[0].cells:
        shade(cell, "F1F5F9")
        set_cell_margins(cell, 150, 180, 150, 180)
    c1, c2 = footer.rows[0].cells
    c1.text = ""
    p = c1.paragraphs[0]
    r = p.add_run("FROM\n")
    r.bold = True
    r.font.size = Pt(9)
    r.font.color.rgb = __import__("docx").shared.RGBColor(71, 85, 105)
    p.add_run(f"{inv['name']}\n{inv['address']}\nABN: {inv['abn']}")
    c2.text = ""
    p = c2.paragraphs[0]
    r = p.add_run("PAYMENT DETAILS\n")
    r.bold = True
    r.font.size = Pt(9)
    r.font.color.rgb = __import__("docx").shared.RGBColor(71, 85, 105)
    p.add_run(f"{inv['bank']}\nBSB: {inv['bsb']}\nAccount: {inv['acc']}")
    doc.save(path)


# -------------------------------------------------------------- PDF export
def export_pdf(path, inv):
    styles = getSampleStyleSheet()
    navy = colors.HexColor("#1F3755")
    pale = colors.HexColor("#F1F5F9")
    stripe = colors.HexColor("#F8FAFC")
    muted = colors.HexColor("#475569")
    normal = ParagraphStyle("InvoiceNormal", parent=styles["Normal"], fontName="Helvetica",
                            fontSize=9.5, leading=13, textColor=colors.HexColor("#1F2937"))
    small_label = ParagraphStyle("InvoiceLabel", parent=normal, fontName="Helvetica-Bold",
                                 fontSize=8, leading=10, textColor=muted)
    title = ParagraphStyle("InvoiceTitle", parent=normal, fontName="Helvetica-Bold",
                           fontSize=26, leading=30, textColor=navy)
    esc = lambda s: str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    doc = SimpleDocTemplate(path, pagesize=A4, leftMargin=1.7 * cm, rightMargin=1.7 * cm,
                            topMargin=1.6 * cm, bottomMargin=1.6 * cm)
    story = [
        Paragraph("INVOICE", title),
        Spacer(1, 3),
        Paragraph(f"<b>Invoice No. {esc(inv['number'])}</b>", normal),
        Spacer(1, 14),
    ]

    info = Table([
        [Paragraph("<b>BILL TO</b><br/>" + esc(inv["bill_to"]), normal),
         Paragraph("<b>PERIOD</b><br/>" + esc(f"Week Ending {inv['week_ending']}"), normal)]
    ], colWidths=[8.2 * cm, 8.2 * cm])
    info.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), pale),
        ("BOX", (0, 0), (-1, -1), 0.4, colors.HexColor("#D8E0E8")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
    ]))
    story += [info, Spacer(1, 16)]

    rows = inv["rows"]
    header_style = ParagraphStyle(
        "InvoiceTableHeader", parent=normal, fontName="Helvetica-Bold",
        fontSize=9.5, leading=12, textColor=colors.white
    )
    data = [[Paragraph(esc(h.upper()), header_style) for h in HEADERS]]
    data += [[Paragraph(esc(v), normal) for v in row_cells(r)] for r in rows]
    data += [[""] * len(HEADERS)] * max(0, 3 - len(rows))
    t = Table(data, colWidths=[3.2 * cm, 4.4 * cm, 4.4 * cm, 3.8 * cm], repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), navy),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D8E0E8")),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (-1, 0), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
    ]
    for row_idx in range(2, len(data), 2):
        style.append(("BACKGROUND", (0, row_idx), (-1, row_idx), stripe))
    t.setStyle(TableStyle(style))
    story += [t, Spacer(1, 10)]

    _, _, total, gst_label = compute_totals(rows, inv["gst_mode"])
    summary = Table([
        [Paragraph("<b>GST</b>", normal), Paragraph(esc(gst_label), normal)],
        [Paragraph("<b>TOTAL</b>", normal), Paragraph(f"<b>{money(total)}</b>", normal)]
    ], colWidths=[4.5 * cm, 4.0 * cm], hAlign="RIGHT")
    summary.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D8E0E8")),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#E8EEF5")),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    story += [summary, Spacer(1, 18),
              Paragraph("<b><font color='#1F3755'>Thank you for your business!</font></b>", normal),
              Spacer(1, 12)]

    footer = Table([
        [Paragraph("<b>FROM</b><br/>" + esc(inv["name"]) + "<br/>" + esc(inv["address"]) +
                   "<br/>ABN: " + esc(inv["abn"]), normal),
         Paragraph("<b>PAYMENT DETAILS</b><br/>" + esc(inv["bank"]) +
                   "<br/>BSB: " + esc(inv["bsb"]) + "<br/>Account: " + esc(inv["acc"]), normal)]
    ], colWidths=[8.2 * cm, 8.2 * cm])
    footer.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), pale),
        ("BOX", (0, 0), (-1, -1), 0.4, colors.HexColor("#D8E0E8")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
    ]))
    story.append(footer)
    doc.build(story)


# --------------------------------------------------------------------- GUI
class InvoiceApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Invoice Maker")
        self.geometry("940x780")
        self.minsize(860, 700)
        self.settings = load_settings()
        self.rows = []  # (date, start, break_start, break_end, finish, position, depot, rate, amount)

        self.v = {k: tk.StringVar(value=str(self.settings[k])) for k in
                  ("name", "address", "abn", "bank", "bsb", "acc", "bill_to")}
        self.v_number = tk.StringVar(value=f"{int(self.settings['next_invoice']):03d}")
        self.v_week = tk.StringVar()
        self.v_gst = tk.StringVar(value=self.settings.get("gst_mode", "Included"))

        self.v_date = tk.StringVar(value=datetime.now().strftime(DATE_FMT))
        self.v_start = tk.StringVar(value=self.settings["start"])
        self.v_finish = tk.StringVar(value=self.settings["finish"])
        self.v_break_start = tk.StringVar(value=self.settings.get("break_start", "10:00 AM"))
        self.v_break_end = tk.StringVar(value=self.settings.get("break_end", "10:30 AM"))
        self.v_nobreak = tk.BooleanVar(value=bool(self.settings["no_break"]))
        self.v_rate = tk.StringVar(value=str(self.settings.get("hourly_rate", "30.00")))
        self.v_pos = tk.StringVar(value=self.settings["position"])
        self.v_depot = tk.StringVar(value=self.settings["depot"])
        self.v_amount = tk.StringVar()
        self.v_hours = tk.StringVar()
        self.v_total = tk.StringVar()

        self._build()
        for var in (self.v_start, self.v_finish, self.v_break_start, self.v_break_end, self.v_nobreak, self.v_rate):
            var.trace_add("write", lambda *_: self._update_hours())
        self._toggle_break()
        self._refresh()

    # ---- layout
    def _build(self):
        pad = {"padx": 6, "pady": 4}
        main = ttk.Frame(self, padding=10)
        main.pack(fill="both", expand=True)

        # Invoice details
        f1 = ttk.LabelFrame(main, text="Invoice details")
        f1.pack(fill="x", **pad)
        ttk.Label(f1, text="Invoice No.").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(f1, textvariable=self.v_number, width=10).grid(row=0, column=1, sticky="w", **pad)
        ttk.Label(f1, text="Bill To").grid(row=0, column=2, sticky="w", **pad)
        ttk.Entry(f1, textvariable=self.v["bill_to"], width=34).grid(row=0, column=3, sticky="we", **pad)
        ttk.Label(f1, text="Week Ending").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(f1, textvariable=self.v_week, width=14).grid(row=1, column=1, sticky="w", **pad)
        ttk.Button(f1, text="Auto (next Sunday)", command=self._auto_week).grid(row=1, column=2, sticky="w", **pad)
        gst_box = ttk.Frame(f1)
        gst_box.grid(row=1, column=3, sticky="e", **pad)
        ttk.Label(gst_box, text="GST ").pack(side="left")
        ttk.Combobox(gst_box, textvariable=self.v_gst, state="readonly", width=16,
                     values=["Included", "Added (+10%)", "None"]).pack(side="left")
        self.v_gst.trace_add("write", lambda *_: self._refresh())
        f1.columnconfigure(3, weight=1)

        # Entry area
        f2 = ttk.LabelFrame(main, text="Add / edit work entry")
        f2.pack(fill="x", **pad)

        def field(parent, label, var, width, row, col):
            ttk.Label(parent, text=label).grid(row=row, column=col, sticky="w", padx=6)
            e = ttk.Entry(parent, textvariable=var, width=width)
            e.grid(row=row + 1, column=col, sticky="w", padx=6, pady=(0, 6))
            return e

        field(f2, "Date (dd/mm/yyyy)", self.v_date, 16, 0, 0)
        field(f2, "Start (e.g. 6:00 AM)", self.v_start, 14, 0, 1)
        field(f2, "Finish (e.g. 2:00 PM)", self.v_finish, 14, 0, 2)
        self.break_start_entry = field(f2, "Break Start (e.g. 10:00 AM)", self.v_break_start, 16, 0, 3)
        self.break_end_entry = field(f2, "Break End (e.g. 10:30 AM)", self.v_break_end, 16, 0, 4)
        ttk.Checkbutton(f2, text="No break", variable=self.v_nobreak,
                        command=self._toggle_break).grid(row=1, column=5, sticky="w", padx=6, pady=(0, 6))

        field(f2, "Position", self.v_pos, 16, 2, 0)
        field(f2, "Depot", self.v_depot, 14, 2, 1)
        field(f2, "Hourly Rate ($)", self.v_rate, 14, 2, 2)
        ttk.Label(f2, text="Hours worked").grid(row=2, column=3, sticky="w", padx=6)
        ttk.Label(f2, textvariable=self.v_hours, font=("Segoe UI", 11, "bold")).grid(
            row=3, column=3, sticky="w", padx=6, pady=(0, 6))
        ttk.Label(f2, text="Calculated Pay ($)").grid(row=2, column=4, sticky="w", padx=6)
        ttk.Label(f2, textvariable=self.v_amount, font=("Segoe UI", 11, "bold")).grid(
            row=3, column=4, sticky="w", padx=6, pady=(0, 6))

        btns = ttk.Frame(f2)
        btns.grid(row=4, column=0, columnspan=5, sticky="w", padx=4, pady=(0, 6))
        ttk.Button(btns, text="Add", command=self.add_row).pack(side="left", padx=2)
        ttk.Button(btns, text="Update selected", command=self.update_row).pack(side="left", padx=2)
        ttk.Button(btns, text="Delete selected", command=self.delete_row).pack(side="left", padx=2)

        # Table
        f3 = ttk.Frame(main)
        f3.pack(fill="both", expand=True, **pad)
        cols = ("date", "start", "break", "finish", "hours", "position", "depot", "total")
        self.tree = ttk.Treeview(f3, columns=cols, show="headings", height=8, selectmode="browse")
        heads = ("From-To", "Start", "Break", "Finish", "Hours", "Position", "Depot Name", "Total")
        widths = (90, 80, 80, 80, 70, 120, 120, 90)
        for c, t, w in zip(cols, heads, widths):
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w, anchor="e" if c == "total" else "w")
        sb = ttk.Scrollbar(f3, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        ttk.Label(main, textvariable=self.v_total, font=("Segoe UI", 12, "bold")).pack(anchor="e", padx=10)

        # My details
        f4 = ttk.LabelFrame(main, text="Your details (saved automatically)")
        f4.pack(fill="x", **pad)
        fields = [("Name", "name"), ("Address", "address"), ("ABN", "abn"),
                  ("Bank", "bank"), ("BSB", "bsb"), ("Account", "acc")]
        for i, (lbl, key) in enumerate(fields):
            r, c = divmod(i, 2)
            ttk.Label(f4, text=lbl).grid(row=r, column=c * 2, sticky="w", **pad)
            ttk.Entry(f4, textvariable=self.v[key], width=34).grid(row=r, column=c * 2 + 1, sticky="we", **pad)
        f4.columnconfigure(1, weight=1)
        f4.columnconfigure(3, weight=1)

        # Actions
        f5 = ttk.Frame(main)
        f5.pack(fill="x", **pad)
        ttk.Button(f5, text="Export Word (.docx)", command=lambda: self.export("docx")).pack(side="left", padx=4)
        pdf_btn = ttk.Button(f5, text="Export PDF", command=lambda: self.export("pdf"))
        pdf_btn.pack(side="left", padx=4)
        if not HAVE_PDF:
            pdf_btn.state(["disabled"])
        ttk.Button(f5, text="New invoice (clear rows)", command=self.new_invoice).pack(side="right", padx=4)

    # ---- break / hours
    def _toggle_break(self):
        state = ["disabled"] if self.v_nobreak.get() else ["!disabled"]
        self.break_start_entry.state(state)
        self.break_end_entry.state(state)
        self._update_hours()

    def _break_minutes(self):
        """Return break minutes (0 if no break), or None for invalid break times."""
        if self.v_nobreak.get():
            return 0
        return break_duration(self.v_break_start.get(), self.v_break_end.get())

    def _update_hours(self):
        m = self._break_minutes()
        h = None if m is None else hours_worked(self.v_start.get(), self.v_finish.get(), m)
        self.v_hours.set(fmt_hours(h))
        try:
            rate = float(self.v_rate.get().replace("$", "").replace(",", "").strip())
            self.v_amount.set(money(h * rate) if h is not None and rate >= 0 else "-")
        except ValueError:
            self.v_amount.set("-")

    # ---- actions
    def _auto_week(self):
        dates = [parse_date(r[0]) for r in self.rows] or [parse_date(self.v_date.get())]
        dates = [d for d in dates if d]
        base = max(dates) if dates else datetime.now()
        sunday = base + timedelta(days=(6 - base.weekday()) % 7)
        self.v_week.set(sunday.strftime(DATE_FMT))

    def _read_entry(self):
        d = self.v_date.get().strip()
        if not parse_date(d):
            messagebox.showerror("Invalid date", "Enter the date as dd/mm/yyyy, e.g. 20/10/2026")
            return None
        s, f = parse_time(self.v_start.get()), parse_time(self.v_finish.get())
        if not s or not f:
            messagebox.showerror("Invalid time", "Enter times like 6:00 AM, 2:30 PM or 14:30.")
            return None
        brk = self._break_minutes()
        if brk is None:
            messagebox.showerror("Invalid break", "Enter valid Break Start and Break End times, or tick 'No break'.")
            return None
        h = hours_worked(self.v_start.get(), self.v_finish.get(), brk)
        if h is None:
            messagebox.showerror("Invalid break", "The break is longer than the shift.")
            return None
        try:
            rate = float(self.v_rate.get().replace("$", "").replace(",", "").strip())
            if rate < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Invalid hourly rate", "Enter a valid hourly rate, e.g. 30.00")
            return None
        amt = round(h * rate, 2)
        bs = "" if self.v_nobreak.get() else fmt_time(parse_time(self.v_break_start.get()))
        be = "" if self.v_nobreak.get() else fmt_time(parse_time(self.v_break_end.get()))
        return (d, fmt_time(s), bs, be, fmt_time(f), self.v_pos.get().strip(),
                self.v_depot.get().strip(), rate, amt)

    def add_row(self):
        row = self._read_entry()
        if row:
            self.rows.append(row)
            self._refresh()
            if not self.v_week.get():
                self._auto_week()

    def _selected_index(self):
        sel = self.tree.selection()
        return self.tree.index(sel[0]) if sel else None

    def update_row(self):
        i = self._selected_index()
        if i is None:
            messagebox.showinfo("Select a row", "Select a row in the table first.")
            return
        row = self._read_entry()
        if row:
            self.rows[i] = row
            self._refresh()

    def delete_row(self):
        i = self._selected_index()
        if i is not None:
            del self.rows[i]
            self._refresh()

    def _on_select(self, _event):
        i = self._selected_index()
        if i is None:
            return
        d, start, break_start, break_end, finish, pos, dep, rate, amt = self.rows[i]
        self.v_date.set(d)
        self.v_start.set(start)
        self.v_finish.set(finish)
        self.v_nobreak.set(not break_start and not break_end)
        if break_start:
            self.v_break_start.set(break_start)
        if break_end:
            self.v_break_end.set(break_end)
        self.v_pos.set(pos)
        self.v_depot.set(dep)
        self.v_rate.set(f"{rate:.2f}")
        self._toggle_break()

    def _refresh(self):
        self.tree.delete(*self.tree.get_children())
        total_hours = 0.0
        for d, start, break_start, break_end, finish, pos, dep, rate, amt in self.rows:
            brk = break_duration(break_start, break_end) or 0
            h = hours_worked(start, finish, brk)
            total_hours += h or 0
            self.tree.insert("", "end", values=(d, start, break_text(break_start, break_end), finish,
                                                fmt_hours(h), pos, dep, money(amt)))
        _, _, total, gst_label = compute_totals(self.rows, self.v_gst.get())
        self.v_total.set(f"Hours: {total_hours:.2f}     GST: {gst_label}     Total: {money(total)}")

    def _persist(self):
        s = {k: self.v[k].get().strip() for k in self.v}
        s["position"] = self.v_pos.get().strip()
        s["depot"] = self.v_depot.get().strip()
        s["start"] = self.v_start.get().strip()
        s["finish"] = self.v_finish.get().strip()
        s["no_break"] = bool(self.v_nobreak.get())
        s["break_start"] = self.v_break_start.get().strip()
        s["break_end"] = self.v_break_end.get().strip()
        s["hourly_rate"] = self.v_rate.get().strip()
        try:
            s["break_min"] = int(break_duration(s["break_start"], s["break_end"]) or 0)
        except (TypeError, ValueError):
            s["break_min"] = 30
        s["gst_mode"] = self.v_gst.get()
        try:
            s["next_invoice"] = int(self.v_number.get())
        except ValueError:
            s["next_invoice"] = int(self.settings["next_invoice"])
        self.settings = s
        save_settings(s)

    def export(self, kind):
        if not self.rows:
            messagebox.showwarning("No entries", "Add at least one work entry first.")
            return
        if not self.v_week.get().strip():
            self._auto_week()
        number = self.v_number.get().strip()
        inv = {
            "number": number,
            "bill_to": self.v["bill_to"].get().strip(),
            "week_ending": self.v_week.get().strip(),
            "rows": list(self.rows),
            "gst_mode": self.v_gst.get(),
            **{k: self.v[k].get().strip() for k in ("name", "address", "abn", "bank", "bsb", "acc")},
        }
        ext = ".docx" if kind == "docx" else ".pdf"
        path = filedialog.asksaveasfilename(
            defaultextension=ext, initialfile=f"Invoice{number}{ext}",
            filetypes=[("Word document", "*.docx")] if kind == "docx" else [("PDF", "*.pdf")])
        if not path:
            return
        try:
            (export_docx if kind == "docx" else export_pdf)(path, inv)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Export failed", str(e))
            return
        self._persist()
        messagebox.showinfo("Saved", f"Invoice saved to:\n{path}")

    def new_invoice(self):
        if self.rows and not messagebox.askyesno("New invoice", "Clear all entries and move to the next invoice number?"):
            return
        try:
            self.v_number.set(f"{int(self.v_number.get()) + 1:03d}")
        except ValueError:
            pass
        self.rows.clear()
        self.v_week.set("")
        self._persist()
        self._refresh()


if __name__ == "__main__":
    InvoiceApp().mainloop()
