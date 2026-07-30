#!/usr/bin/env python3
"""Generate slide-style PDF (landscape, PPT-like) for Sustainability Dashboard data flows."""

import textwrap
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph

OUTPUT = "/workspace/docs/Sustainability-Dashboard-Data-Flow.pdf"

# Widescreen slide 13.333" × 7.5" (standard PowerPoint 16:9)
SW, SH = 13.333 * inch, 7.5 * inch
MARGIN = 0.55 * inch
HEADER_H = 0.95 * inch
FOOTER_H = 0.38 * inch

C_DARK = colors.HexColor("#1B5E20")
C_MID = colors.HexColor("#2E7D32")
C_LIGHT = colors.HexColor("#E8F5E9")
C_ACCENT = colors.HexColor("#F57C00")
C_WHITE = colors.white
C_TEXT = colors.HexColor("#263238")
C_MUTED = colors.HexColor("#607D8B")
C_LINE = colors.HexColor("#CFD8DC")


class SlideDeck:
    def __init__(self, path):
        self.c = canvas.Canvas(path, pagesize=(SW, SH))
        self.page = 0
        self.footer_left = "Sustainability Dashboard — Data Flow Documentation"

    def _styles(self):
        return {
            "title": ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=22, leading=26, textColor=C_WHITE),
            "subtitle": ParagraphStyle("s", fontName="Helvetica", fontSize=11, leading=14, textColor=colors.HexColor("#C8E6C9")),
            "h2": ParagraphStyle("h2", fontName="Helvetica-Bold", fontSize=13, leading=16, textColor=C_MID),
            "body": ParagraphStyle("b", fontName="Helvetica", fontSize=10.5, leading=14, textColor=C_TEXT),
            "small": ParagraphStyle("sm", fontName="Helvetica", fontSize=9, leading=12, textColor=C_MUTED),
            "bullet": ParagraphStyle("bu", fontName="Helvetica", fontSize=10, leading=13.5, textColor=C_TEXT, leftIndent=14),
            "box": ParagraphStyle("bx", fontName="Helvetica-Bold", fontSize=9.5, leading=12, textColor=C_WHITE, alignment=TA_CENTER),
            "box_dark": ParagraphStyle("bxd", fontName="Helvetica-Bold", fontSize=9.5, leading=12, textColor=C_TEXT, alignment=TA_CENTER),
        }

    def finish(self):
        self.c.save()

    def _footer(self):
        self.c.setFont("Helvetica", 8)
        self.c.setFillColor(C_MUTED)
        self.c.drawString(MARGIN, 0.22 * inch, self.footer_left)
        self.c.drawRightString(SW - MARGIN, 0.22 * inch, f"Slide {self.page}")

    def _header(self, title, subtitle=None):
        self.c.setFillColor(C_DARK)
        self.c.rect(0, SH - HEADER_H, SW, HEADER_H, fill=1, stroke=0)
        self.c.setFillColor(C_MID)
        self.c.rect(0, SH - HEADER_H, 0.12 * inch, HEADER_H, fill=1, stroke=0)
        st = self._styles()
        p = Paragraph(title, st["title"])
        w, h = p.wrap(SW - 2 * MARGIN, HEADER_H)
        p.drawOn(self.c, MARGIN, SH - HEADER_H + (HEADER_H - h) / 2 + (8 if subtitle else 0))
        if subtitle:
            ps = Paragraph(subtitle, st["subtitle"])
            _, sh = ps.wrap(SW - 2 * MARGIN, 20)
            ps.drawOn(self.c, MARGIN, SH - HEADER_H + (HEADER_H - h) / 2 - sh + 2)

    def new_slide(self, title=None, subtitle=None):
        if self.page:
            self._footer()
            self.c.showPage()
        self.page += 1
        self.c.setFillColor(C_WHITE)
        self.c.rect(0, 0, SW, SH, fill=1, stroke=0)
        if title:
            self._header(title, subtitle)
        return SH - HEADER_H - 0.35 * inch  # top y for content

    def cover(self, title, subtitle, lines):
        self.page += 1
        self.c.setFillColor(C_DARK)
        self.c.rect(0, 0, SW, SH, fill=1, stroke=0)
        # accent strip
        self.c.setFillColor(C_MID)
        self.c.rect(0, SH * 0.42, SW, 0.06 * inch, fill=1, stroke=0)
        st = self._styles()
        pt = ParagraphStyle("ct", parent=st["title"], fontSize=34, leading=40, alignment=TA_CENTER)
        ps = ParagraphStyle("cs", parent=st["subtitle"], fontSize=16, leading=22, alignment=TA_CENTER)
        pb = ParagraphStyle("cb", parent=st["subtitle"], fontSize=11, leading=16, alignment=TA_CENTER, textColor=colors.HexColor("#A5D6A7"))

        p = Paragraph(title, pt)
        w, h = p.wrap(SW - 2 * MARGIN, 120)
        p.drawOn(self.c, MARGIN, SH * 0.58)

        p2 = Paragraph(subtitle, ps)
        p2.wrap(SW - 2 * MARGIN, 60)
        p2.drawOn(self.c, MARGIN, SH * 0.48)

        y = SH * 0.32
        for line in lines:
            p3 = Paragraph(line, pb)
            p3.wrap(SW - 2 * MARGIN, 20)
            p3.drawOn(self.c, MARGIN, y)
            y -= 18

        self._footer()
        self.c.showPage()

    def bullets(self, top_y, items, col_x=MARGIN, col_w=None):
        if col_w is None:
            col_w = SW - 2 * MARGIN
        st = self._styles()
        y = top_y
        for item in items:
            level = item[1] if isinstance(item, tuple) else 0
            text = item[0] if isinstance(item, tuple) else item
            indent = col_x + level * 16
            style = ParagraphStyle(
                "bl", parent=st["bullet"], leftIndent=14 + level * 14, fontSize=10 - level * 0.5
            )
            p = Paragraph(f"• {text}", style)
            w, h = p.wrap(col_w - level * 16, 200)
            if y - h < FOOTER_H + 0.15 * inch:
                break
            p.drawOn(self.c, indent, y - h)
            y -= h + 3
        return y

    def draw_box(self, x, y, w, h, text, fill=C_MID, text_color=C_WHITE, radius=8, font_size=9.5):
        self.c.setFillColor(fill)
        self.c.setStrokeColor(C_DARK if fill != C_LIGHT else C_LINE)
        self.c.setLineWidth(0.8)
        self.c.roundRect(x, y, w, h, radius, fill=1, stroke=1)
        st = ParagraphStyle(
            "boxt",
            fontName="Helvetica-Bold",
            fontSize=font_size,
            leading=font_size + 2.5,
            textColor=text_color,
            alignment=TA_CENTER,
        )
        lines = text.split("\n")
        line_h = font_size + 3
        total_h = len(lines) * line_h
        start_y = y + (h + total_h) / 2 - line_h
        for i, line in enumerate(lines):
            p = Paragraph(line, st)
            pw, ph = p.wrap(w - 10, h)
            p.drawOn(self.c, x + (w - pw) / 2, start_y - i * line_h - ph + line_h)

    def draw_arrow_down(self, x, y, label=""):
        self.c.setStrokeColor(C_ACCENT)
        self.c.setFillColor(C_ACCENT)
        self.c.setLineWidth(2)
        self.c.line(x, y, x, y - 0.22 * inch)
        self.c.line(x, y - 0.22 * inch, x - 5, y - 0.14 * inch)
        self.c.line(x, y - 0.22 * inch, x + 5, y - 0.14 * inch)
        if label:
            self.c.setFont("Helvetica-Bold", 8)
            self.c.setFillColor(C_ACCENT)
            self.c.drawCentredString(x, y - 0.32 * inch, label)

    def draw_arrow_right(self, x1, y, x2):
        self.c.setStrokeColor(C_ACCENT)
        self.c.setFillColor(C_ACCENT)
        self.c.setLineWidth(2)
        self.c.line(x1, y, x2, y)
        self.c.line(x2, y, x2 - 6, y - 4)
        self.c.line(x2, y, x2 - 6, y + 4)

    def table(self, top_y, headers, rows, col_widths=None):
        if col_widths is None:
            n = len(headers)
            usable = SW - 2 * MARGIN
            col_widths = [usable / n] * n

        row_h = 0.38 * inch
        x0 = MARGIN
        y = top_y

        # header
        self.c.setFillColor(C_MID)
        self.c.rect(x0, y - row_h, sum(col_widths), row_h, fill=1, stroke=0)
        x = x0
        self.c.setFillColor(C_WHITE)
        self.c.setFont("Helvetica-Bold", 9)
        for i, h in enumerate(headers):
            self.c.drawString(x + 6, y - row_h + 10, str(h)[:40])
            x += col_widths[i]
        y -= row_h

        self.c.setFont("Helvetica", 8.5)
        for ri, row in enumerate(rows):
            bg = C_LIGHT if ri % 2 == 0 else C_WHITE
            self.c.setFillColor(bg)
            self.c.setStrokeColor(C_LINE)
            self.c.rect(x0, y - row_h, sum(col_widths), row_h, fill=1, stroke=1)
            x = x0
            self.c.setFillColor(C_TEXT)
            for ci, cell in enumerate(row):
                txt = textwrap.fill(str(cell), width=int(col_widths[ci] / 5.5)) if len(str(cell)) > 35 else str(cell)
                first_line = txt.split("\n")[0]
                self.c.drawString(x + 6, y - row_h + 10, first_line[:55])
                x += col_widths[ci]
            y -= row_h
            if y < FOOTER_H + 0.5 * inch:
                break
        return y

    def section_label(self, y, text):
        st = self._styles()
        p = Paragraph(text, st["h2"])
        w, h = p.wrap(SW - 2 * MARGIN, 30)
        p.drawOn(self.c, MARGIN, y - h)
        return y - h - 8

    def note(self, y, text):
        st = self._styles()
        p = Paragraph(text, st["small"])
        w, h = p.wrap(SW - 2 * MARGIN, 40)
        p.drawOn(self.c, MARGIN, y - h)
        return y - h


def build():
    d = SlideDeck(OUTPUT)

    d.cover(
        "Sustainability Dashboard",
        "Dokumentasi Alur Data &amp; Rumus Bisnis",
        [
            "Mill Onboarding · Supplier Due Diligence · Traceability",
            "Sumber valid dari codebase — tanpa perubahan kode aplikasi",
        ],
    )

    # Slide 2 - TOC
    y = d.new_slide("Daftar Isi", "15 slide — alur data & rumus bisnis")
    toc = [
        "1. Arsitektur Sistem",
        "2. Supplier Due Diligence (SDD)",
        "3. SDD Lifecycle (State Machine)",
        "4. SDD → Kemana Data Mengalir",
        "5. Mapping SDD → Mill Onboarding",
        "6. Mapping FFB → Monitoring TTP/TTM",
        "7. Mill Onboarding — Entry Point",
        "8. Mill Onboarding → Output",
        "9. Supply Task List",
        "10. Rumus Google Sheet",
        "11. Rumus App (Grievance · TTM · EUDR)",
        "12. Peta Modul & Koneksi",
        "13. Daftar Google Sheets",
        "14. SDD vs EUDR DDS",
        "15. Aturan Validasi",
    ]
    mid = SW / 2
    d.bullets(y, toc[:8], col_x=MARGIN, col_w=mid - MARGIN - 0.2 * inch)
    d.bullets(y, toc[8:], col_x=mid + 0.1 * inch, col_w=mid - MARGIN - 0.2 * inch)

    # Slide 3 - Architecture
    y = d.new_slide("1. Arsitektur Sistem", "Frontend → API Proxy → Google Apps Script → Google Sheets")
    boxes_top = [
        ("Supplier\nDue Diligence", 0.5 * inch),
        ("Mill\nOnboarding", 2.85 * inch),
        ("Monitoring\nTTM/TTP", 5.2 * inch),
        ("EUDR\nPotential", 7.55 * inch),
        ("Monthly\nReport", 9.9 * inch),
    ]
    bw, bh = 2.1 * inch, 0.72 * inch
    by = y - bh - 0.1 * inch
    for text, bx in boxes_top:
        d.draw_box(bx, by, bw, bh, text, fill=C_LIGHT, text_color=C_DARK)
    cx = SW / 2
    d.draw_arrow_down(cx, by - 0.05 * inch)
    d.draw_box(cx - 2.0 * inch, by - 1.05 * inch, 4.0 * inch, 0.55 * inch, "/api/gas-proxy", fill=C_ACCENT)
    d.draw_arrow_down(cx, by - 1.1 * inch)
    d.draw_box(cx - 2.6 * inch, by - 1.85 * inch, 5.2 * inch, 0.55 * inch, "Google Apps Script  (doGet / doPost)", fill=C_DARK)
    d.draw_arrow_down(cx, by - 1.9 * inch)
    d.draw_box(
        cx - 3.2 * inch, by - 2.65 * inch, 6.4 * inch, 0.55 * inch,
        "Google Sheets: SDD_MAIN · Mill Profile · TTP · EUDR · ...",
        fill=C_MID, font_size=9,
    )
    d.note(by - 3.0 * inch, "<b>Prinsip:</b> Skor/risk mill dihitung di Google Sheet. Web app tidak menulis kolom formula.")

    # Slide 4 - SDD overview
    y = d.new_slide("2. Supplier Due Diligence (SDD)", "Panel #panel-supplier-dd")
    d.bullets(
        y,
        [
            "Penyimpanan: <b>SDD_MAIN</b> · <b>SDD_MILL_LIST</b> (Traceability A) · <b>SDD_FFB_LIST</b> (Traceability B)",
            "Primary key: submission_id  (SUB-YYYYMMDD-XXXXXX)",
            "Supplier type: MILL · KCP · TRADER",
            ("Entry point:", 0),
            ("Upload Excel — Main Form + Traceability A/B", 1),
            ("Screening form — SCR fields, GRV, PRI", 1),
            ("Approve / Hold / Reject", 1),
            ("NBL Check vs sheet NBL & Unilever NBL", 1),
            ("Export PDF screening report", 1),
        ],
    )

    # Slide 5 - SDD lifecycle (vertical flow, center)
    y = d.new_slide("3. SDD Lifecycle", "State machine dari import hingga onboarding selesai")
    steps = [
        ("1. Import Excel", C_MID),
        ("2. Save Draft  →  SCR: Draft", C_MID),
        ("3. Submit  →  SCR: Submitted", C_ACCENT),
        ("4. Boss Decision: Approve / Hold / Reject", C_DARK),
        ("5. Task List  (APPROVED + Submitted)", C_MID),
        ("6. + Add to Mill  →  Mill Profile", C_MID),
        ("7. mill_ttp_sync  →  Monitoring TTP/TTM", C_ACCENT),
        ("8. mill_added = true  (selesai)", colors.HexColor("#388E3C")),
    ]
    bw2 = 5.6 * inch
    bh2 = 0.44 * inch
    bx = (SW - bw2) / 2
    sy = y - 0.08 * inch
    for i, (txt, col) in enumerate(steps):
        by2 = sy - i * 0.54 * inch
        d.draw_box(bx, by2 - bh2, bw2, bh2, txt, fill=col, font_size=9.5)
        if i < len(steps) - 1:
            d.draw_arrow_down(bx + bw2 / 2, by2 - bh2 - 0.02 * inch)
    d.note(sy - len(steps) * 0.54 * inch - 0.15 * inch, "<b>Catatan:</b> TTP sync terjadi SETELAH Mill disimpan — bukan saat approve SDD saja.")

    # Slide 6 - SDD downstream
    y = d.new_slide("4. SDD → Kemana Data Mengalir?", "Trigger · target · fungsi backend")
    d.table(
        y,
        ["Trigger", "Target", "Fungsi"],
        [
            ["APPROVED", "Contact List", "syncContactFromApprovedSubmission_"],
            ["APPROVED + Submitted", "Mill Task List", "isApprovedSubmittedSddRow_"],
            ["Save Mill", "Mill Profile", "add / update sheet mill"],
            ["mill_ttp_sync", "TTP/TTM", "syncTtpFromMillOnboarding_*"],
            ["NBL Check", "SCR - No Buy List", "runSddNblCheck()"],
            ["Report period", "Monthly Report §01", "fetchSddList()"],
        ],
        col_widths=[2.4 * inch, 2.5 * inch, 3.4 * inch],
    )

    # Slide 7 - SDD to Mill mapping
    y = d.new_slide("5. Mapping SDD → Mill Onboarding", "mapSddRowToMillPayload() / mapSddTmlRowToMillPayload()")
    d.table(
        y,
        ["Kolom Mill", "Sumber SDD"],
        [
            ["GROUP NAME", "Group Name"],
            ["COMPANY NAME", "Company Name"],
            ["MILL NAME", "Mill Name / KCP Name / TML - Mill Name"],
            ["COORDINATES", "Latitude + Longitude"],
            ["MILL CAPACITY", "Mill Capacity (Ton/Hour)"],
            ["SOURCE TYPE", "MILL / KCP / TRADER"],
            ["TRADER NAME", "Nama trader parent"],
            ["UML ID", "TML - UML ID"],
        ],
        col_widths=[3.2 * inch, 4.8 * inch],
    )

    # Slide 8 - FFB to TTP
    y = d.new_slide("6. Mapping FFB → TTP/TTM", "1 baris TTP = 1 pemasok FFB + konteks Main Form")
    d.table(
        y,
        ["Kolom TTP", "Sumber"],
        [
            ["GROUP / COMPANY / MILL", "Mill Onboarding identity"],
            ["FFB SUPPLIER GROUP/NAME", "FFB - Supplier Group/Name"],
            ["CATEGORY", "FFB - Supplier Category"],
            ["LAT / LONG", "FFB - Latitude/Longitude"],
            ["AREA / LEGALITAS / CERT", "FFB fields"],
            ["FFB SUPPLY to MILL (TON)", "FFB - Total Supply FFB"],
            ["submission_id", "Metadata upsert"],
        ],
        col_widths=[3.5 * inch, 4.5 * inch],
    )

    # Slide 9 - Mill entry
    y = d.new_slide("7. Mill Onboarding — Entry Point", "Mill Profile (CPO/PK) · Mill Waste (POME/SHELL)")
    d.bullets(
        y,
        [
            "<b>3 cara data masuk:</b>",
            ("Manual Add/Edit — modal #btn-add-mill", 1),
            ("SDD Task List — '+ Add to Mill'", 1),
            ("Supply Excel Import — upload → draft → submit", 1),
            "<b>Supply routing:</b>",
            ("CPO/PK → sheet mill", 1),
            ("POME/SHELL → sheet millWaste", 1),
            ("KCP plant → FACILITY NAME PK", 1),
            ("Refinery → FACILITY NAME CPO", 1),
        ],
    )

    # Slide 10 - Mill output
    y = d.new_slide("8. Mill Onboarding → Output", "Consumer modules & join key")
    d.table(
        y,
        ["Consumer", "Join Key", "Data"],
        [
            ["TTP/TTM", "GROUP|COMPANY|MILL", "Identity + FFB"],
            ["Monthly Report", "Period + risk", "High-risk mills"],
            ["Performa Facility", "Facility name", "Supply, trace %"],
            ["Questionnaire", "group|company|mill", "Registry"],
            ["EUDR Potential", "group|company|mill", "Identity sync"],
            ["Executive Report", "Quarter", "Supply, NBL, cert"],
        ],
        col_widths=[2.3 * inch, 2.5 * inch, 3.2 * inch],
    )

    # Slide 11 - Supply task list
    y = d.new_slide("9. Supply Task List", "Embedded di Mill Onboarding")
    flow_x = MARGIN
    flow_y = y - 0.55 * inch
    fw, fh = 2.0 * inch, 0.5 * inch
    flow_steps = ["Excel Upload", "Match Profil", "Save Draft", "Review", "Submit", "Mill Sheet"]
    for i, step in enumerate(flow_steps):
        bx = flow_x + i * (fw + 0.28 * inch)
        d.draw_box(bx, flow_y, fw, fh, step, fill=C_MID if i == len(flow_steps) - 1 else C_LIGHT, text_color=C_WHITE if i == len(flow_steps) - 1 else C_DARK, font_size=8.5)
        if i < len(flow_steps) - 1:
            d.draw_arrow_right(bx + fw + 0.02 * inch, flow_y + fh / 2, bx + fw + 0.26 * inch)
    d.bullets(
        flow_y - 0.35 * inch,
        [
            "match_status: matched · new · group_mismatch",
            "Formula columns di-restore — tidak di-overwrite",
            "Repair drift: reconcileSupplyDraft_()",
        ],
    )

    # Slide 12 - Sheet formulas
    y = d.new_slide("10. Rumus Google Sheet", "Read-only dari web — MILL_FORMULA_HEADERS_")
    left_items = [
        "SCORE · TOTAL SCORE · LEGALITY SCORE",
        "DEFORESTATION / BURN / PEAT SCORE",
        "TOTAL GRIEVANCES · POLICY · CERTIFICATION",
        "RISK LEVEL · RESULT RISK LEVEL",
        "BUYER NO BUY LIST · SUPPLIER STATUS",
    ]
    right_items = [
        "PERCENTAGE SUPPLY CPO/PK/ISCC/INS/SHELL",
        "PRODUCT SUPPLY · STATUS SUPPLY",
        "TOTAL POME SUPPLY · MAX SUPPLY",
        "REMAINING STOCK · TOTAL SCORE SUPPLY",
        "SD Monitoring: DAY LEFT · Status · Result",
    ]
    d.bullets(y, left_items, col_x=MARGIN, col_w=SW / 2 - MARGIN - 0.15 * inch)
    d.bullets(y, right_items, col_x=SW / 2 + 0.05 * inch, col_w=SW / 2 - MARGIN - 0.15 * inch)

    # Slide 13 - App formulas
    y = d.new_slide("11. Rumus App (Frontend)", "Dihitung di web app — bukan di sheet")
    cards = [
        ("Grievance Risk", "Total = Σ 6 indikator\n≤8 Low · ≤11 Med · >11 High", 0),
        ("TTM Coordinate %", "Σ(qty traceable)\n/ Σ(qty total) × 100", 1),
        ("TTP Aggregate %", "Σ(trace vol)\n/ Σ(supply) × 100", 2),
        ("EUDR Status", "All criteria PASS\n→ Potential", 3),
    ]
    cw, ch = 2.85 * inch, 1.15 * inch
    for title, body, idx in cards:
        row, col = idx // 2, idx % 2
        bx = MARGIN + col * (cw + 0.35 * inch)
        by = y - 0.2 * inch - row * (ch + 0.35 * inch)
        d.draw_box(bx, by - ch, cw, 0.32 * inch, title, fill=C_DARK, font_size=10)
        d.draw_box(bx, by - ch - 0.82 * inch, cw, 0.82 * inch, body, fill=C_LIGHT, text_color=C_DARK, font_size=9)

    # Slide 14 - Module map
    y = d.new_slide("12. Peta Modul & Koneksi", "Dependensi antar modul")
    flows = [
        "SDD (APPROVED) ──────────→ Contact List",
        "SDD ──→ Task List ──→ Mill ──→ TTP sync",
        "Supply Excel ──────────────→ Mill Profile",
        "Mill ──→ EUDR Potential    TTP ──→ EUDR FFB %",
        "Mill + TTP + GRV + NBL ───→ Monthly Report",
        "Mill + TTP + Supplied ────→ Performa Facility",
        "NBL ──→ SDD check + Mill BUYER NO BUY LIST",
    ]
    fy = y - 0.15 * inch
    for line in flows:
        d.draw_box(MARGIN, fy - 0.42 * inch, SW - 2 * MARGIN, 0.42 * inch, line, fill=C_LIGHT, text_color=C_DARK, font_size=9, radius=6)
        fy -= 0.52 * inch

    # Slide 15 - Sheets table
    y = d.new_slide("13. Daftar Google Sheets", "Database tabs per modul")
    d.table(
        y,
        ["Key", "Tab Name", "Modul"],
        [
            ["mill", "Mill Onboarding Profile", "Mill Onboarding"],
            ["millWaste", "Mill Onboarding Waste", "Mill Waste"],
            ["sddMain/Mill/Ffb", "SDD_MAIN / MILL / FFB", "Supplier DD"],
            ["ttp", "Monitoring TTP/TTM", "Traceability"],
            ["grievance", "Grievance Monitoring", "Grievance"],
            ["eudrPotential", "EUDR Potential", "EUDR"],
            ["contactSupplier", "Contact List Supplier", "Contact"],
        ],
        col_widths=[2.0 * inch, 3.5 * inch, 2.5 * inch],
    )

    # Slide 16 - SDD vs DDS
    y = d.new_slide("14. SDD vs EUDR DDS", "Dua modul terpisah — tidak ada coupling langsung")
    d.table(
        y,
        ["Aspek", "SDD", "EUDR DDS"],
        [
            ["Panel", "supplier-dd", "due-diligence-statement"],
            ["Storage", "SDD_MAIN + child", "EUDR DDS + child"],
            ["Tujuan", "Screening + onboarding", "Export paket EU"],
            ["Output", "Mill → TTP", "DOCX/PDF buyer"],
        ],
        col_widths=[1.8 * inch, 2.8 * inch, 2.8 * inch],
    )

    # Slide 17 - Validation + closing
    y = d.new_slide("15. Aturan Validasi", "Guard rules mencegah inkonsistensi data")
    d.bullets(
        y,
        [
            "SDD Submitted tidak bisa downgrade ke Draft",
            "Task List: SCR = submitted AND statusSDD = APPROVED",
            "TTP sync: identity lengkap (GROUP / COMPANY / MILL / UML)",
            "Supply submit: match_status = matched atau new",
            "Mill save: kolom formula di-strip sebelum write",
            "EUDR: semua kriteria enabled harus pass",
            "Grievance: 6 indikator wajib sebelum klasifikasi",
        ],
    )

    # Closing slide
    d.page += 1
    d.c.setFillColor(C_DARK)
    d.c.rect(0, 0, SW, SH, fill=1, stroke=0)
    st = d._styles()
    pt = ParagraphStyle("end", fontName="Helvetica-Bold", fontSize=32, textColor=C_WHITE, alignment=TA_CENTER)
    p = Paragraph("Terima Kasih", pt)
    p.wrap(SW, 60)
    p.drawOn(d.c, 0, SH * 0.55)
    ps = ParagraphStyle("ens", fontName="Helvetica", fontSize=13, textColor=colors.HexColor("#C8E6C9"), alignment=TA_CENTER)
    p2 = Paragraph("Dokumentasi valid · Tanpa perubahan kode aplikasi", ps)
    p2.wrap(SW, 30)
    p2.drawOn(d.c, 0, SH * 0.45)
    d._footer()
    d.c.showPage()

    d.finish()
    print(f"Saved: {OUTPUT} ({d.page} slides)")


if __name__ == "__main__":
    build()
