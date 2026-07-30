#!/usr/bin/env python3
"""Generate Sustainability Dashboard data-flow documentation as PDF."""

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm, mm
from reportlab.platypus import (
    HRFlowable,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

OUTPUT = "/workspace/docs/Sustainability-Dashboard-Data-Flow.pdf"

DARK_GREEN = colors.HexColor("#1B5E20")
MID_GREEN = colors.HexColor("#2E7D32")
LIGHT_GREEN = colors.HexColor("#E8F5E9")
ACCENT = colors.HexColor("#FF8F00")
DARK_TEXT = colors.HexColor("#212121")
GRAY = colors.HexColor("#757575")


def build_styles():
    base = getSampleStyleSheet()
    return {
        "cover_title": ParagraphStyle(
            "cover_title",
            parent=base["Title"],
            fontSize=26,
            leading=32,
            textColor=colors.white,
            alignment=TA_CENTER,
            spaceAfter=14,
        ),
        "cover_sub": ParagraphStyle(
            "cover_sub",
            parent=base["Normal"],
            fontSize=13,
            leading=18,
            textColor=colors.HexColor("#C8E6C9"),
            alignment=TA_CENTER,
        ),
        "h1": ParagraphStyle(
            "h1",
            parent=base["Heading1"],
            fontSize=18,
            leading=24,
            textColor=DARK_GREEN,
            spaceBefore=16,
            spaceAfter=8,
        ),
        "h2": ParagraphStyle(
            "h2",
            parent=base["Heading2"],
            fontSize=13,
            leading=18,
            textColor=MID_GREEN,
            spaceBefore=12,
            spaceAfter=6,
        ),
        "body": ParagraphStyle(
            "body",
            parent=base["Normal"],
            fontSize=10,
            leading=14,
            textColor=DARK_TEXT,
            spaceAfter=6,
        ),
        "bullet": ParagraphStyle(
            "bullet",
            parent=base["Normal"],
            fontSize=10,
            leading=14,
            textColor=DARK_TEXT,
            leftIndent=14,
            bulletIndent=6,
            spaceAfter=4,
        ),
        "note": ParagraphStyle(
            "note",
            parent=base["Normal"],
            fontSize=9,
            leading=12,
            textColor=GRAY,
            spaceAfter=6,
        ),
        "mono": ParagraphStyle(
            "mono",
            parent=base["Code"],
            fontSize=9,
            leading=12,
            textColor=DARK_TEXT,
            backColor=LIGHT_GREEN,
            leftIndent=8,
            rightIndent=8,
            spaceBefore=4,
            spaceAfter=8,
        ),
    }


def section_header(st, title, subtitle=None):
    out = [Paragraph(title, st["h1"])]
    if subtitle:
        out.append(Paragraph(subtitle, st["note"]))
        out.append(Spacer(1, 4))
    out.append(HRFlowable(width="100%", thickness=1, color=MID_GREEN, spaceAfter=10))
    return out


def bullets(st, items):
    out = []
    for item in items:
        if isinstance(item, tuple):
            text, level = item
            indent = 14 + level * 12
            style = ParagraphStyle(
                "b" + str(level),
                parent=st["bullet"],
                leftIndent=indent,
                bulletIndent=indent - 8,
            )
            out.append(Paragraph(f"• {text}", style))
        else:
            out.append(Paragraph(f"• {item}", st["bullet"]))
    return out


def table(st, headers, rows, col_widths=None):
    data = [headers] + [[Paragraph(str(c), st["body"]) for c in row] for row in rows]
    tbl = Table(data, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), MID_GREEN),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 9),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#BDBDBD")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, LIGHT_GREEN]),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    return tbl


def flow_box(st, lines):
    text = "<br/>".join(lines)
    return Paragraph(text, st["mono"])


def cover_page(st):
    return [
        Spacer(1, 5 * cm),
        Paragraph("Sustainability Dashboard", st["cover_title"]),
        Paragraph("Dokumentasi Alur Data &amp; Rumus Bisnis", st["cover_sub"]),
        Spacer(1, 1.2 * cm),
        Paragraph(
            "Mill Onboarding · Supplier Due Diligence · Modul Terkait",
            st["cover_sub"],
        ),
        Spacer(1, 0.6 * cm),
        Paragraph(
            "Sumber valid dari codebase (main.js, GoogleAppsScript-backend-v3-full.gs)",
            st["cover_sub"],
        ),
        PageBreak(),
    ]


def build_pdf():
    st = build_styles()
    doc = SimpleDocTemplate(
        OUTPUT,
        pagesize=A4,
        rightMargin=1.8 * cm,
        leftMargin=1.8 * cm,
        topMargin=1.8 * cm,
        bottomMargin=1.8 * cm,
        title="Sustainability Dashboard - Data Flow Documentation",
        author="Generated from codebase",
    )

    story = cover_page(st)

    # TOC
    story += section_header(st, "Daftar Isi", "Ringkasan 15 bagian")
    toc = [
        "1. Arsitektur Sistem",
        "2. Supplier Due Diligence (SDD)",
        "3. SDD — State Machine & Alur",
        "4. SDD → Kemana Data Mengalir",
        "5. Mapping SDD → Mill Onboarding",
        "6. Mapping SDD FFB → TTP/TTM",
        "7. Mill Onboarding — Entry & Output",
        "8. Supply Task List",
        "9. Rumus Google Sheet (Read-Only)",
        "10. Rumus App — Grievance, TTM/TTP, EUDR",
        "11. Mill Risk Reason",
        "12. Peta Modul & Koneksi",
        "13. Daftar Google Sheets",
        "14. SDD vs EUDR DDS",
        "15. Aturan Validasi",
    ]
    story += bullets(st, toc)
    story.append(PageBreak())

    # 1 Architecture
    story += section_header(
        st,
        "1. Arsitektur Sistem",
        "Frontend → API Proxy → Google Apps Script → Google Sheets",
    )
    story += bullets(
        st,
        [
            "Frontend: Vite SPA (src/main.js + modul panel per fitur)",
            "API: /api/gas-proxy → doGet / doPost di Google Apps Script",
            "Database: Google Sheets — setiap modul = satu atau lebih tab sheet",
            "Backend canonical: scripts/GoogleAppsScript-backend-v3-full.gs",
            "Prinsip penting:",
            ("Skor dan risk level mill dihitung di Google Sheet (formula), bukan di web app", 1),
            ("Web app tidak menulis ulang kolom formula saat save", 1),
        ],
    )
    story.append(Spacer(1, 8))
    story.append(
        flow_box(
            st,
            [
                "[Supplier DD] [Mill Onboarding] [TTM/TTP] [EUDR] [Monthly Report]",
                "                            ↓",
                "                      /api/gas-proxy",
                "                            ↓",
                "                   Google Apps Script",
                "                            ↓",
                "         SDD_MAIN · Mill Profile · Monitoring TTP/TTM · ...",
            ],
        )
    )
    story.append(PageBreak())

    # 2 SDD
    story += section_header(
        st,
        "2. Supplier Due Diligence (SDD)",
        "Panel: #panel-supplier-dd",
    )
    story += bullets(
        st,
        [
            "Penyimpanan: SDD_MAIN (1 baris/submission), SDD_MILL_LIST (Traceability A), SDD_FFB_LIST (Traceability B)",
            "Primary key: submission_id (format SUB-YYYYMMDD-XXXXXX)",
            "Supplier type: MILL, KCP, atau TRADER",
            "Entry point:",
            ("Upload Excel — Main Form + Traceability A/B", 1),
            ("Screening form manual — SCR fields, GRV, PRI", 1),
            ("Approve / Hold / Reject oleh approver", 1),
            ("NBL Check vs sheet NBL & Unilever NBL", 1),
            "Legacy: sheet SDD Data (flat) masih didukung via adapter",
        ],
    )
    story.append(PageBreak())

    # 3 SDD state machine
    story += section_header(st, "3. SDD — State Machine & Alur")
    story.append(
        flow_box(
            st,
            [
                "Import Excel",
                "    ↓ Save Draft → SCR Status: Draft",
                "    ↓ Submit     → SCR Status: Submitted",
                "    ↓ Boss Decision:",
                "         Approve → statusSDD = APPROVED",
                "         Hold    → statusSDD = ON HOLD",
                "         Reject  → statusSDD = REJECTED",
                "    ↓ (APPROVED + Submitted) → Mill Onboarding Task List",
                "    ↓ '+ Add to Mill' → Mill Onboarding Profile",
                "    ↓ mill_ttp_sync → Monitoring TTP/TTM",
                "    ↓ mill_added = true → hilang dari Task List",
            ],
        )
    )
    story.append(Spacer(1, 10))
    story.append(
        Paragraph(
            "<b>Catatan:</b> TTP tidak di-sync saat SDD di-approve saja. "
            "Sync TTP terjadi setelah Mill disimpan dari Task List dengan payload mill_ttp_sync.",
            st["body"],
        )
    )
    story.append(PageBreak())

    # 4 SDD downstream
    story += section_header(st, "4. SDD → Kemana Data Mengalir?")
    story.append(
        table(
            st,
            ["Trigger", "Target Modul", "Fungsi / Mekanisme"],
            [
                ["statusSDD = APPROVED", "Contact List Supplier", "syncContactFromApprovedSubmission_()"],
                ["APPROVED + Submitted", "Mill Task List", "isApprovedSubmittedSddRow_()"],
                ["Save Mill dari Task List", "Mill Onboarding Profile", "add / update sheet mill"],
                ["mill_ttp_sync", "Monitoring TTP/TTM", "syncTtpFromMillOnboarding_*()"],
                ["NBL Check", "SCR - No Buy List", "runSddNblCheck()"],
                ["Monthly Report", "Section 01 SDD", "fetchSddList()"],
                ["Export PDF", "SDD Screening PDF", "sddExportPdf()"],
            ],
            col_widths=[4.2 * cm, 4.5 * cm, 7.5 * cm],
        )
    )
    story.append(PageBreak())

    # 5 SDD to Mill mapping
    story += section_header(
        st,
        "5. Mapping Field: SDD → Mill Onboarding",
        "Fungsi: mapSddRowToMillPayload() / mapSddTmlRowToMillPayload()",
    )
    story.append(
        table(
            st,
            ["Kolom Mill", "Sumber SDD"],
            [
                ["GROUP NAME", "Group Name"],
                ["COMPANY NAME", "Company Name"],
                ["MILL NAME", "Mill Name / KCP Name / TML - Mill Name"],
                ["COORDINATES", "Latitude + Longitude (comma-separated)"],
                ["MILL CAPACITY (TON/HOUR)", "Mill Capacity (Ton/Hour)"],
                ["MILL CATEGORY", "Mill Category / KCP Category"],
                ["ADDRESS", "Mill Address / KCP Address"],
                ["MONTH, YEAR", "Periode SDD"],
                ["SOURCE TYPE", "MILL / KCP / TRADER"],
                ["TRADER NAME", "Nama trader parent (jika applicable)"],
                ["UML ID", "TML - UML ID (TML lines)"],
            ],
            col_widths=[5.5 * cm, 10.7 * cm],
        )
    )
    story.append(PageBreak())

    # 6 FFB to TTP
    story += section_header(
        st,
        "6. Mapping: SDD FFB → Monitoring TTP/TTM",
        "Prinsip: 1 baris TTP = 1 pemasok FFB (Traceability B) + konteks Main Form",
    )
    story.append(
        table(
            st,
            ["Kolom TTP", "Sumber SDD"],
            [
                ["GROUP / COMPANY / MILL NAME", "Mill Onboarding identity"],
                ["FFB SUPPLIER GROUP / NAME", "FFB - Supplier Group/Name"],
                ["CATEGORY", "FFB - Supplier Category"],
                ["LAT, LONG", "FFB - Latitude/Longitude"],
                ["VILLAGE, SUBDISTRICT, DISTRICT", "FFB location fields"],
                ["CONCESION / PLANTED AREA", "FFB area fields"],
                ["LEGALITAS, ISPO/RSPO/ISCC", "FFB certification"],
                ["FFB SUPPLY to MILL (TON)", "FFB - Total Supply FFB (Ton)"],
                ["submission_id, ffb_line_id", "Metadata sync (upsert)"],
            ],
            col_widths=[5.5 * cm, 10.7 * cm],
        )
    )
    story.append(
        Paragraph(
            "Referensi lengkap: docs/sdd-to-ttm-ttp-field-mapping.md",
            st["note"],
        )
    )
    story.append(PageBreak())

    # 7 Mill onboarding
    story += section_header(
        st,
        "7. Mill Onboarding — Entry & Output",
        "Sheet: Mill Onboarding Profile (CPO/PK) · Mill Onboarding Waste (POME/SHELL)",
    )
    story.append(Paragraph("Cara data masuk:", st["h2"]))
    story += bullets(
        st,
        [
            "Manual Add/Edit — modal #btn-add-mill",
            "SDD Task List — '+ Add to Mill' dari approved screening",
            "Supply Excel Import — upload → draft → submit",
        ],
    )
    story.append(Paragraph("Supply routing:", st["h2"]))
    story += bullets(
        st,
        [
            "CPO/PK → sheet mill",
            "POME (ISCC/INS) / SHELL → sheet millWaste",
            "KCP plant → FACILITY NAME PK",
            "Refinery plant → FACILITY NAME CPO",
        ],
    )
    story.append(Spacer(1, 8))
    story.append(Paragraph("Data mengalir ke:", st["h2"]))
    story.append(
        table(
            st,
            ["Consumer", "Join Key", "Data Dipakai"],
            [
                ["Monitoring TTP/TTM", "GROUP|COMPANY|MILL|UML", "Identity + FFB sync"],
                ["Monthly Report §02", "Period + risk level", "High-risk MILL rows"],
                ["Performa Facility", "FACILITY NAME CPO/PK", "Supply, traceability %"],
                ["Questionnaire Monitoring", "group|company|mill", "Registry mill"],
                ["EUDR Potential", "group|company|mill", "Identity sync"],
                ["Mill Executive Report", "Quarter", "Supply, risk, NBL, cert"],
            ],
            col_widths=[4.2 * cm, 4.5 * cm, 7.5 * cm],
        )
    )
    story.append(PageBreak())

    # 8 Supply task list
    story += section_header(st, "8. Supply Task List — Alur")
    story.append(
        flow_box(
            st,
            [
                "Upload Excel (PLANT, CATEGORY, COMPANY, MILL, QTY)",
                "    ↓ Parse & match profil mill",
                "    ↓ match_status: matched | new | group_mismatch",
                "    ↓ Save Draft → Supply Import Draft",
                "    ↓ Review di Task List modal",
                "    ↓ Submit → submitSupplyDraft_",
                "    ↓ Append row → Mill Profile / Mill Waste",
                "    ↓ Google Sheet hitung kolom formula",
            ],
        )
    )
    story.append(Spacer(1, 10))
    story += bullets(
        st,
        [
            "CPO/PK submit: kolom produk lain dikosongkan (CPO submit clear PK, dst.)",
            "Formula columns di-restore dari baris di atas — tidak di-overwrite",
            "Repair drift: reconcileSupplyDraft_()",
        ],
    )
    story.append(PageBreak())

    # 9 Sheet formulas
    story += section_header(
        st,
        "9. Rumus di Google Sheet (Read-Only dari Web)",
        "Kolom MILL_FORMULA_HEADERS_ — tidak pernah ditulis dari web app",
    )
    story += bullets(
        st,
        [
            "Skor: SCORE, DEFORESTATION/BURN/PEAT SCORE, TOTAL SCORE SPATIAL, TOTAL SCORE, LEGALITY SCORE",
            "Agregat: TOTAL GRIEVANCES, TOTAL POLICY, TOTAL CERTIFICATION",
            "Risk: RISK LEVEL, RESULT RISK LEVEL, BUYER NO BUY LIST, PRIORITY ENGAGEMENT, SUPPLIER STATUS",
            "Supply %: PERCENTAGE SUPPLY CPO/PK/ISCC/INS/SHELL, PRODUCT SUPPLY",
            "Waste: TOTAL POME SUPPLY, MAX SUPPLY POME/SHELL, REMAINING STOCK, TOTAL SCORE SUPPLY",
            "Status: COMPLIMENT/NOT COMPLIMENT, DECLARATION MONITORING, CPO, PK",
            "SD Monitoring: LAST SD DATE, DAY LEFT, Status, Result, Risk Number",
        ],
    )
    story.append(PageBreak())

    # 10 App formulas
    story += section_header(st, "10. Rumus App — Grievance, TTM/TTP, EUDR")
    story.append(Paragraph("Grievance Risk Score (src/grievance-risk.js)", st["h2"]))
    story += bullets(
        st,
        [
            "Total Score = jumlah 6 indikator (skor 1–3 masing-masing)",
            "Total ≤ 8 → Low · Total ≤ 11 → Medium · Total > 11 → High",
            "Indikator: Publish, Subject, Repeat, Consequence, Group Scale, No Response",
        ],
    )
    story.append(Paragraph("TTM Coordinate % (ttpCalcTtmCoordinatePct_)", st["h2"]))
    story += bullets(
        st,
        [
            "Pool: SOURCE TYPE MILL/TRADER/REFINERY dan SUPPLY CPO/PK > 0",
            "Traceable: baris dengan koordinat valid",
            "% = Σ(qty traceable) / Σ(qty total) × 100",
        ],
    )
    story.append(Paragraph("TTP Aggregate % (ttpAggregateTotalTraceablePct_)", st["h2"]))
    story += bullets(
        st,
        [
            "Primary: Σ(Traceable Volume) / Σ(CPO/PK SUPPLY) × 100",
            "Fallback: rata-rata kolom % CPO/PK TRACEABLE",
        ],
    )
    story.append(Paragraph("EUDR Potential Status (eudrComputeStatus_)", st["h2"]))
    story += bullets(
        st,
        [
            "Semua kriteria ENABLED harus PASS → Potential",
            "Salah satu FAIL → Not Potential",
            "Kriteria: legality, millCategory (Integrated), ownPlasmaFfb (≥70%), resultRiskLevel (Low),",
            ("millLocation (APL), certification, grievance, ndpePolicy, noBuyList, deforestation", 1),
        ],
    )
    story.append(PageBreak())

    # 11 Mill risk reason
    story += section_header(
        st,
        "11. Mill Risk Reason Pills",
        "src/mill-risk-reason.js — interpretasi gap field untuk UI",
    )
    story += bullets(
        st,
        [
            "No Coordinate — koordinat kosong/invalid",
            "Legality Not Complete — LEGALITY SCORE ≠ 1",
            "Non APL Area — MILL LOC bukan APL",
            "No NDPE / No Certification",
            "Deforestation — RISK REDUCTION FACTOR = 1 atau width > 0 ha",
            "High Deforestation — factor = 2 atau width > 25 ha",
            "On No Buy List — BUYER NO BUY LIST = Yes",
            "Mills dengan RESULT RISK LEVEL = Low tidak menampilkan reason pills",
        ],
    )
    story.append(PageBreak())

    # 12 Module map
    story += section_header(st, "12. Peta Modul & Koneksi")
    story.append(
        flow_box(
            st,
            [
                "SDD (APPROVED) ──→ Contact List",
                "SDD ──→ Mill Task List ──→ Mill Profile ──→ TTP sync",
                "Supply Excel ──→ Mill Profile",
                "Mill ──→ EUDR Potential    TTP FFB ──→ EUDR FFB %",
                "Mill + TTP + Grievance + NBL + EUDR ──→ Monthly Report",
                "Mill + TTP + Supplied CPO/PK ──→ Performa Facility",
                "NBL ──→ SDD screening + Mill BUYER NO BUY LIST (sheet formula)",
            ],
        )
    )
    story.append(PageBreak())

    # 13 Sheets
    story += section_header(st, "13. Daftar Google Sheets (Database)")
    story.append(
        table(
            st,
            ["Key", "Tab Name", "Modul"],
            [
                ["mill", "Mill Onboarding Profile", "Mill Onboarding"],
                ["millWaste", "Mill Onboarding Waste", "Mill Waste"],
                ["supplyDraft", "Supply Import Draft", "Supply Task List"],
                ["sddMain", "SDD_MAIN", "Supplier Due Diligence"],
                ["sddMill", "SDD_MILL_LIST", "SDD Traceability A"],
                ["sddFfb", "SDD_FFB_LIST", "SDD Traceability B"],
                ["ttp", "Monitoring TTP/TTM", "TTM/TTP"],
                ["grievance", "Grievance Monitoring", "Grievance"],
                ["nbl", "NBL / Unilever NBL", "No Buy List"],
                ["eudrPotential", "EUDR Potential", "EUDR Potential"],
                ["eudrDds", "EUDR DDS (+ child)", "EUDR DDS export EU"],
                ["contactSupplier", "Contact List Supplier", "Contact List"],
            ],
            col_widths=[3.2 * cm, 5.8 * cm, 7.2 * cm],
        )
    )
    story.append(PageBreak())

    # 14 SDD vs DDS
    story += section_header(st, "14. Perbedaan: SDD vs EUDR DDS")
    story.append(
        table(
            st,
            ["Aspek", "SDD", "EUDR DDS"],
            [
                ["Panel", "#panel-supplier-dd", "#panel-due-diligence-statement"],
                ["Storage", "SDD_MAIN / MILL / FFB", "EUDR DDS + child sheets"],
                ["Tujuan", "Screening internal + onboarding", "Paket export EU EUDR"],
                ["Output", "Task List → Mill → TTP", "DOCX/PDF untuk buyer"],
                ["Code", "src/main.js", "src/dds-ui.js"],
            ],
            col_widths=[3 * cm, 6.5 * cm, 6.7 * cm],
        )
    )
    story.append(Spacer(1, 10))
    story.append(
        Paragraph(
            "Keduanya terpisah — tidak ada coupling langsung di codebase.",
            st["body"],
        )
    )
    story.append(PageBreak())

    # 15 Validation
    story += section_header(st, "15. Aturan Validasi & Guard Rules")
    story += bullets(
        st,
        [
            "SDD Submitted tidak bisa downgrade ke Draft",
            "Task List: SCR Status = submitted AND statusSDD = APPROVED",
            "TTP sync: APPROVED + submitted + identity lengkap (GROUP/COMPANY/MILL/UML)",
            "Supply submit: match_status = matched atau new; company wajib",
            "Mill save: kolom formula di-strip sebelum write",
            "EUDR: semua kriteria enabled harus pass untuk Potential",
            "Grievance: 6 indikator wajib diisi sebelum klasifikasi",
            "Koordinat: normalizeCoordinate() + recoverCoord() untuk locale Indonesia",
        ],
    )
    story.append(Spacer(1, 20))
    story.append(HRFlowable(width="100%", thickness=1, color=MID_GREEN))
    story.append(Spacer(1, 8))
    story.append(
        Paragraph(
            "<i>Dokumentasi valid dari codebase · Tidak ada perubahan kode aplikasi</i>",
            st["note"],
        )
    )

    def on_page(canvas, doc_):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(GRAY)
        canvas.drawString(1.8 * cm, 1.2 * cm, "Sustainability Dashboard — Data Flow Documentation")
        canvas.drawRightString(A4[0] - 1.8 * cm, 1.2 * cm, f"Halaman {doc_.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    print(f"Saved: {OUTPUT}")


if __name__ == "__main__":
    build_pdf()
