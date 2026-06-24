"""
PDF versions of the reports, generated directly with reportlab.

These are built fresh from the data (not converted from the Excel files), so they
work anywhere — including hosts without LibreOffice. Same numbers and layout idea
as the Excel reports; meant for quick viewing/printing in the browser.
"""

from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

import fleet_logic

NAVY = colors.HexColor("#1F4E79")
LIGHT = colors.HexColor("#DCE6F1")
ALT = colors.HexColor("#F3F6FC")
LINE = colors.HexColor("#BFBFBF")
GREY = colors.HexColor("#595959")
FLEET_NAME = "MASTER FLEET"

_title = ParagraphStyle("title", fontName="Helvetica-Bold", fontSize=13,
                        textColor=colors.white, alignment=1, leading=15)
_subtitle = ParagraphStyle("subtitle", fontName="Helvetica-Oblique", fontSize=9,
                           textColor=GREY, alignment=1, leading=12)
_info = ParagraphStyle("info", fontName="Helvetica", fontSize=10, leading=15)
_note = ParagraphStyle("note", fontName="Helvetica", fontSize=8, textColor=GREY, leading=11)
_hcell = ParagraphStyle("hcell", fontName="Helvetica-Bold", fontSize=10,
                        textColor=colors.white, alignment=1, leading=12)
_sumhdr = ParagraphStyle("sumhdr", fontName="Helvetica-Bold", fontSize=10,
                         textColor=colors.white, alignment=0, leading=12)


def _m(value):
    """1234.5 -> '1,234.50'."""
    return "{:,.2f}".format(float(value))


def _title_bar(text, usable_w):
    return Table([[Paragraph(text, _title)]], colWidths=[usable_w],
                 style=TableStyle([
                     ("BACKGROUND", (0, 0), (-1, -1), NAVY),
                     ("TOPPADDING", (0, 0), (-1, -1), 7),
                     ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
                 ]))


def write_driver_pdf(driver, combined, period, path):
    rate = fleet_logic.FIRM_RATE
    rows = combined[combined["driver"] == driver]

    body, t_gross = [], 0.0
    t_comm = t_cash = t_tips = t_bonus = t_firm = t_transfer = 0.0
    for r in rows.itertuples():
        firm = round(r.gross * rate, 2)
        transfer = round(r.gross - r.platform_commission - r.cash + r.tips + r.bonus - firm, 2)
        body.append([r.platform, _m(r.gross), _m(r.platform_commission), _m(r.cash),
                     _m(r.tips), _m(r.bonus), _m(firm), _m(transfer)])
        t_gross += r.gross; t_comm += r.platform_commission; t_cash += r.cash
        t_tips += r.tips; t_bonus += r.bonus; t_firm += firm; t_transfer += transfer

    net = round(t_gross - t_comm - t_firm, 2)

    header = ["Platform", "Gross", "Platform commission", "Cash collected",
              "Tips", "Bonus", "Firm commission (7%)", "Bank transfer"]
    header = [Paragraph(h, _hcell) for h in header]
    total = ["TOTAL", _m(t_gross), _m(t_comm), _m(t_cash), _m(t_tips), _m(t_bonus),
             _m(t_firm), _m(t_transfer)]

    usable_w = landscape(A4)[0] - 24 * mm
    col_ratios = [33, 30, 40, 33, 24, 24, 40, 33]
    col_w = [r / sum(col_ratios) * usable_w for r in col_ratios]
    table = Table([header] + body + [total], colWidths=col_w, repeatRows=1, hAlign="LEFT")
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
        ("ALIGN", (0, 0), (0, -1), "LEFT"),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 1), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("GRID", (0, 0), (-1, -1), 0.4, LINE),
        ("BACKGROUND", (0, -1), (-1, -1), LIGHT),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
    ]))

    summary = [
        ["Gross earnings", _m(t_gross)],
        ["Platform commission", _m(t_comm)],
        ["Firm commission (7%)", _m(t_firm)],
        ["Net earnings", _m(net)],
        ["Tips", _m(t_tips)],
        ["Bonus", _m(t_bonus)],
        ["Cash already collected", _m(t_cash)],
        ["Bank transfer to driver", _m(t_transfer)],
    ]
    sum_table = Table([[Paragraph("PAYMENT SUMMARY", _sumhdr), ""]] + summary,
                      colWidths=[70 * mm, 45 * mm], hAlign="LEFT")
    sum_table.setStyle(TableStyle([
        ("SPAN", (0, 0), (1, 0)),
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 1), (-1, -1), 10),
        ("ALIGN", (1, 1), (1, -1), "RIGHT"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LINEBELOW", (0, 1), (-1, -2), 0.4, LINE),
        ("BACKGROUND", (0, 4), (-1, 4), LIGHT),
        ("FONTNAME", (0, 4), (-1, 4), "Helvetica-Bold"),
        ("BACKGROUND", (0, 8), (-1, 8), LIGHT),
        ("FONTNAME", (0, 8), (-1, 8), "Helvetica-Bold"),
    ]))

    doc = SimpleDocTemplate(path, pagesize=landscape(A4),
                            leftMargin=12 * mm, rightMargin=12 * mm,
                            topMargin=12 * mm, bottomMargin=12 * mm,
                            title=f"{driver} - {period}")
    doc.build([
        _title_bar(FLEET_NAME, usable_w),
        Spacer(1, 4),
        Paragraph("Weekly driver payment report", _subtitle),
        Spacer(1, 12),
        Paragraph(f"<b>Driver:</b> {driver}", _info),
        Paragraph(f"<b>Period:</b> {period}", _info),
        Paragraph(f"<b>Firm commission rate:</b> {round(rate * 100)}%", _info),
        Paragraph(f"<b>Generated:</b> {datetime.now().strftime('%d.%m.%Y %H:%M')}", _info),
        Spacer(1, 12),
        table,
        Spacer(1, 8),
        sum_table,
        Spacer(1, 10),
        Paragraph("Amounts in RON.", _note),
    ])


def write_fleet_summary_pdf(combined, period, path):
    s = fleet_logic.compute_summary(combined)

    header = ["Driver", "Gross", "Platform commission", "Firm commission (7%)",
              "Net earnings", "Tips", "Bonus", "Cash collected", "Bank transfer"]
    header = [Paragraph(h, _hcell) for h in header]

    body = []
    for t in s.itertuples():
        body.append([t.driver, _m(t.gross), _m(t.platform_commission), _m(t.firm_commission),
                     _m(t.net), _m(t.tips), _m(t.bonus), _m(t.cash), _m(t.bank_transfer)])

    total = ["TOTAL",
             _m(s["gross"].sum()), _m(s["platform_commission"].sum()), _m(s["firm_commission"].sum()),
             _m(s["net"].sum()), _m(s["tips"].sum()), _m(s["bonus"].sum()),
             _m(s["cash"].sum()), _m(s["bank_transfer"].sum())]

    col_w = [41, 29, 29, 29, 29, 29, 29, 29, 29]
    table = Table([header] + body + [total], colWidths=[w * mm for w in col_w], repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), NAVY),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
        ("ALIGN", (0, 0), (0, -1), "LEFT"),
        ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 1), (-1, -1), 8.5),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("GRID", (0, 0), (-1, -1), 0.4, LINE),
        ("BACKGROUND", (0, -1), (-1, -1), LIGHT),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
    ]
    for i in range(1, len(body) + 1):
        if i % 2 == 0:
            style.append(("BACKGROUND", (0, i), (-1, i), ALT))
    table.setStyle(TableStyle(style))

    usable_w = landscape(A4)[0] - 24 * mm
    doc = SimpleDocTemplate(path, pagesize=landscape(A4),
                            leftMargin=12 * mm, rightMargin=12 * mm,
                            topMargin=12 * mm, bottomMargin=12 * mm,
                            title=f"Fleet summary - {period}")
    doc.build([
        _title_bar(f"{FLEET_NAME} \u2014 Fleet summary", usable_w),
        Spacer(1, 6),
        Paragraph(f"Period: {period} &nbsp;|&nbsp; {len(s)} drivers &nbsp;|&nbsp; "
                  f"Generated: {datetime.now().strftime('%d.%m.%Y %H:%M')}", _subtitle),
        Spacer(1, 12),
        table,
        Spacer(1, 10),
        Paragraph("Amounts in RON.", _note),
    ])
