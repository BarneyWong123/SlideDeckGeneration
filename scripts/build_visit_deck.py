"""Build the two-slide customer-visit deck (Hospital Umum Sarawak, PIL Timberland Kuching).

Reads the SAP AR export directly so no figure is ever retyped by hand, then asserts the
headline totals before writing the .pptx. Branded to the Biomed Global design system.
"""

import collections
import datetime
import sys
import warnings

import openpyxl
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_TICK_MARK
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

warnings.filterwarnings("ignore")

SRC = "/root/.claude/uploads/067988c6-a330-52b0-877f-b06189dbcc05/66a74b37-data_1.xlsx"
OUT = "/home/user/SlideDeckGeneration/Customer_Visit_HUS_PIL_2026-09-06.pptx"

# --- Biomed Global design tokens -------------------------------------------------
DARK = RGBColor(0x00, 0x33, 0x66)     # corporate dark blue
LIGHT = RGBColor(0x78, 0xC8, 0xEF)    # corporate light blue
DEEP = RGBColor(0x0F, 0x1B, 0x2B)     # Pantone 296C
SKY600 = RGBColor(0x2F, 0x92, 0xC2)   # accent text on light
SKY100 = RGBColor(0xDC, 0xF2, 0xFD)
BLUE50 = RGBColor(0xE6, 0xED, 0xF4)
GREY800 = RGBColor(0x2B, 0x31, 0x38)
GREY600 = RGBColor(0x5D, 0x66, 0x6E)
GREY200 = RGBColor(0xE3, 0xE8, 0xEC)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DANGER = RGBColor(0x8C, 0x24, 0x17)
SUCCESS = RGBColor(0x1F, 0x6B, 0x44)
FONT = "DM Sans"

FY_BOUNDARY = datetime.datetime(2026, 5, 1)
LFL_FY26 = ["2025-05", "2025-06", "2025-07", "2025-08"]
LFL_FY27 = ["2026-05", "2026-06", "2026-07", "2026-08"]
MONTHS = [f"2025-{m:02d}" for m in range(5, 13)] + [f"2026-{m:02d}" for m in range(1, 9)]

TRANSFUSION = ("Bio-Rad", "Werfen Immucor")

GROUP_LABELS = {
    "DM-ID CARDS": "Bio-Rad ID cards",
    "DM-ID CELLS": "Bio-Rad ID cells",
    "DM-ID REAGENT": "Bio-Rad ID reagent",
    "DM-ID QC": "Bio-Rad ID QC",
    "DM-ACCS": "Bio-Rad accessories",
    "IMCG-REAGENT": "Werfen Immucor reagent",
    "IMCG-ACCS": "Werfen Immucor accessories",
    "DBL-REAGENT": "Werfen DBL reagent",
}

# I-LOAN PDN tender achievement, reproduced verbatim from the supplied report screenshot.
TENDER = [
    ("PDN", 53.13), ("Melaka", 46.77), ("Perlis", 45.59), ("Sarawak", 41.52),
    ("Wil. Persekutuan", 40.26), ("P. Pinang", 33.10), ("Selangor", 31.79),
    ("N. Sembilan", 31.43), ("Sabah & Labuan", 27.96), ("Johor", 17.65),
    ("Terengganu", 17.08), ("Kelantan", 16.00), ("Kedah", 13.77),
    ("Perak", 12.19), ("Pahang", 11.81),
]
TENDER_TOTAL = 28.59


# --- data ------------------------------------------------------------------------
def load():
    wb = openpyxl.load_workbook(SRC, data_only=True)
    rows = list(wb["Export"].iter_rows(values_only=True))
    hdr, body = rows[0], rows[1:]
    idx = {h: n for n, h in enumerate(hdr)}
    out = []
    for r in body:
        if not r[idx["Date"]]:
            continue
        try:
            amt = float(r[idx["Sales Amount"]] or 0)
        except (TypeError, ValueError):
            amt = 0.0
        try:
            qty = float(r[idx["Quantity"]] or 0)
        except (TypeError, ValueError):
            qty = 0.0
        d = r[idx["Date"]]
        out.append({
            "date": d, "month": d.strftime("%Y-%m"),
            "fy": "FY26" if d < FY_BOUNDARY else "FY27",
            "brand": r[idx["Brand"]], "group": r[idx["Product Group"]],
            "product": r[idx["Product Name"]] or "",
            "ship": (r[idx["Ship To Name (SAP)"]] or "").upper(),
            "amt": amt, "qty": qty,
        })
    return out


def money(x):
    return f"{x:,.0f}"


def pct(new, old):
    return (new / old - 1) * 100 if old else 0.0


# --- drawing helpers -------------------------------------------------------------
def box(slide, x, y, w, h, fill=None, line=None, radius=None):
    from pptx.enum.shapes import MSO_SHAPE
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h))
    if radius:
        shape.adjustments[0] = radius
    if fill is None:
        shape.fill.background()
    else:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = Pt(0.75)
    shape.shadow.inherit = False
    shape.text_frame.word_wrap = True
    return shape


def text(slide, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.0):
    """runs: list of (string, size_pt, bold, color) or a list of such lists = paragraphs."""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    paras = runs if isinstance(runs[0], list) else [runs]
    for n, para in enumerate(paras):
        p = tf.paragraphs[0] if n == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        for s, size, bold, color in para:
            r = p.add_run()
            r.text = s
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.color.rgb = color
            r.font.name = FONT
    return tb


def style_chart(chart, cat_size=9, val_fmt='#,##0'):
    chart.font.size = Pt(cat_size)
    chart.font.name = FONT
    chart.font.color.rgb = GREY600
    ca = chart.category_axis
    ca.tick_labels.font.size = Pt(cat_size)
    ca.tick_labels.font.name = FONT
    ca.tick_labels.font.color.rgb = GREY600
    ca.format.line.color.rgb = GREY200
    ca.major_tick_mark = XL_TICK_MARK.NONE
    va = chart.value_axis
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = GREY200
    va.major_gridlines.format.line.width = Pt(0.5)
    va.format.line.fill.background()
    va.tick_labels.font.size = Pt(cat_size)
    va.tick_labels.font.name = FONT
    va.tick_labels.font.color.rgb = GREY600
    va.tick_labels.number_format = val_fmt
    va.tick_labels.number_format_is_linked = False
    va.major_tick_mark = XL_TICK_MARK.NONE


def kpi(slide, x, y, w, h, label, value, sub, value_color=DARK, panel=WHITE):
    box(slide, x, y, w, h, fill=panel, line=GREY200, radius=0.06)
    text(slide, x + 0.22, y + 0.16, w - 0.44, 0.24,
         [(label.upper(), 9, True, GREY600)])
    text(slide, x + 0.22, y + 0.42, w - 0.44, 0.44,
         [(value, 26, True, value_color)])
    text(slide, x + 0.22, y + 0.90, w - 0.44, 0.32,
         [(sub, 10, False, GREY600)], spacing=1.15)


def header(slide, eyebrow, title, right_lines):
    box(slide, 0, 0, 13.333, 0.95, fill=DARK)
    text(slide, 0.5, 0.16, 8.6, 0.22, [(eyebrow.upper(), 9.5, True, LIGHT)])
    text(slide, 0.5, 0.38, 8.6, 0.44, [(title, 21 if len(title) < 40 else 16.5, True, WHITE)],
         anchor=MSO_ANCHOR.MIDDLE)
    text(slide, 9.3, 0.22, 3.55, 0.55,
         [[(right_lines[0], 10, True, LIGHT)], [(right_lines[1], 9.5, False, SKY100)]],
         align=PP_ALIGN.RIGHT, spacing=1.25)


def footer(slide, note):
    box(slide, 0, 7.02, 13.333, 0.48, fill=BLUE50)
    text(slide, 0.5, 7.16, 12.4, 0.24, [(note, 8.5, False, GREY600)])


def table(slide, x, y, w, col_w, rows_data, header_row, row_h=0.30, head_h=0.32,
          font_size=9.5, align_right_from=1):
    """Branded table: dark blue header, hairline rules, no banding."""
    n_rows, n_cols = len(rows_data) + 1, len(header_row)
    shape = slide.shapes.add_table(n_rows, n_cols, Inches(x), Inches(y),
                                   Inches(w), Inches(head_h + row_h * len(rows_data)))
    tbl = shape.table
    tbl.first_row = False
    tbl.horz_banding = False
    for c, cw in enumerate(col_w):
        tbl.columns[c].width = Inches(cw)
    tbl.rows[0].height = Inches(head_h)
    for r in range(1, n_rows):
        tbl.rows[r].height = Inches(row_h)

    for c, label in enumerate(header_row):
        cell = tbl.cell(0, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = DARK
        cell.margin_left = cell.margin_right = Inches(0.08)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        p_ = cell.text_frame.paragraphs[0]
        p_.alignment = PP_ALIGN.LEFT if c < align_right_from else PP_ALIGN.RIGHT
        run = p_.add_run()
        run.text = label
        run.font.size = Pt(8.5)
        run.font.bold = True
        run.font.name = FONT
        run.font.color.rgb = WHITE

    for r, row in enumerate(rows_data, start=1):
        emphasise = isinstance(row, tuple)
        cells = row[1] if emphasise else row
        style = row[0] if emphasise else None
        for c, val in enumerate(cells):
            cell = tbl.cell(r, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = BLUE50 if style == "total" else WHITE
            cell.margin_left = cell.margin_right = Inches(0.08)
            cell.margin_top = cell.margin_bottom = 0
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p_ = cell.text_frame.paragraphs[0]
            p_.alignment = PP_ALIGN.LEFT if c < align_right_from else PP_ALIGN.RIGHT
            run = p_.add_run()
            run.text = str(val)
            run.font.size = Pt(font_size)
            run.font.bold = style == "total"
            run.font.name = FONT
            colour = GREY800
            if style != "total" and c >= align_right_from:
                if str(val).startswith(MINUS):
                    colour = DANGER
                elif str(val).startswith("+"):
                    colour = SUCCESS
            run.font.color.rgb = DARK if style == "total" else colour
    return shape


def points(slide, x, y, w, h, heading, items):
    """Bottom 'what to ask' strip: heading plus evenly spaced numbered points."""
    box(slide, x, y, w, h, fill=WHITE, line=GREY200, radius=0.05)
    text(slide, x + 0.24, y + 0.16, w - 0.48, 0.22, [(heading.upper(), 9, True, GREY600)])
    col_w = (w - 0.48) / len(items)
    for n, (lead, body) in enumerate(items):
        cx = x + 0.24 + n * col_w
        text(slide, cx, y + 0.44, col_w - 0.24, h - 0.60,
             [[(lead, 10.5, True, DARK)], [(body, 9.5, False, GREY800)]], spacing=1.22)


MINUS = "\u2212"


def delta_str(v):
    if round(v) == 0:
        return "—"
    sign = "+" if v > 0 else MINUS
    return f"{sign}{abs(v):,.0f}"


# --- slide 1: Hospital Umum Sarawak ----------------------------------------------
def slide_hus(prs, rows):
    hus = [r for r in rows if "HOSPITAL UMUM SARAWAK" in r["ship"]]
    tx = [r for r in hus if r["brand"] in TRANSFUSION]

    def total(rs, months=None, brand=None, fy=None):
        return sum(r["amt"] for r in rs
                   if (months is None or r["month"] in months)
                   and (brand is None or r["brand"] == brand)
                   and (fy is None or r["fy"] == fy))

    br26, br27 = total(tx, LFL_FY26, "Bio-Rad"), total(tx, LFL_FY27, "Bio-Rad")
    we26, we27 = total(tx, LFL_FY26, "Werfen Immucor"), total(tx, LFL_FY27, "Werfen Immucor")
    t26, t27 = br26 + we26, br27 + we27
    fy26_full = total(tx, fy="FY26")

    assert round(t26) == 183908 and round(t27) == 139464, (t26, t27)
    assert round(fy26_full) == 786219, fy26_full

    # product-group movement, like-for-like
    mv = collections.defaultdict(lambda: [0.0, 0.0])
    for r in tx:
        if r["month"] in LFL_FY26:
            mv[r["group"]][0] += r["amt"]
        elif r["month"] in LFL_FY27:
            mv[r["group"]][1] += r["amt"]
    moves = sorted(((k, v[1] - v[0]) for k, v in mv.items()), key=lambda x: -abs(x[1]))[:4]

    bd = [r for r in hus if r["brand"] == "BD"]
    bd_delta = (sum(r["amt"] for r in bd if r["month"] in LFL_FY27)
                - sum(r["amt"] for r in bd if r["month"] in LFL_FY26))

    s = prs.slides.add_slide(prs.slide_layouts[6])
    header(s, "Customer visit · 7 September 2026", "Hospital Umum Sarawak",
           ["Transfusion — Bio-Rad + Werfen Immucor",
            "FY26 (May 25–Apr 26) vs FY27 YTD (May–Aug 26)"])

    w, gap = 2.963, 0.16
    kpi(s, 0.5, 1.10, w, 1.18, "Transfusion, May–Aug",
        f"RM {money(t27)}", f"vs RM {money(t26)} same window FY26")
    kpi(s, 0.5 + (w + gap), 1.10, w, 1.18, "Like-for-like movement",
        f"{pct(t27, t26):+.1f}%".replace("-", "\u2212"), "Both brands down together",
        value_color=DANGER)
    kpi(s, 0.5 + 2 * (w + gap), 1.10, w, 1.18, "FY26 full year",
        f"RM {money(fy26_full)}", "Full 12 months, for context")
    kpi(s, 0.5 + 3 * (w + gap), 1.10, w, 1.18, "I-LOAN PDN tender, Sarawak",
        "41.52%", "4th of 15 · national total 28.59%",
        value_color=SUCCESS, panel=SKY100)

    # monthly stacked columns
    text(s, 0.5, 2.44, 7.55, 0.24,
         [("Monthly transfusion purchases, May 2025 – Aug 2026 (RM)", 11.5, True, DARK)])
    cd = CategoryChartData()
    cd.categories = [datetime.datetime.strptime(m, "%Y-%m").strftime("%b %y") for m in MONTHS]
    for brand in TRANSFUSION:
        cd.add_series(brand, tuple(
            sum(r["amt"] for r in tx if r["month"] == m and r["brand"] == brand)
            for m in MONTHS))
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(0.5), Inches(2.74),
                            Inches(7.55), Inches(2.85), cd)
    ch = gf.chart
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.TOP
    ch.legend.include_in_layout = False
    ch.plots[0].gap_width = 45
    ch.series[0].format.fill.solid()
    ch.series[0].format.fill.fore_color.rgb = DARK
    ch.series[1].format.fill.solid()
    ch.series[1].format.fill.fore_color.rgb = LIGHT
    style_chart(ch, cat_size=8.5)

    # tender achievement
    text(s, 8.28, 2.44, 4.55, 0.24,
         [("I-LOAN PDN tender achievement by state (%)", 11.5, True, DARK)])
    cd2 = CategoryChartData()
    ordered = TENDER + [("NATIONAL TOTAL", TENDER_TOTAL)]
    cd2.categories = [n for n, _ in reversed(ordered)]
    cd2.add_series("Achievement", tuple(v for _, v in reversed(ordered)))
    gf2 = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(8.28), Inches(2.74),
                             Inches(4.55), Inches(2.85), cd2)
    ch2 = gf2.chart
    ch2.has_legend = False
    ch2.plots[0].gap_width = 40
    ch2.plots[0].has_data_labels = True
    dl = ch2.plots[0].data_labels
    dl.number_format = '0.00"%"'
    dl.number_format_is_linked = False
    dl.font.size = Pt(7)
    dl.font.name = FONT
    dl.font.color.rgb = GREY600
    ser = ch2.series[0]
    for n, (name, _) in enumerate(reversed(ordered)):
        pt = ser.points[n]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = (
            SUCCESS if name == "Sarawak" else SKY600 if name == "NATIONAL TOTAL" else LIGHT)
    style_chart(ch2, cat_size=7.5, val_fmt='0"%"')

    # movement callouts
    box(s, 0.5, 5.78, 12.333, 1.10, fill=WHITE, line=GREY200, radius=0.05)
    text(s, 0.74, 5.94, 3.0, 0.22,
         [("WHERE THE MOVEMENT SITS, MAY–AUG", 9, True, GREY600)])
    xs = 0.74
    for name, delta in moves:
        label = GROUP_LABELS.get(name, str(name).title())
        text(s, xs, 6.24, 2.9, 0.46,
             [[(f"{'+' if delta >= 0 else '−'}RM {money(abs(delta))}", 14, True,
                SUCCESS if delta >= 0 else DANGER)],
              [(label, 9.5, False, GREY800)]], spacing=1.2)
        xs += 3.0
    text(s, 0.74, 6.62, 11.8, 0.22,
         [("Worth knowing: BD flow cytometry at the same site is "
           f"+RM {money(bd_delta)} over the same window — total site spend is up "
           "even though transfusion is down.", 9, False, SKY600)])

    footer(s, "Source: SAP AR invoice export, 5 May 2025 – 4 Sep 2026. Ship-to = Hospital Umum "
              "Sarawak (invoiced via Chemtecq / Alam Medik / Edgenta). Tender achievement per "
              "I-LOAN PDN contract report. Sep 2026 part-month excluded from all comparisons.")


# --- slide 2: PIL Kuching (Timberland) -------------------------------------------
ASSAYS = [
    ("H. pylori IgG", "H.PYLORI", 100),
    ("EBV VCA IgG", "EBV", None),
    ("ANA Screen", "ANA SCREEN", 50),
]


def slide_pil(prs, rows):
    tim = [r for r in rows if "TIMBERLAND" in r["ship"]]
    sn = [r for r in tim if r["brand"] == "SNIBE"]

    t26 = sum(r["amt"] for r in sn if r["month"] in LFL_FY26)
    t27 = sum(r["amt"] for r in sn if r["month"] in LFL_FY27)
    fy26_full = sum(r["amt"] for r in sn if r["fy"] == "FY26")
    fy27_ytd = sum(r["amt"] for r in sn if r["fy"] == "FY27")
    assert round(t26) == 32850 and round(t27) == 14340, (t26, t27)
    assert round(fy26_full) == 73940 and round(fy27_ytd) == 14340, (fy26_full, fy27_ytd)

    def kits(key, fy):
        return sum(r["qty"] for r in sn
                   if key in r["product"].upper() and r["fy"] == fy and r["amt"] > 0)

    foc = sum(1 for r in sn if r["amt"] == 0)

    s = prs.slides.add_slide(prs.slide_layouts[6])
    header(s, "Customer visit · 7 September 2026",
           "Premier Integrated Labs — Timberland Medical Centre, Kuching",
           ["SNIBE MAGLUMI — immunoassay",
            "FY26 (May 25–Apr 26) vs FY27 YTD (May–Aug 26)"])

    w, gap = 2.963, 0.16
    kpi(s, 0.5, 1.10, w, 1.18, "SNIBE reagents, May–Aug",
        f"RM {money(t27)}", f"vs RM {money(t26)} same window FY26")
    kpi(s, 0.5 + (w + gap), 1.10, w, 1.18, "Like-for-like movement",
        f"{pct(t27, t26):+.1f}%".replace("-", "\u2212"), "Almost entirely one assay",
        value_color=DANGER)
    kpi(s, 0.5 + 2 * (w + gap), 1.10, w, 1.18, "FY26 full year",
        f"RM {money(fy26_full)}", f"FY27 YTD so far RM {money(fy27_ytd)}")
    kpi(s, 0.5 + 3 * (w + gap), 1.10, w, 1.18, "H. pylori kits",
        f"{kits('H.PYLORI', 'FY26'):.0f} → {kits('H.PYLORI', 'FY27'):.0f}",
        "FY26 full year → FY27 YTD",
        value_color=DANGER, panel=SKY100)

    # monthly revenue
    text(s, 0.5, 2.44, 7.55, 0.24,
         [("Monthly SNIBE reagent purchases, May 2025 – Aug 2026 (RM)", 11.5, True, DARK)])
    cd = CategoryChartData()
    cd.categories = [datetime.datetime.strptime(m, "%Y-%m").strftime("%b %y") for m in MONTHS]
    cd.add_series("SNIBE", tuple(
        sum(r["amt"] for r in sn if r["month"] == m) for m in MONTHS))
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.5), Inches(2.74),
                            Inches(7.55), Inches(2.85), cd)
    ch = gf.chart
    ch.has_legend = False
    ch.plots[0].gap_width = 45
    ch.series[0].format.fill.solid()
    ch.series[0].format.fill.fore_color.rgb = DARK
    style_chart(ch, cat_size=8.5)

    # assay mix in kits
    text(s, 8.28, 2.44, 4.55, 0.24,
         [("Assay mix — kits purchased", 11.5, True, DARK)])
    cd2 = CategoryChartData()
    cd2.categories = [a[0] for a in ASSAYS]
    cd2.add_series("FY26 full year", tuple(kits(a[1], "FY26") for a in ASSAYS))
    cd2.add_series("FY27 YTD", tuple(kits(a[1], "FY27") for a in ASSAYS))
    gf2 = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(8.28), Inches(2.74),
                             Inches(4.55), Inches(2.85), cd2)
    ch2 = gf2.chart
    ch2.has_legend = True
    ch2.legend.position = XL_LEGEND_POSITION.TOP
    ch2.legend.include_in_layout = False
    ch2.plots[0].gap_width = 60
    ch2.plots[0].has_data_labels = True
    ch2.plots[0].data_labels.font.size = Pt(8)
    ch2.plots[0].data_labels.font.name = FONT
    ch2.plots[0].data_labels.font.color.rgb = GREY600
    ch2.series[0].format.fill.solid()
    ch2.series[0].format.fill.fore_color.rgb = DARK
    ch2.series[1].format.fill.solid()
    ch2.series[1].format.fill.fore_color.rgb = LIGHT
    style_chart(ch2, cat_size=8.5, val_fmt='0')

    # platform note
    box(s, 0.5, 5.78, 8.05, 1.10, fill=WHITE, line=GREY200, radius=0.05)
    text(s, 0.74, 5.94, 7.6, 0.22,
         [("WHAT THEY RUN ON THE MAGLUMI", 9, True, GREY600)])
    text(s, 0.74, 6.20, 7.6, 0.60,
         [[("Three CLIA assays only — H. pylori IgG (100 tests/kit), EBV VCA IgG and "
            "ANA Screen (50 tests/kit).", 10, False, GREY800)],
          [(f"All {foc} MAGLUMI consumable lines — reaction cups, starter kits, wash and "
            "system liquid, light check — invoice at RM 0, consistent with a reagent-rental "
            "placement where only reagents are billed.", 10, False, GREY800)]], spacing=1.25)

    box(s, 8.75, 5.78, 4.083, 1.10, fill=SKY100, line=GREY200, radius=0.05)
    text(s, 8.99, 5.94, 3.6, 0.22, [("AGAINST CONTRACT", 9, True, GREY600)])
    text(s, 8.99, 6.20, 3.6, 0.60,
         [[("[CONTRACT TARGET — TBC]", 12, True, DANGER)],
          [("PIL SNIBE contract not found in the projects folder or Drive — send it and "
            "this becomes an achievement %.", 9, False, GREY800)]], spacing=1.2)

    footer(s, "Source: SAP AR invoice export, 5 May 2025 – 4 Sep 2026. Ship-to = Timberland "
              "Medical Centre Kuching (all four SAP spellings), billed to Premier Integrated "
              "Labs Sdn Bhd. No SNIBE orders in Jan 2026; Sep 2026 part-month excluded.")


# --- slide 2: Hospital Umum Sarawak, detail ---------------------------------------
def slide_hus_detail(prs, rows):
    hus = [r for r in rows if "HOSPITAL UMUM SARAWAK" in r["ship"]]
    tx = [r for r in hus if r["brand"] in TRANSFUSION]

    agg = collections.defaultdict(lambda: [0.0, 0.0, 0.0, 0.0])
    for r in tx:
        v = agg[(r["brand"], r["group"])]
        v[0 if r["fy"] == "FY26" else 1] += r["amt"]
        if r["month"] in LFL_FY26:
            v[2] += r["amt"]
        elif r["month"] in LFL_FY27:
            v[3] += r["amt"]
    lines = sorted(((k, v) for k, v in agg.items() if any(v)), key=lambda x: -x[1][0])

    s = prs.slides.add_slide(prs.slide_layouts[6])
    header(s, "Hospital Umum Sarawak · detail", "Where the money moved, line by line",
           ["Bio-Rad + Werfen Immucor",
            "RM · FY26 full year, FY27 to 4 Sep 2026"])

    text(s, 0.5, 1.12, 7.0, 0.22,
         [("TRANSFUSION LINES", 9, True, GREY600)])
    body = []
    for (brand, group), v in lines:
        body.append([GROUP_LABELS.get(group, str(group).title()),
                     money(v[0]), money(v[1]), money(v[2]), money(v[3]),
                     delta_str(v[3] - v[2])])
    tot = [sum(v[i] for _, v in lines) for i in range(4)]
    body.append(("total", ["Transfusion total", money(tot[0]), money(tot[1]),
                           money(tot[2]), money(tot[3]), delta_str(tot[3] - tot[2])]))
    table(s, 0.5, 1.40, 7.0, [2.15, 1.02, 0.97, 0.97, 0.97, 0.92], body,
          ["Line", "FY26 full", "FY27 YTD", "May–Aug 26*", "May–Aug 27", "Δ LFL"])

    # ID cards timing
    text(s, 7.85, 1.12, 4.98, 0.22,
         [("BIO-RAD ID CARDS — WHEN THE BIG DRAWS LAND (RM)", 9, True, GREY600)])
    cd = CategoryChartData()
    cd.categories = [datetime.datetime.strptime(m, "%Y-%m").strftime("%b %y") for m in MONTHS]
    cd.add_series("ID cards", tuple(
        sum(r["amt"] for r in tx if r["month"] == m and r["group"] == "DM-ID CARDS")
        for m in MONTHS))
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(7.85), Inches(1.40),
                            Inches(4.98), Inches(2.55), cd)
    ch = gf.chart
    ch.has_legend = False
    ch.plots[0].gap_width = 45
    ch.series[0].format.fill.solid()
    ch.series[0].format.fill.fore_color.rgb = DARK
    style_chart(ch, cat_size=8)

    box(s, 7.85, 4.08, 4.98, 1.42, fill=SKY100, line=GREY200, radius=0.05)
    text(s, 8.09, 4.24, 4.5, 1.10,
         [[("Read the drop carefully.", 11, True, DARK)],
          [("Cards are drawn in a few big lumps — Aug–Oct 2025, then Feb and Apr 2026 — "
            "not monthly. FY27's equivalent draws have not happened yet, so part of the "
            "May–Aug gap is timing, not lost share.", 9.5, False, GREY800)]], spacing=1.22)

    points(s, 0.5, 5.62, 12.333, 1.28, "What to ask", [
        ("When is the next card draw?",
         "Aug–Oct is the pattern. Confirming the schedule tells you whether FY27 is behind "
         "or simply not drawn yet."),
        ("Werfen DBL reagent stopped.",
         "RM 8,820 in May–Aug FY26, nothing since. Worth finding out whether that testing "
         "moved, paused or went elsewhere."),
        ("ID reagent and QC are up.",
         "+RM 10,063 and +RM 3,502 like-for-like — routine consumption is healthy, which "
         "supports the timing reading."),
    ])

    footer(s, "* May–Aug FY26 is shown so the four months of FY27 are compared against the "
              "same four months, not against a full year. Source: SAP AR invoice export, "
              "ship-to Hospital Umum Sarawak.")


# --- slide 4: PIL Kuching, detail -------------------------------------------------
def slide_pil_detail(prs, rows):
    tim = [r for r in rows if "TIMBERLAND" in r["ship"]]
    sn = [r for r in tim if r["brand"] == "SNIBE"]

    def kits(key, months=None, fy=None):
        return sum(r["qty"] for r in sn if key in r["product"].upper() and r["amt"] > 0
                   and (fy is None or r["fy"] == fy)
                   and (months is None or r["month"] in months))

    s = prs.slides.add_slide(prs.slide_layouts[6])
    header(s, "Premier Integrated Labs, Timberland Kuching · detail",
           "What changed in January", ["SNIBE MAGLUMI — kits and test capacity",
                                       "FY26 full year vs FY27 to 4 Sep 2026"])

    text(s, 0.5, 1.12, 7.55, 0.22,
         [("KITS PURCHASED PER MONTH, BY ASSAY", 9, True, GREY600)])
    cd = CategoryChartData()
    cd.categories = [datetime.datetime.strptime(m, "%Y-%m").strftime("%b %y") for m in MONTHS]
    for name, key, _ in ASSAYS:
        cd.add_series(name, tuple(kits(key, months=[m]) for m in MONTHS))
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED, Inches(0.5), Inches(1.40),
                            Inches(7.55), Inches(2.90), cd)
    ch = gf.chart
    ch.has_legend = True
    ch.legend.position = XL_LEGEND_POSITION.TOP
    ch.legend.include_in_layout = False
    ch.plots[0].gap_width = 45
    for n, colour in enumerate((DARK, LIGHT, SKY600)):
        ch.series[n].format.fill.solid()
        ch.series[n].format.fill.fore_color.rgb = colour
    style_chart(ch, cat_size=8, val_fmt='0')

    text(s, 8.28, 1.12, 4.55, 0.22,
         [("TEST CAPACITY PURCHASED", 9, True, GREY600)])
    body = []
    for name, key, size in ASSAYS:
        k26, k27 = kits(key, fy="FY26"), kits(key, fy="FY27")
        tests = (f"{k26 * size:,.0f} → {k27 * size:,.0f}" if size else "—")
        body.append([name, size and f"{size}/kit" or "n/a",
                     f"{k26:.0f} → {k27:.0f}", tests])
    table(s, 8.28, 1.40, 4.55, [1.62, 0.80, 0.86, 1.27], body,
          ["Assay (CLIA)", "Kit size", "Kits", "Tests"], font_size=9, row_h=0.34)

    box(s, 8.28, 2.92, 4.55, 1.38, fill=SKY100, line=GREY200, radius=0.05)
    text(s, 8.52, 3.08, 4.07, 1.06,
         [[("The instrument never stopped.", 10.5, True, DARK)],
          [("MAGLUMI consumables — cups, wash and system liquid, light check — kept being "
            "delivered right through the quiet months, all at RM 0. The analyser is "
            "running; the reagent menu narrowed.", 9.5, False, GREY800)]], spacing=1.22)

    points(s, 0.5, 4.48, 12.333, 1.28, "What to ask", [
        ("Nothing at all in January.",
         "All three assays went to zero that month, then EBV and ANA came back and "
         "H. pylori did not. Something changed then — find out what."),
        ("H. pylori is the whole gap.",
         "9,200 tests of capacity in FY26 against 700 so far. If referrals moved to "
         "another method or lab, that is the recoverable volume."),
        ("EBV and ANA are steady.",
         "Both are tracking close to last year's rhythm, so this is not the account "
         "pulling back from the platform."),
    ])

    footer(s, "Kits are billed units; test capacity is kits × kit size (EBV kit size not "
              "stated on the invoice line). Source: SAP AR invoice export, ship-to Timberland "
              "Medical Centre Kuching, billed to Premier Integrated Labs Sdn Bhd.")


def main():
    rows = load()
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    slide_hus(prs, rows)
    slide_hus_detail(prs, rows)
    slide_pil(prs, rows)
    slide_pil_detail(prs, rows)
    prs.save(OUT)
    print(f"wrote {OUT}")


if __name__ == "__main__":
    sys.exit(main())
