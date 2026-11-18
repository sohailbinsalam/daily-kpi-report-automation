"""Builds the daily PowerPoint report with python-pptx. Charts are native PowerPoint charts, so they stay editable."""
from __future__ import annotations

import math
from datetime import datetime
from pathlib import Path

import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_MARKER_STYLE
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

from .config import Config
from .etl import EtlResult

NAVY = RGBColor(0x0F, 0x2A, 0x4A)
TEAL = RGBColor(0x0E, 0x9F, 0x8E)
INK = RGBColor(0x1B, 0x24, 0x30)
MUTED = RGBColor(0x66, 0x72, 0x82)
LINE = RGBColor(0xD9, 0xE0, 0xE8)
CARD = RGBColor(0xF5, 0xF7, 0xFA)
GOOD = RGBColor(0x15, 0x80, 0x3D)
BAD = RGBColor(0xC0, 0x26, 0x3D)
GOOD_BG = RGBColor(0xE3, 0xF5, 0xE9)
BAD_BG = RGBColor(0xFB, 0xE4, 0xE8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FONT = "Segoe UI"

W, H = Inches(13.333), Inches(7.5)
TREND_KPIS = ["Uptime_pct", "Error_Rate_pct", "Response_Time_ms", "Speed_Mbps"]


# --------------------------------------------------------------------------- helpers

def _text(slide, x, y, w, h, text, size=12, bold=False, color=INK, align=PP_ALIGN.LEFT,
          anchor=MSO_ANCHOR.TOP):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.05)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    tf.vertical_anchor = anchor
    lines = text if isinstance(text, list) else [text]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        r = p.add_run()
        r.text = line
        r.font.size, r.font.bold, r.font.name = Pt(size), bold, FONT
        r.font.color.rgb = color
    return box


def _rect(slide, x, y, w, h, fill, line=None, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, x, y, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    if line is None:
        s.line.fill.background()
    else:
        s.line.color.rgb = line
        s.line.width = Pt(0.75)
    s.shadow.inherit = False
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = 0.08
    return s


def _pill(slide, x, y, text, good: bool | None):
    fill, color = (GOOD_BG, GOOD) if good else (BAD_BG, BAD)
    if good is None:
        fill, color = (CARD, MUTED)
    p = _rect(slide, x, y, Inches(1.05), Inches(0.28), fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    p.adjustments[0] = 0.5
    tf = p.text_frame
    tf.margin_top = tf.margin_bottom = 0
    r = tf.paragraphs[0].add_run()
    r.text = text
    r.font.size, r.font.bold, r.font.name = Pt(9), True, FONT
    r.font.color.rgb = color
    tf.paragraphs[0].alignment = PP_ALIGN.CENTER
    return p


def _header(slide, title: str, subtitle: str, cfg: Config, page: int):
    _rect(slide, 0, 0, W, Inches(0.95), NAVY)
    _rect(slide, 0, Inches(0.95), W, Inches(0.05), TEAL)
    _text(slide, Inches(0.5), Inches(0.14), Inches(9), Inches(0.45), title, 22, True, WHITE)
    _text(slide, Inches(0.5), Inches(0.55), Inches(9), Inches(0.3), subtitle, 11, False, RGBColor(0xC7, 0xD6, 0xE8))
    _text(slide, Inches(0.5), H - Inches(0.42), Inches(8), Inches(0.3), cfg.report.get("footer", ""), 9, False, MUTED)
    _text(slide, W - Inches(1.5), H - Inches(0.42), Inches(1), Inches(0.3), str(page), 9, False, MUTED, PP_ALIGN.RIGHT)


def nice_axis(lo: float, hi: float, ticks: int = 5) -> tuple[float, float, float]:
    """Round axis limits to a clean step (1, 2, 2.5 or 5 x 10^n) so tick labels read 99.5, 99.6 ..."""
    span = (hi - lo) or abs(hi) * 0.02 or 1.0
    raw = span / ticks
    mag = 10 ** math.floor(math.log10(raw))
    step = next(m * mag for m in (1, 2, 2.5, 5, 10) if m * mag >= raw)
    lo_n = math.floor(lo / step) * step
    hi_n = math.ceil(hi / step) * step
    if hi_n - hi < step * 0.2:
        hi_n += step
    if lo - lo_n < step * 0.2:
        lo_n -= step
    return round(lo_n, 6), round(hi_n, 6), step


def _value_txt(k, v) -> str:
    return f"{k.fmt(v)} {k.unit}".strip()


# --------------------------------------------------------------------------- slides

def _slide_title(prs, res: EtlResult, cfg: Config):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _rect(s, 0, 0, W, H, NAVY)
    _rect(s, Inches(0.8), Inches(2.55), Inches(1.2), Inches(0.08), TEAL)
    _text(s, Inches(0.8), Inches(2.8), Inches(11), Inches(0.9), cfg.report["title"], 40, True, WHITE)
    _text(s, Inches(0.8), Inches(3.7), Inches(11), Inches(0.5),
          f"{cfg.report['platform_name']}  ·  {res.day:%A %d %B %Y}", 18, False, RGBColor(0xC7, 0xD6, 0xE8))
    on = sum(1 for r in res.status if r["on_target"])
    total = sum(1 for r in res.status if r["on_target"] is not None)
    _text(s, Inches(0.8), Inches(4.5), Inches(11), Inches(0.5),
          f"{on} of {total} KPIs on target  ·  {len(res.breaches)} zone issue(s)", 16, True, TEAL)
    _text(s, Inches(0.8), H - Inches(0.9), Inches(11), Inches(0.4),
          f"Generated automatically {datetime.now():%d %b %Y %H:%M}  ·  {cfg.report.get('footer', '')}",
          10, False, RGBColor(0x9F, 0xB3, 0xC8))


def _slide_summary(prs, res: EtlResult, cfg: Config):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _header(s, "Overall summary", f"Overall KPIs for {res.day:%d %b %Y}, compared with the day before", cfg, 2)
    cols, cw, ch, gx, gy = 3, Inches(3.95), Inches(1.72), Inches(0.22), Inches(0.2)
    x0, y0 = Inches(0.5), Inches(1.3)
    for i, r in enumerate(res.status):
        k = r["kpi"]
        x = x0 + (cw + gx) * (i % cols)
        y = y0 + (ch + gy) * (i // cols)
        _rect(s, x, y, cw, ch, CARD, LINE, MSO_SHAPE.ROUNDED_RECTANGLE)
        _text(s, x + Inches(0.18), y + Inches(0.12), cw - Inches(1.5), Inches(0.3), k.label, 11, True, MUTED)
        if r["on_target"] is not None:
            _pill(s, x + cw - Inches(1.22), y + Inches(0.14), "ON TARGET" if r["on_target"] else "OFF TARGET", r["on_target"])
        val_color = INK if r["on_target"] is None else (GOOD if r["on_target"] else BAD)
        _text(s, x + Inches(0.18), y + Inches(0.48), cw - Inches(0.3), Inches(0.6), _value_txt(k, r["value"]), 26, True, val_color)
        target = f"Target {'≥' if k.better == 'higher' else '≤'} {k.fmt(k.target)} {k.unit}" if k.target is not None else "Volume KPI, no target"
        _text(s, x + Inches(0.18), y + Inches(1.15), Inches(2.2), Inches(0.3), target, 10, False, MUTED)
        if r["improved"] is None:
            ccolor = MUTED
            arrow = ""
        else:
            ccolor = GOOD if r["improved"] else BAD
            arrow = "▲ " if (r["delta"] or 0) >= 0 else "▼ "
        _text(s, x + cw - Inches(1.9), y + Inches(1.15), Inches(1.72), Inches(0.3),
              f"{arrow}{r['change_txt']} vs yesterday", 10, True, ccolor, PP_ALIGN.RIGHT)


def _slide_zones(prs, res: EtlResult, cfg: Config):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _header(s, "KPIs by zone", "Red cells missed the target", cfg, 3)
    kpis = cfg.kpis
    df = res.today_zones.sort_values("Zone")
    n_rows, n_cols = len(df) + 1, len(kpis) + 2
    tbl = s.shapes.add_table(n_rows, n_cols, Inches(0.5), Inches(1.35), W - Inches(1.0),
                             Inches(0.5) * n_rows).table
    widths = [Inches(1.4), Inches(0.9)] + [Emu(int((W - Inches(1.0) - Inches(2.3)) / len(kpis)))] * len(kpis)
    for i, w in enumerate(widths):
        tbl.columns[i].width = w

    def cell(r, c, text, bold=False, color=INK, fill=WHITE, size=11, align=PP_ALIGN.CENTER):
        cl = tbl.cell(r, c)
        cl.fill.solid()
        cl.fill.fore_color.rgb = fill
        cl.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cl.text_frame
        tf.paragraphs[0].alignment = align
        run = tf.paragraphs[0].add_run()
        run.text = text
        run.font.size, run.font.bold, run.font.name = Pt(size), bold, FONT
        run.font.color.rgb = color

    cell(0, 0, "Zone", True, WHITE, NAVY, 10, PP_ALIGN.LEFT)
    cell(0, 1, "Units", True, WHITE, NAVY, 10)
    for j, k in enumerate(kpis, start=2):
        cell(0, j, f"{k.label}\n({k.unit})", True, WHITE, NAVY, 9)
    for i, (_, row) in enumerate(df.iterrows(), start=1):
        cell(i, 0, row["Zone"], True, align=PP_ALIGN.LEFT)
        cell(i, 1, f"{int(row['Unit_Count']):,}", color=MUTED)
        for j, k in enumerate(kpis, start=2):
            v = row[k.column]
            ok = k.meets_target(v) if pd.notna(v) else None
            fill = BAD_BG if ok is False else WHITE
            color = BAD if ok is False else (MUTED if pd.isna(v) else INK)
            cell(i, j, k.fmt(v) if pd.notna(v) else "blank", ok is False, color, fill)


def _slide_trends(prs, res: EtlResult, cfg: Config):
    days = int(cfg.report.get("trend_days", 14))
    net = res.overall.tail(days)
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _header(s, f"Trends, last {len(net)} days", "Overall daily value against the target", cfg, 4)
    by_col = {k.column: k for k in cfg.kpis}
    cats = [d.strftime("%d %b") for d in net["Date"]]
    cw, chh = Inches(6.05), Inches(2.75)
    for i, col in enumerate(c for c in TREND_KPIS if c in by_col):
        k = by_col[col]
        x = Inches(0.5) + (cw + Inches(0.23)) * (i % 2)
        y = Inches(1.25) + (chh + Inches(0.2)) * (i // 2)
        _rect(s, x, y, cw, chh, CARD, LINE, MSO_SHAPE.ROUNDED_RECTANGLE)
        _text(s, x + Inches(0.2), y + Inches(0.1), cw, Inches(0.3), f"{k.label} ({k.unit})", 12, True, INK)
        data = CategoryChartData()
        data.categories = cats
        data.add_series("Actual", [round(v, k.decimals) if pd.notna(v) else None for v in net[col]])
        if k.target is not None:
            data.add_series("Target", [k.target] * len(cats))
        gf = s.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, x + Inches(0.1), y + Inches(0.4),
                                cw - Inches(0.2), chh - Inches(0.45), data)
        ch = gf.chart
        ch.font.size, ch.font.name = Pt(9), FONT
        ch.font.color.rgb = MUTED
        ch.has_legend = True
        ch.legend.position = XL_LEGEND_POSITION.TOP
        ch.legend.include_in_layout = False
        ch.value_axis.has_major_gridlines = True
        ch.value_axis.major_gridlines.format.line.color.rgb = LINE
        ch.value_axis.format.line.fill.background()
        ch.category_axis.format.line.color.rgb = LINE
        ch.category_axis.tick_labels.font.size = Pt(8)
        vals = [v for v in net[col] if pd.notna(v)] + ([k.target] if k.target is not None else [])
        lo, hi, step = nice_axis(min(vals), max(vals))
        ch.value_axis.minimum_scale = lo
        ch.value_axis.maximum_scale = hi
        ch.value_axis.major_unit = step
        dec = max(0, -math.floor(math.log10(step))) if step < 1 else 0
        ch.value_axis.tick_labels.number_format = "0" if dec == 0 else "0." + "0" * dec
        ch.value_axis.tick_labels.number_format_is_linked = False
        actual = ch.plots[0].series[0]
        actual.format.line.color.rgb = TEAL
        actual.format.line.width = Pt(2.25)
        actual.smooth = False
        actual.marker.format.fill.solid()
        actual.marker.format.fill.fore_color.rgb = TEAL
        actual.marker.format.line.color.rgb = TEAL
        actual.marker.size = 5
        if k.target is not None:
            tgt = ch.plots[0].series[1]
            tgt.format.line.color.rgb = BAD
            tgt.format.line.width = Pt(1.25)
            tgt.format.line.dash_style = MSO_LINE_DASH_STYLE.DASH
            tgt.marker.style = XL_MARKER_STYLE.NONE


def _slide_issues(prs, res: EtlResult, cfg: Config):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _header(s, "Issues and data quality", "What needs attention today, and what the ETL fixed in the input file", cfg, 5)

    # left: off-target list
    lx, ly, lw = Inches(0.5), Inches(1.3), Inches(7.6)
    _rect(s, lx, ly, lw, Inches(5.5), CARD, LINE, MSO_SHAPE.ROUNDED_RECTANGLE)
    _text(s, lx + Inches(0.25), ly + Inches(0.15), lw, Inches(0.35), "Zones off target", 14, True, INK)
    if not res.breaches:
        _text(s, lx + Inches(0.25), ly + Inches(0.7), lw - Inches(0.5), Inches(0.4),
              "Every zone met every target today.", 12, False, GOOD)
    for i, b in enumerate(res.breaches[:9]):
        y = ly + Inches(0.65) + Inches(0.52) * i
        k = b["kpi"]
        _rect(s, lx + Inches(0.25), y, Inches(0.07), Inches(0.38), BAD)
        _text(s, lx + Inches(0.42), y - Inches(0.02), Inches(3.6), Inches(0.25), f"{b['zone']}  ·  {k.label}", 11, True, INK)
        _text(s, lx + Inches(0.42), y + Inches(0.19), Inches(4.5), Inches(0.22),
              f"{_value_txt(k, b['value'])} against target {'≥' if k.better == 'higher' else '≤'} {k.fmt(k.target)} {k.unit}",
              9, False, MUTED)
        _text(s, lx + lw - Inches(1.6), y + Inches(0.03), Inches(1.35), Inches(0.3), f"{b['gap']:.1%} off", 11, True, BAD, PP_ALIGN.RIGHT)
    if len(res.breaches) > 9:
        _text(s, lx + Inches(0.25), ly + Inches(5.1), lw, Inches(0.3), f"+ {len(res.breaches) - 9} more in the zone table", 10, False, MUTED)

    # right: data quality
    rx, rw = Inches(8.35), Inches(4.48)
    _rect(s, rx, ly, rw, Inches(5.5), CARD, LINE, MSO_SHAPE.ROUNDED_RECTANGLE)
    _text(s, rx + Inches(0.25), ly + Inches(0.15), rw, Inches(0.35), "Input file checks", 14, True, INK)
    v = res.validation
    checks = [
        ("Rows read", str(v.rows_read), None),
        ("Duplicates removed", str(v.duplicates_removed), v.duplicates_removed == 0),
        ("Blank values", str(len(v.missing_values)), not v.missing_values),
        ("Impossible values blanked", str(len(v.rejected)), not v.rejected),
    ]
    for i, (label, val, good) in enumerate(checks):
        y = ly + Inches(0.7) + Inches(0.62) * i
        _text(s, rx + Inches(0.25), y, Inches(2.6), Inches(0.35), label, 12, False, INK)
        color = INK if good is None else (GOOD if good else BAD)
        _text(s, rx + rw - Inches(1.3), y, Inches(1.05), Inches(0.35), val, 16, True, color, PP_ALIGN.RIGHT)
    details = [f"Row {d['row']} · {d['zone']} · {d['column']}: {d['reason']}" for d in (v.rejected + [
        {**m, "reason": "blank"} for m in v.missing_values])][:5]
    if details:
        _text(s, rx + Inches(0.25), ly + Inches(3.3), rw - Inches(0.5), Inches(1.9), details, 9, False, MUTED)
    _text(s, rx + Inches(0.25), ly + Inches(4.95), rw - Inches(0.5), Inches(0.45),
          f"Source: {res.source.name}", 9, False, MUTED)


# --------------------------------------------------------------------------- entry

def build_report(res: EtlResult, cfg: Config) -> Path:
    prs = Presentation()
    prs.slide_width, prs.slide_height = W, H
    _slide_title(prs, res, cfg)
    _slide_summary(prs, res, cfg)
    _slide_zones(prs, res, cfg)
    _slide_trends(prs, res, cfg)
    _slide_issues(prs, res, cfg)
    out_dir = cfg.paths["reports"]
    out_dir.mkdir(parents=True, exist_ok=True)
    out = out_dir / f"Daily_KPI_Report_{res.day:%Y-%m-%d}.pptx"
    prs.save(out)
    return out
