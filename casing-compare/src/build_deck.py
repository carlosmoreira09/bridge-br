"""Gera o deck .pptx comparando SCORE-PySafe x WellCat para todos os
carregamentos registrados em loadcases.yaml.

Uso:
    python src/build_deck.py                       # todos os loadcases
    python src/build_deck.py --only prod_surg_pico_oleo
    python src/build_deck.py --out out/meu_deck.pptx
"""
from __future__ import annotations

import argparse
import datetime as dt
import math
from pathlib import Path

import numpy as np
import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_MARKER_STYLE, XL_TICK_LABEL_POSITION
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from extract import ROOT, compare, export_filled_xlsx, extract_loadcase, get_arch, get_pressure_map, load_config

# ------------------------------------------------------------------ estilo
FONT = "Calibri"
HEAD = "Cambria"
NAVY = RGBColor(0x0E, 0x24, 0x33)
INK = RGBColor(0x1B, 0x2B, 0x36)
MUTED = RGBColor(0x5B, 0x6B, 0x76)
PANEL = RGBColor(0xF1, 0xF4, 0xF6)
GRID = RGBColor(0xDD, 0xE3, 0xE8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ICE = RGBColor(0xB9, 0xD3, 0xDE)
SW = {"PySafe": RGBColor(0x0F, 0x7C, 0x8C), "WellCat": RGBColor(0xD9, 0x82, 0x2B)}
SW_LABEL = {"PySafe": "SCORE-PySafe", "WellCat": "WellCat"}
GEO = RGBColor(0x9A, 0xA5, 0xAE)
EST_DARK = RGBColor(0xB8, 0x6A, 0x1E)
STR_COL = {"22": RGBColor(0x34, 0x49, 0x5E), "14": RGBColor(0x6C, 0x5B, 0x9E),
           "10.75": RGBColor(0x2E, 0x8B, 0x57), "COP": RGBColor(0xB2, 0x3A, 0x48)}
RAW_COL = {"Fluido": RGBColor(0xB2, 0x3A, 0x48), "Anular A": RGBColor(0x6C, 0x5B, 0x9E),
           "Anular B": RGBColor(0x34, 0x49, 0x5E)}
UNIT = {"temp": "°C", "press": "psi"}
QTY = {"temp": "Temperatura", "press": "Pressão interna"}

W, H = Inches(13.333), Inches(7.5)
MX = Inches(0.5)


# ------------------------------------------------------------ utilidades
def fmt(v, kind, sign=False, dec=None):
    d = (1 if kind == "temp" else 0) if dec is None else dec
    s = f"{v:+,.{d}f}" if sign else f"{v:,.{d}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def fm(md):
    return fmt(md, "press")  # MD sem casas decimais, milhar com ponto


def text(slide, x, y, w, h, runs, size=14, color=INK, bold=False, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, font=FONT, name=None):
    tb = slide.shapes.add_textbox(x, y, w, h)
    if name:
        tb.name = name
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    paras = runs if isinstance(runs, list) else [runs]
    for i, p in enumerate(paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = align
        items = p if isinstance(p, list) else [(p, {})] if isinstance(p, str) else [p]
        for t, o in items:
            r = para.add_run()
            r.text = t
            f = r.font
            f.name = o.get("font", font)
            f.size = Pt(o.get("size", size))
            f.bold = o.get("bold", bold)
            f.italic = o.get("italic", False)
            f.color.rgb = o.get("color", color)
        if isinstance(p, list) and p and p[0][1].get("space_after"):
            para.space_after = Pt(p[0][1]["space_after"])
    return tb


def rect(slide, x, y, w, h, fill, shape=MSO_SHAPE.ROUNDED_RECTANGLE, name=None):
    s = slide.shapes.add_shape(shape, x, y, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    s.shadow.inherit = False
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = 0.08
    if name:
        s.name = name
    return s


class Deck:
    def __init__(self, cfg):
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = W, H
        self.blank = self.prs.slide_layouts[6]
        self.cfg = cfg
        self.n = 0

    def slide(self, dark=False, footer=None):
        s = self.prs.slides.add_slide(self.blank)
        self.n += 1
        bg = s.background.fill
        bg.solid()
        bg.fore_color.rgb = NAVY if dark else WHITE
        if footer and not dark:
            text(s, MX, Inches(7.08), Inches(10), Inches(0.25), footer, size=10, color=MUTED,
                 name="Footer")
            text(s, Inches(11.83), Inches(7.08), Inches(1.0), Inches(0.25), str(self.n), size=10,
                 color=MUTED, align=PP_ALIGN.RIGHT, name="SlideNumber")
        return s

    def title(self, s, t, sub=None):
        text(s, MX, Inches(0.35), Inches(12.3), Inches(0.6), t, size=30, bold=True, font=HEAD,
             name="Title")
        if sub:
            text(s, MX, Inches(0.98), Inches(12.3), Inches(0.35), sub, size=15, color=MUTED,
                 name="Subtitle")


# ------------------------------------------------------------- gráficos
def _style_axis(ax, title, size=10):
    ax.has_major_gridlines = True
    ax.major_gridlines.format.line.color.rgb = GRID
    ax.major_gridlines.format.line.width = Pt(0.5)
    ax.format.line.color.rgb = GRID
    ax.tick_labels.font.size = Pt(size)
    ax.tick_labels.font.name = FONT
    ax.tick_labels.font.color.rgb = MUTED
    ax.has_title = True
    ax.axis_title.text_frame.text = title
    r = ax.axis_title.text_frame.paragraphs[0].runs[0]
    r.font.size = Pt(size)
    r.font.bold = False
    r.font.name = FONT
    r.font.color.rgb = MUTED


def depth_chart(slide, x, y, w, h, series, xlab, title=None, xr=None, yr=None, legend=True,
                name=None):
    """series: list of dicts {name, md, val, color, dash?, width?, marker?}.
    Profundidade (MD) no eixo vertical, crescendo para baixo."""
    cd = XyChartData()
    # Todas as séries com o mesmo nº de pontos (repete o último ponto): o LibreOffice
    # descarta séries XY mais curtas que as seguintes; no PowerPoint é inócuo.
    nmax = max(len(list(s["md"])) for s in series)
    for s in series:
        ser = cd.add_series(s["name"])
        md, val = [float(m) for m in s["md"]], [float(v) for v in s["val"]]
        md += [md[-1]] * (nmax - len(md))
        val += [val[-1]] * (nmax - len(val))
        for m, v in zip(md, val):
            ser.add_data_point(v, m)
    gf = slide.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS, x, y, w, h, cd)
    if name:
        gf.name = name
    ch = gf.chart
    ch.font.name = FONT
    ch.font.size = Pt(10)
    ch.has_title = bool(title)
    if title:
        ch.chart_title.text_frame.text = title
        r = ch.chart_title.text_frame.paragraphs[0].runs[0]
        r.font.size = Pt(13)
        r.font.bold = True
        r.font.name = FONT
        r.font.color.rgb = INK
    ch.has_legend = legend
    if legend:
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
        ch.legend.font.size = Pt(10)
        ch.legend.font.color.rgb = MUTED
    xa, ya = ch.category_axis, ch.value_axis
    _style_axis(xa, xlab)
    _style_axis(ya, "MD (m)")
    ya.reverse_order = True
    xa.tick_label_position = XL_TICK_LABEL_POSITION.LOW  # rótulos no topo (eixo invertido)
    ya.tick_label_position = XL_TICK_LABEL_POSITION.LOW  # MD sempre à esquerda (mesmo com Δ<0)
    if xr:
        xa.minimum_scale, xa.maximum_scale, xa.major_unit = xr
    if yr:
        ya.minimum_scale, ya.maximum_scale, ya.major_unit = yr
    ya.tick_labels.number_format = "#,##0"
    ya.tick_labels.number_format_is_linked = False
    xa.tick_labels.number_format = "#,##0"
    xa.tick_labels.number_format_is_linked = False
    for plot in ch.plots:
        for ser, s in zip(plot.series, series):
            ser.smooth = False
            ln = ser.format.line
            ln.color.rgb = s["color"]
            ln.width = Pt(s.get("width", 2.0))
            if s.get("dash"):
                ln.dash_style = s["dash"]
            mk = s.get("marker")
            if mk:
                ser.marker.style = XL_MARKER_STYLE.CIRCLE
                ser.marker.size = mk
                ser.marker.format.fill.solid()
                ser.marker.format.fill.fore_color.rgb = s["color"]
                ser.marker.format.line.color.rgb = s["color"]
            else:
                ser.marker.style = XL_MARKER_STYLE.NONE
    return gf


def nice_step(span, target=8):
    raw = span / target
    mag = 10 ** math.floor(math.log10(raw))
    for m in (1, 2, 2.5, 5, 10):
        if raw <= m * mag:
            return m * mag
    return 10 * mag


def nice_range(vals, pad=0.04, step=None, target=8):
    """Retorna (min, max, passo) arredondados."""
    lo, hi = float(np.nanmin(vals)), float(np.nanmax(vals))
    span = hi - lo or 1.0
    lo, hi = lo - span * pad, hi + span * pad
    step = step or nice_step(hi - lo, target)
    return math.floor(lo / step) * step, math.ceil(hi / step) * step, step


# ------------------------------------------------------------- slides
def kpi(slide, x, y, w, h, value, label, color=INK, name=None):
    rect(slide, x, y, w, h, PANEL, name=name)
    text(slide, x + Inches(0.18), y + Inches(0.1), w - Inches(0.36), Inches(0.45), value,
         size=22, bold=True, color=color)
    text(slide, x + Inches(0.18), y + Inches(0.58), w - Inches(0.36), h - Inches(0.62), label,
         size=11, color=MUTED)


def reading(r, kind):
    u = UNIT[kind]
    if r["d_min"] > 0:
        return f"WellCat acima do PySafe em toda a faixa comparada (Δ {fmt(r['d_min'], kind, True)} a {fmt(r['d_max'], kind, True)} {u})"
    if r["d_max"] < 0:
        return f"WellCat abaixo do PySafe em toda a faixa comparada (Δ {fmt(r['d_min'], kind, True)} a {fmt(r['d_max'], kind, True)} {u})"
    return f"Sinal do Δ alterna ao longo do poço (Δ {fmt(r['d_min'], kind, True)} a {fmt(r['d_max'], kind, True)} {u})"


def coverage_note(r):
    ps = r["pysafe_md"][1] - r["pysafe_md"][0]
    ov = r["overlap"][1] - r["overlap"][0]
    pct = 100 * ov / ps if ps else 100
    return pct, f"Faixa comparada {fm(r['overlap'][0])}–{fm(r['overlap'][1])} m ({pct:.0f}% do trecho PySafe)"


EST = RGBColor(0xF2, 0xB8, 0x80)  # WellCat estimado (preenchimento)


def est_segments(wc):
    """Trechos contíguos 'estimado', cada um com o ponto vizinho para dar continuidade."""
    wc = wc.sort_values("md", kind="stable").reset_index(drop=True)
    flag = (wc.src == "estimado").values
    segs, i = [], 0
    while i < len(flag):
        if flag[i]:
            j = i
            while j + 1 < len(flag) and flag[j + 1]:
                j += 1
            segs.append(wc.iloc[max(i - 1, 0): min(j + 2, len(wc))])
            i = j + 1
        else:
            i += 1
    return segs


def fill_delta(tidy, key):
    """Δ ponto a ponto no grid preenchido, com a origem do valor WellCat."""
    F = tidy[(tidy.kind == "temp_fill") & (tidy.string == key)]
    pv = F.pivot_table(index="md", columns="software", values="value", aggfunc="last")
    src = F[F.software == "WellCat"].groupby("md").src.last()
    d = (pv["WellCat"] - pv["PySafe"]).dropna()
    return d.index.values, d.values, src.reindex(d.index).values


def pct_measured(r_meas):
    if not r_meas:
        return 0.0
    ps = r_meas["pysafe_md"][1] - r_meas["pysafe_md"][0]
    ov = r_meas["overlap"][1] - r_meas["overlap"][0]
    return 100 * ov / ps if ps else 100.0


def string_slide(deck, lc, tidy, st, kind, foot, results):
    key, label = st["key"], st["label"]
    r = results.get((kind, key))
    rm = results.get((kind + "_meas", key), r) if kind == "temp" else r
    if not r:
        return None
    s = deck.slide(footer=foot)
    u = UNIT[kind]
    pct = pct_measured(rm)
    sub = f"Δ máx (WellCat − PySafe) = {fmt(r['d_absmax'], kind, True)} {u} a {fm(r['d_absmax_md'])} m MD"
    if kind == "temp":
        sub += (f" · WellCat medido em {pct:.0f}% do trecho ({fm(rm['wellcat_md'][0])}–{fm(rm['wellcat_md'][1])} m)"
                if pct < 99 else " · WellCat medido em todo o trecho")
    deck.title(s, f"{QTY[kind]} – {label}", sub)
    dk = "temp_fill" if kind == "temp" else kind
    sel = tidy[(tidy.kind == dk) & (tidy.string == key)]
    ps = sel[sel.software == "PySafe"]
    if kind == "temp":
        ps = tidy[(tidy.kind == "temp") & (tidy.software == "PySafe") & (tidy.string == key)]
        wc_all = sel[sel.software == "WellCat"]
        wc = tidy[(tidy.kind == "temp") & (tidy.software == "WellCat") & (tidy.string == key)]
        segs = est_segments(wc_all)
    else:
        wc, segs = sel[sel.software == "WellCat"], []
    geo = tidy[(tidy.kind == "temp") & (tidy.software == "PySafe") & (tidy.string == "GEO")]
    yr = nice_range(sel.md, pad=0.02, target=10)
    geo = geo[(geo.md >= yr[0]) & (geo.md <= yr[1])]
    xr = nice_range(sel.value, target=6)
    if kind == "temp":
        xr = nice_range(np.r_[sel.value.values, geo.value.values], target=6)
    xlab = f"{QTY[kind]} ({u})"
    cw, ch, cy = Inches(4.0), Inches(4.35), Inches(1.45)
    xs = [MX, MX + Inches(4.15), MX + Inches(8.3)]
    geo_s = ([dict(name="Geotérmico", md=geo.md, val=geo.value, color=GEO, width=1.25,
                   dash=MSO_LINE_DASH_STYLE.DASH)] if kind == "temp" else [])
    wc_mark = 5 if len(wc) < 40 else None
    est_s = [dict(name="WellCat estimado", md=g.md, val=g.value, color=EST, width=2.25,
                  dash=MSO_LINE_DASH_STYLE.ROUND_DOT) for g in segs]
    depth_chart(s, xs[0], cy, cw, ch,
                [dict(name="SCORE-PySafe", md=ps.md, val=ps.value, color=SW["PySafe"])] + geo_s,
                xlab, "SCORE-PySafe", xr, yr, name="Chart PySafe")
    depth_chart(s, xs[1], cy, cw, ch,
                [dict(name="WellCat", md=wc.md, val=wc.value, color=SW["WellCat"], marker=wc_mark)]
                + est_s + geo_s, xlab, "WellCat", xr, yr, name="Chart WellCat")
    depth_chart(s, xs[2], cy, cw, ch,
                [dict(name="SCORE-PySafe", md=ps.md, val=ps.value, color=SW["PySafe"]),
                 dict(name="WellCat", md=wc.md, val=wc.value, color=SW["WellCat"],
                      dash=MSO_LINE_DASH_STYLE.DASH, marker=wc_mark)] + est_s,
                xlab, "Sobreposição", xr, yr, name="Chart Overlay")
    ky, kh, kw, gap = Inches(5.95), Inches(0.98), Inches(2.93), Inches(0.19)
    cards = [
        (f"{fmt(r['pysafe_max'], kind)} {u}", f"Máx. PySafe @ {fm(r['pysafe_md_at_max'])} m", SW["PySafe"]),
        (f"{fmt(rm['wellcat_max'], kind)} {u}", f"Máx. WellCat medido @ {fm(rm['wellcat_md_at_max'])} m", SW["WellCat"]),
    ]
    if kind == "temp" and pct < 99:
        cards += [
            (f"{fmt(rm['d_mean'], kind, True)} {u}", "Δ médio só no trecho medido", INK),
            (f"{fmt(r['d_mean'], kind, True)} {u}", "Δ médio no trecho completo (com estimado)", EST_DARK),
        ]
    else:
        cards += [
            (f"{fmt(r['d_mean'], kind, True)} {u}", f"Δ médio em {r['n']} pontos", INK),
            (f"{pct:.0f}%", "do trecho com dado WellCat original", INK),
        ]
    for i, (v, l, c) in enumerate(cards):
        kpi(s, MX + i * (kw + gap), ky, kw, kh, v, l, c, name=f"KPI {i+1}")
    note = (f"{reading(r, kind)}. WellCat medido: {fm(rm['wellcat_md'][0])}–{fm(rm['wellcat_md'][1])} m "
            f"({pct:.0f}% do trecho PySafe {fm(r['pysafe_md'][0])}–{fm(r['pysafe_md'][1])} m). ")
    if kind == "temp":
        note += ("Lacunas internas do WellCat preenchidas por interpolação linear entre pontos medidos; "
                 "trechos fora da faixa medida (linha pontilhada) estimados conforme temp_fill no loadcases.yaml. "
                 f"Δ médio só no trecho medido: {fmt(rm['d_mean'], kind, True)} {u}.")
    else:
        note += "Δ calculado nas profundidades do WellCat, com PySafe interpolado linearmente."
    s.notes_slide.notes_text_frame.text = note
    return r


def overview_slide(deck, lc, tidy, kind, foot, strings):
    s = deck.slide(footer=foot)
    u = UNIT[kind]
    if kind == "temp":
        deck.title(s, "Temperatura – visão geral dos revestimentos",
                   f"Todos os revestimentos com o gradiente geotérmico de referência · {lc['name']}")
    else:
        deck.title(s, "Pressão – visão geral",
                   f"PySafe: curvas contínuas (Fluido, Anular A, Anular B) · WellCat: pressão interna por revestimento · {lc['name']}")
    sel = tidy[tidy.kind == kind]
    if kind == "temp":
        sel = tidy[tidy.kind.isin(["temp", "temp_fill"]) & (tidy.string != "GEO")]
    geo = tidy[(tidy.kind == "temp") & (tidy.software == "PySafe") & (tidy.string == "GEO")]
    all_md = np.r_[sel.md.values]
    yr = nice_range(all_md, pad=0.02, step=500)
    geo = geo[(geo.md >= yr[0]) & (geo.md <= yr[1])]
    cw, ch, cy = Inches(6.05), Inches(5.45), Inches(1.45)
    for i, sw in enumerate(["PySafe", "WellCat"]):
        ser = []
        if kind == "press" and sw == "PySafe":
            raw = tidy[(tidy.kind == "press_raw")]
            for c in ["Fluido", "Anular A", "Anular B"]:
                g = raw[raw.string == c]
                ser.append(dict(name=c, md=g.md, val=g.value, color=RAW_COL[c]))
        else:
            for st in strings:
                g = sel[(sel.kind != "temp_fill") & (sel.software == sw) & (sel.string == st["key"])]
                if len(g):
                    ser.append(dict(name=st["short"],
                                    md=g.md, val=g.value, color=STR_COL[st["key"]],
                                    marker=4 if len(g) < 40 else None))
                if kind == "temp" and sw == "WellCat":
                    gf = tidy[(tidy.kind == "temp_fill") & (tidy.software == "WellCat") & (tidy.string == st["key"])]
                    for seg in est_segments(gf):
                        ser.append(dict(name=f"{st['short']} est.", md=seg.md, val=seg.value,
                                        color=STR_COL[st["key"]], width=1.75, dash=MSO_LINE_DASH_STYLE.ROUND_DOT))
        if kind == "temp":
            ser.append(dict(name="Geotérmico", md=geo.md, val=geo.value, color=GEO, width=1.25,
                            dash=MSO_LINE_DASH_STYLE.DASH))
        vals = np.r_[sel.value.values, geo.value.values] if kind == "temp" else \
            np.r_[sel.value.values, tidy[tidy.kind == "press_raw"].value.values]
        xr = nice_range(vals, target=7)
        depth_chart(s, MX + i * Inches(6.25), cy, cw, ch, ser, f"{QTY[kind]} ({u})",
                    SW_LABEL[sw], xr, yr, name=f"Chart {sw}")
    return s


def delta_slide(deck, lc, tidy, kind, foot, strings, results):
    s = deck.slide(footer=foot)
    u = UNIT[kind]
    deck.title(s, f"Δ{'T' if kind == 'temp' else 'P'} ao longo do poço (WellCat − PySafe)",
               f"Valores positivos: WellCat mais {'quente' if kind == 'temp' else 'alto'} que o PySafe · {lc['name']}")
    ser, allm, alld = [], [], []
    for st in strings:
        r = results.get((kind, st["key"]))
        if not r:
            continue
        if kind == "temp":
            md, dv, src = fill_delta(tidy, st["key"])
            df = pd.DataFrame({"md": md, "value": dv, "src": src})
            m = df[df.src != "estimado"]
            ser.append(dict(name=st["short"], md=m.md, val=m.value, color=STR_COL[st["key"]]))
            for seg in est_segments(df):
                ser.append(dict(name=f"{st['short']} est.", md=seg.md, val=seg.value,
                                color=STR_COL[st["key"]], width=1.75, dash=MSO_LINE_DASH_STYLE.ROUND_DOT))
            allm.append(md)
            alld.append(dv)
            continue
        ser.append(dict(name=st["short"],
                        md=r["delta_md"], val=r["delta"], color=STR_COL[st["key"]],
                        marker=4 if r["n"] < 40 else None))
        allm.append(r["delta_md"])
        alld.append(r["delta"])
    ser.append(dict(name="Δ = 0", md=[np.concatenate(allm).min(), np.concatenate(allm).max()],
                    val=[0, 0], color=GEO, width=1.0, dash=MSO_LINE_DASH_STYLE.DASH))
    yr = nice_range(np.concatenate(allm), pad=0.02, step=500)
    xr = nice_range(np.r_[np.concatenate(alld), 0], target=8)
    depth_chart(s, MX, Inches(1.45), Inches(7.2), Inches(5.45), ser, f"Δ ({u})", None, xr, yr,
                name="Chart Delta")
    # tabela-resumo à direita
    x0, y0 = MX + Inches(7.5), Inches(1.55)
    rows = [("Revestimento", "Δ médio", "Δ máx", "MD Δ máx", "Medido")]
    for st in strings:
        r = results.get((kind, st["key"]))
        if r:
            rm = results.get((kind + "_meas", st["key"]), r)
            rows.append((st["short"],
                         f"{fmt(r['d_mean'], kind, True)}", f"{fmt(r['d_absmax'], kind, True)}",
                         f"{fm(r['d_absmax_md'])} m", f"{pct_measured(rm):.0f}%"))
    tbl = s.shapes.add_table(len(rows), 5, x0, y0, Inches(4.8), Inches(0.42) * len(rows)).table
    widths = [1.3, 0.85, 0.85, 1.0, 0.8]
    for j, wv in enumerate(widths):
        tbl.columns[j].width = Inches(wv)
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            c = tbl.cell(i, j)
            c.text = v
            p = c.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.RIGHT
            f = p.runs[0].font
            f.size, f.name = Pt(12), FONT
            f.bold = i == 0
            f.color.rgb = WHITE if i == 0 else INK
            c.fill.solid()
            c.fill.fore_color.rgb = NAVY if i == 0 else (PANEL if i % 2 else WHITE)
            c.margin_left = c.margin_right = Inches(0.08)
    text(s, x0, y0 + Inches(0.42) * len(rows) + Inches(0.3), Inches(4.8), Inches(1.6),
         [f"Unidade: {u}. " + (
             "Δ no grid completo de cada revestimento: lacunas do WellCat interpoladas e trechos sem dado "
             "estimados (pontilhado). % medido = parcela do trecho com dado original do WellCat."
             if kind == "temp" else
             "Δ avaliado nos pontos do WellCat dentro da faixa comum; PySafe interpolado linearmente, "
             "sem atravessar mudanças de curva (ex.: packer).")],
         size=12, color=MUTED)
    return s


def summary_table(slide, x, y, w, entries, size=12):
    """entries: list of (qty, string_label, r, kind)."""
    hdr = ("Grandeza", "Revestimento", "Faixa (m)", "% medido", "Δ médio", "Δ máx (MD)", "Leitura")
    tbl = slide.shapes.add_table(len(entries) + 1, len(hdr), x, y, w, Inches(0.4) * (len(entries) + 1)).table
    widths = [1.15, 1.4, 1.5, 0.95, 1.05, 1.75, 4.5]
    for j, wv in enumerate(widths):
        tbl.columns[j].width = Inches(wv)
    rows = [hdr]
    for qty, lab, r, kind, rm in entries:
        u = UNIT[kind]
        rows.append((qty, lab, f"{fm(r['overlap'][0])}–{fm(r['overlap'][1])}", f"{pct_measured(rm):.0f}%",
                     f"{fmt(r['d_mean'], kind, True)} {u}",
                     f"{fmt(r['d_absmax'], kind, True)} {u} ({fm(r['d_absmax_md'])})",
                     reading(r, kind).split(" (")[0]))
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            c = tbl.cell(i, j)
            c.text = v
            p = c.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.RIGHT if j in (3, 4, 5) and i else PP_ALIGN.LEFT
            f = p.runs[0].font
            f.size, f.name = Pt(size), FONT
            f.bold = i == 0
            f.color.rgb = WHITE if i == 0 else INK
            c.fill.solid()
            c.fill.fore_color.rgb = NAVY if i == 0 else (PANEL if i % 2 else WHITE)
            c.margin_left = c.margin_right = Inches(0.08)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
    return tbl


def coverage_slide(deck, lc, tidy, foot, strings, pmap):
    s = deck.slide(footer=foot)
    deck.title(s, "Escopo, cobertura de dados e premissas", lc["name"])
    cats, base, before, length, after, colors = [], [], [], [], [], []
    for st in strings:
        for sw in ["PySafe", "WellCat"]:
            g = tidy[(tidy.kind == "temp") & (tidy.software == sw) & (tidy.string == st["key"])]
            if len(g):
                f = tidy[(tidy.kind == "temp_fill") & (tidy.software == sw) & (tidy.string == st["key"])]
                lo = min(g.md.min(), f.md.min()) if len(f) else g.md.min()
                hi = max(g.md.max(), f.md.max()) if len(f) else g.md.max()
                cats.append(f"{st['short']} · {SW_LABEL[sw]}")
                base.append(lo)
                before.append(g.md.min() - lo)
                length.append(g.md.max() - g.md.min())
                after.append(hi - g.md.max())
                colors.append(SW[sw])
    cd = CategoryChartData()
    cd.categories = cats
    cd.add_series("Topo", base)
    cd.add_series("Estimado acima", before)
    cd.add_series("Medido", length)
    cd.add_series("Estimado abaixo", after)
    gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_STACKED, MX, Inches(1.45), Inches(6.4), Inches(5.45), cd)
    gf.name = "Chart Coverage"
    ch = gf.chart
    ch.font.name, ch.font.size = FONT, Pt(10)
    ch.has_title = True
    ch.chart_title.text_frame.text = "Cobertura térmica (MD, m) · claro = WellCat estimado"
    rr = ch.chart_title.text_frame.paragraphs[0].runs[0]
    rr.font.size, rr.font.bold, rr.font.color.rgb = Pt(13), True, INK
    ch.has_legend = False
    p = ch.plots[0]
    p.gap_width = 60
    p.overlap = 100
    p.series[0].format.fill.background()
    p.series[0].format.line.fill.background()
    for i, c in enumerate(colors):
        pt = p.series[2].points[i]
        pt.format.fill.solid()
        pt.format.fill.fore_color.rgb = c
    for si in (1, 3):
        p.series[si].format.fill.solid()
        p.series[si].format.fill.fore_color.rgb = EST
    ca, va = ch.category_axis, ch.value_axis
    ca.reverse_order = True
    ca.tick_labels.font.size = Pt(10)
    ca.tick_labels.font.color.rgb = INK
    ca.format.line.color.rgb = GRID
    va.minimum_scale = math.floor(min(base) / 500) * 500
    va.maximum_scale = math.ceil(max(b + x + l + y for b, x, l, y in zip(base, before, length, after)) / 500) * 500
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = GRID
    va.tick_labels.font.size = Pt(10)
    va.tick_labels.font.color.rgb = MUTED
    va.tick_labels.number_format = "#,##0"
    va.tick_labels.number_format_is_linked = False
    va.format.line.color.rgb = GRID
    # premissas à direita
    x0, wcol = MX + Inches(6.75), Inches(5.55)
    items = [
        ("Comparação", "WellCat × SCORE-PySafe, revestimento a revestimento (22\", 14\", liner 10 3/4\" e COP/COI). Condutor de 36\" fora do escopo."),
        ("Δ", "Δ = WellCat − PySafe; média ponderada pela profundidade. Temperatura: no grid completo de cada revestimento. "
              "Pressão: nos pontos do WellCat, com PySafe interpolado linearmente."),
        ("Pressão interna", "; ".join(
            f"{k.replace('10.75', '10 3/4')}\" ↔ " + " + ".join(
                f"{seg['curve']} ({fm(seg['md_from'])}–{fm(seg['md_to'])} m)" for seg in v["pysafe"])
            for k, v in sorted(pmap.items(), key=lambda kv: [st["key"] for st in strings].index(kv[0])
                               if kv[0] in [st["key"] for st in strings] else 99)) + "."),
        ("Preenchimento térmico", "Lacunas do WellCat: interpolação linear entre pontos medidos. Fora da faixa medida: "
         "14\" e COP/COI por reta até o 1º ponto do liner 10 3/4\" e, abaixo dele, seguindo o liner; demais por reta curta (≤ 300 m)."),
        ("Unidades", "Temperatura em °C, pressão em psi, profundidade em MD (m)."),
    ]
    y = Inches(1.5)
    for head, body in items:
        text(s, x0, y, wcol, Inches(0.3), head, size=15, bold=True, color=NAVY)
        tb = text(s, x0, y + Inches(0.33), wcol, Inches(0.9), body, size=13, color=INK)
        y += Inches(0.33) + Inches(0.24) * math.ceil(len(body) / 66) + Inches(0.22)
    return s


REG_COL = {"Fluido": RGBColor(0xF2, 0xC4, 0xCA), "Anular A": RGBColor(0xD9, 0xD2, 0xEE),
           "Anular B": RGBColor(0xC6, 0xD6, 0xE3), "Anular C": RGBColor(0xCF, 0xE6, 0xD6),
           "Cimento": RGBColor(0xB8, 0xBE, 0xC4)}
FORMATION = RGBColor(0xE8, 0xE2, 0xD6)


def schematic_slide(deck, lc, cfg, foot, strings):
    """Esquemático (meio-poço) com os espaços do PySafe e o mapa para o WellCat."""
    from architecture import bands, describe
    arch = get_arch(lc, cfg)
    s = deck.slide(footer=foot)
    deck.title(s, "PySafe nomeia espaços, WellCat nomeia revestimentos",
               "Pressão interna de cada revestimento = curva do espaço imediatamente dentro dele, trecho a trecho")
    label = {st["key"]: st["short"] for st in strings}
    order = [c["key"] for c in arch["casings"]]
    y0, y1 = Inches(1.6), Inches(6.75)
    d0, d1 = arch["wellhead"], arch["td"]
    Y = lambda d: int(y0 + (y1 - y0) * (d - d0) / (d1 - d0))
    X0, wall, step, hole = Inches(2.3), Inches(0.05), Inches(0.5), Inches(0.32)
    xcop = X0 + Inches(0.55)
    xw = {k: xcop + step * (i + 1) for i, k in enumerate(order)}
    xmax = xw[order[-1]] + wall + hole + Inches(0.25)
    rect(s, X0, y0, xmax - X0, y1 - y0, FORMATION, shape=MSO_SHAPE.RECTANGLE, name="Formação")

    def xr(edge, inner):
        if edge is None:
            return X0
        if edge == "COP":
            return xcop
        if edge == "OH":
            return xw[inner] + wall + hole
        return xw[edge]

    for bd in bands(arch):
        ya, yb = Y(bd["top"]), Y(bd["base"])
        for lab, inner, outer in bd["regions"]:
            xa = X0 if inner is None else (xcop + wall if inner == "COP" else xw[inner] + wall)
            xb = xr(outer, inner)
            r = rect(s, xa, ya, xb - xa, yb - ya, REG_COL.get(lab, PANEL), shape=MSO_SHAPE.RECTANGLE)
            if yb - ya > Inches(0.32) and xb - xa > Inches(0.2) and lab != "Cimento":
                text(s, xa, ya, xb - xa, yb - ya, lab.replace("Anular ", ""), size=11, bold=True,
                     color=INK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    # paredes
    rect(s, xcop, Y(arch["wellhead"]), wall, Y(arch["packer"]) - Y(arch["wellhead"]), INK,
         shape=MSO_SHAPE.RECTANGLE, name="COP/COI")
    for c in arch["casings"]:
        rect(s, xw[c["key"]], Y(c["top"]), wall, Y(c["shoe"]) - Y(c["top"]), INK,
             shape=MSO_SHAPE.RECTANGLE, name=f"Parede {c['key']}")
    inner_at_packer = next(c["key"] for c in arch["casings"] if c["top"] <= arch["packer"] <= c["shoe"])
    rect(s, xcop, Y(arch["packer"]) - Inches(0.04), xw[inner_at_packer] - xcop, Inches(0.08), INK,
         shape=MSO_SHAPE.RECTANGLE, name="Packer")
    # rótulos de profundidade (à esquerda, sem sobreposição)
    marks = [(arch["wellhead"], "Cabeça do poço")]
    for c in arch["casings"]:
        if c["top"] > arch["wellhead"]:
            marks.append((c["top"], f"Topo {label.get(c['key'], c['key'])}"))
        marks.append((c["shoe"], f"Sapata {label.get(c['key'], c['key'])}"))
        if c.get("toc", c["top"]) > c["top"]:
            marks.append((c["toc"], f"TOC {label.get(c['key'], c['key'])}"))
    marks.append((arch["packer"], "Packer / base COP"))
    marks = sorted({m[0]: m for m in marks}.values())
    last = -1e9
    for d, t in marks:
        yy = max(Y(d) - Inches(0.1), last + Inches(0.22))
        last = yy
        text(s, MX, yy, X0 - MX - Inches(0.12), Inches(0.2), [[(f"{fm(d)} m ", {"bold": True, "size": 10}),
             (t, {"size": 10, "color": MUTED})]], size=10, align=PP_ALIGN.RIGHT)
        ln = s.shapes.add_connector(1, X0 - Inches(0.1), yy + Inches(0.1), X0, Y(d))
        ln.line.color.rgb = MUTED
        ln.line.width = Pt(0.75)
    # legenda
    ly = Inches(6.82)
    for i, (lab, col) in enumerate([("Fluido", REG_COL["Fluido"]), ("Anular A", REG_COL["Anular A"]),
                                    ("Anular B", REG_COL["Anular B"]), ("Cimento", REG_COL["Cimento"]),
                                    ("Formação", FORMATION)]):
        lx = MX + Inches(0.85) * i
        rect(s, lx, ly + Inches(0.03), Inches(0.14), Inches(0.14), col, shape=MSO_SHAPE.RECTANGLE)
        text(s, lx + Inches(0.18), ly, Inches(0.7), Inches(0.2), lab, size=10, color=MUTED)
    # tabela de correspondência
    x0, wtab = Inches(5.35), Inches(7.45)
    rows = [("Revestimento (WellCat)", "Trecho MD (m)", "Pressão interna no PySafe")]
    rank = {st["key"]: i for i, st in enumerate(strings)}
    for k, curve, a, b in sorted(describe(arch), key=lambda r: (rank.get(r[0], 99), r[2])):
        rows.append((label.get(k, k), f"{fm(a)}–{fm(b)}", curve))
    tbl = s.shapes.add_table(len(rows), 3, x0, Inches(1.6), wtab, Inches(0.38) * len(rows)).table
    for j, wv in enumerate([2.4, 2.0, 3.05]):
        tbl.columns[j].width = Inches(wv)
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            c = tbl.cell(i, j)
            c.text = v
            f = c.text_frame.paragraphs[0].runs[0].font
            f.size, f.name, f.bold = Pt(12), FONT, i == 0
            f.color.rgb = WHITE if i == 0 else INK
            c.fill.solid()
            c.fill.fore_color.rgb = NAVY if i == 0 else (PANEL if i % 2 else WHITE)
    c14 = next((c for c in arch["casings"] if c.get("toc", c["top"]) > c["top"]), None)
    rules = [
        "Anular A: entre a COP/COI e o 1º revestimento presente. De baixo para cima: liner 10 3/4\" até o "
        "topo do liner e, acima dele, o 14\".",
        "Anular B: só existe onde não há cimento atrás do revestimento interno. Começa no TOC"
        + (f" do {label.get(c14['key'])} ({fm(c14['toc'])} m), em poço aberto com o fluido de perfuração,"
           if c14 else ",") + " e, acima da sapata do 22\", passa a ser o interior do 22\".",
        "Abaixo do packer, o interior do liner é o próprio Fluido de produção.",
        "O Anular B também é a pressão externa do 14\"; a exportação do WellCat traz só pressão interna, "
        "então essa face fica fora da comparação.",
    ]
    ty = Inches(1.6) + Inches(0.38) * len(rows) + Inches(0.3)
    text(s, x0, ty, wtab, y1 - ty + Inches(0.3),
         [[(f"{i+1}. {r}", {"space_after": 6})] for i, r in enumerate(rules)], size=13)
    s.notes_slide.notes_text_frame.text = (
        "Correspondência gerada a partir de architecture no loadcases.yaml (src/architecture.py). "
        "Verificação nos dados: a 2.500 m, 22\" WellCat = 6.721 psi; Anular B PySafe = 6.750 psi; "
        "Anular A = 10.774 psi.")
    return s


def section_slide(deck, kicker, title, sub):
    s = deck.slide(dark=True)
    text(s, Inches(0.8), Inches(2.55), Inches(11.5), Inches(0.4), kicker.upper(), size=14, bold=True,
         color=ICE)
    text(s, Inches(0.8), Inches(3.0), Inches(11.5), Inches(1.0), title, size=40, bold=True,
         color=WHITE, font=HEAD)
    if sub:
        text(s, Inches(0.8), Inches(4.05), Inches(11.5), Inches(0.8), sub, size=16, color=ICE)
    return s


# --------------------------------------------------------------- montagem
def build(cfg, only=None, out=None):
    lcs = [lc for lc in cfg["loadcases"] if not only or lc["id"] in only]
    deck = Deck(cfg)
    strings = cfg["strings"]
    data = {}
    for lc in lcs:
        tidy = extract_loadcase(lc, cfg)
        res = {}
        for st in strings:
            res[("temp", st["key"])] = compare(tidy, "temp_fill", st["key"])      # trecho completo
            res[("temp_meas", st["key"])] = compare(tidy, "temp", st["key"])      # só medido
            res[("press", st["key"])] = compare(tidy, "press", st["key"])
        data[lc["id"]] = (tidy, {k: v for k, v in res.items() if v})

    # capa
    s = deck.slide(dark=True)
    text(s, Inches(0.8), Inches(1.9), Inches(11.5), Inches(0.4), "PROJETO DE REVESTIMENTO · VERIFICAÇÃO CRUZADA",
         size=14, bold=True, color=ICE)
    text(s, Inches(0.8), Inches(2.4), Inches(11.5), Inches(1.6),
         ["SCORE-PySafe × WellCat", "Temperatura e pressão por revestimento"],
         size=40, bold=True, color=WHITE, font=HEAD)
    text(s, Inches(0.8), Inches(4.35), Inches(11.5), Inches(1.2),
         [f"Poço {cfg['well']}",
          "Carregamento" + ("s: " if len(lcs) > 1 else ": ") + " · ".join(lc["name"] for lc in lcs)],
         size=18, color=ICE)
    text(s, Inches(0.8), Inches(6.5), Inches(6), Inches(0.4),
         dt.date.today().strftime("%m/%Y"), size=14, color=ICE)

    # resumo executivo
    for lc in lcs:
        tidy, res = data[lc["id"]]
        foot = f"{cfg['well']} · {lc['short']} · Fontes: SCORE-PySafe, WellCat"
        s = deck.slide(footer=foot)
        deck.title(s, "Resumo da comparação", lc["name"])
        entries = [(("Temperatura" if k == "temp" else "Pressão"), st["short"], res[(k, st["key"])], k,
                    res.get((k + "_meas", st["key"]), res[(k, st["key"])]))
                   for k in ("temp", "press") for st in strings if (k, st["key"]) in res]
        summary_table(s, MX, Inches(1.5), Inches(12.3), entries, size=12)
        obs = lc.get("observations") or []
        if obs:
            yb = Inches(1.5) + Inches(0.4) * (len(entries) + 1) + Inches(0.35)
            text(s, MX, yb, Inches(12.3), Inches(0.3), "Leitura", size=15, bold=True, color=NAVY)
            text(s, MX, yb + Inches(0.38), Inches(12.3), Inches(7.0) - yb - Inches(0.38),
                 [[(f"{i+1}. {o}", {"space_after": 4})] for i, o in enumerate(obs[:3])], size=13)

        coverage_slide(deck, lc, tidy, foot, strings, get_pressure_map(lc, cfg))
        schematic_slide(deck, lc, cfg, foot, strings)

        section_slide(deck, lc["short"], "Temperatura", "Curvas por revestimento e comparação um-a-um")
        overview_slide(deck, lc, tidy, "temp", foot, strings)
        for st in strings:
            string_slide(deck, lc, tidy, st, "temp", foot, res)
        delta_slide(deck, lc, tidy, "temp", foot, strings, res)

        section_slide(deck, lc["short"], "Pressão interna", "Curvas por revestimento e comparação um-a-um")
        overview_slide(deck, lc, tidy, "press", foot, strings)
        for st in strings:
            string_slide(deck, lc, tidy, st, "press", foot, res)
        delta_slide(deck, lc, tidy, "press", foot, strings, res)

    # consolidado (quando houver mais de um carregamento)
    if len(lcs) > 1:
        s = deck.slide(footer=f"{cfg['well']} · Consolidado")
        deck.title(s, "Consolidado – Δ máximo por carregamento", "WellCat − PySafe; °C para temperatura, psi para pressão")
        keys = [(k, st) for k in ("temp", "press") for st in strings]
        hdr = ["Carregamento"] + [f"{'T' if k == 'temp' else 'P'} {st['key'].replace('10.75', '10 3/4')}" for k, st in keys]
        tbl = s.shapes.add_table(len(lcs) + 1, len(hdr), MX, Inches(1.5), Inches(12.3),
                                 Inches(0.4) * (len(lcs) + 1)).table
        tbl.columns[0].width = Inches(3.1)
        for j in range(1, len(hdr)):
            tbl.columns[j].width = Inches(9.2 / (len(hdr) - 1))
        for j, h in enumerate(hdr):
            tbl.cell(0, j).text = h
        for i, lc in enumerate(lcs, 1):
            _, res = data[lc["id"]]
            tbl.cell(i, 0).text = lc["short"]
            for j, (k, st) in enumerate(keys, 1):
                r = res.get((k, st["key"]))
                tbl.cell(i, j).text = fmt(r["d_absmax"], k, True) if r else "–"
        for i in range(len(lcs) + 1):
            for j in range(len(hdr)):
                c = tbl.cell(i, j)
                f = c.text_frame.paragraphs[0].runs[0].font
                f.size, f.name, f.bold = Pt(12), FONT, i == 0
                f.color.rgb = WHITE if i == 0 else INK
                c.fill.solid()
                c.fill.fore_color.rgb = NAVY if i == 0 else (PANEL if i % 2 else WHITE)

    # fechamento
    s = deck.slide(dark=True)
    text(s, Inches(0.8), Inches(0.7), Inches(11.5), Inches(0.8), "Pontos a confirmar", size=36,
         bold=True, color=WHITE, font=HEAD)
    pts = []
    for lc in lcs:
        pts += lc.get("next_steps") or []
    if not pts:
        pts = ["Confirmar mapeamento curva PySafe ↔ revestimento na pressão interna."]
    text(s, Inches(0.8), Inches(1.8), Inches(11.5), Inches(5),
         [[(f"{i+1}.  {p}", {"space_after": 16})] for i, p in enumerate(pts)], size=20, color=ICE)

    out = Path(out or ROOT / "out" / f"Comparativo_PySafe_x_WellCat_{cfg['well'].replace(' ', '_')}.pptx")
    out.parent.mkdir(exist_ok=True, parents=True)
    deck.prs.save(out)
    # exporta também o tidy + estatísticas para auditoria
    for lc in lcs:
        tidy, res = data[lc["id"]]
        tidy.to_csv(out.parent / f"{lc['id']}_tidy.csv", index=False)
        export_filled_xlsx(tidy, cfg, lc, out.parent / f"{lc['id']}_temperaturas_preenchidas.xlsx")
        rows = []
        for (k, key), r in res.items():
            rows.append({"loadcase": lc["id"], "kind": k, "string": key,
                         **{n: (v if not isinstance(v, (tuple, np.ndarray)) else None) for n, v in r.items()},
                         "overlap_from": r["overlap"][0], "overlap_to": r["overlap"][1]})
        import pandas as pd
        pd.DataFrame(rows).drop(columns=["delta", "delta_md", "overlap", "pysafe_md", "wellcat_md"]).to_csv(
            out.parent / f"{lc['id']}_stats.csv", index=False)
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--only", nargs="*")
    ap.add_argument("--out")
    a = ap.parse_args()
    print(build(load_config(), a.only, a.out))
