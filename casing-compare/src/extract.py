"""Extração e normalização dos dados SCORE-PySafe x WellCat.

Saída: DataFrame "tidy" com colunas
    loadcase, kind (temp|press), software (PySafe|WellCat), string, md, value
"""
from __future__ import annotations

import re
import unicodedata
from pathlib import Path

import numpy as np
import pandas as pd
import yaml

ROOT = Path(__file__).resolve().parents[1]


def load_config(path: Path | None = None) -> dict:
    return yaml.safe_load((path or ROOT / "loadcases.yaml").read_text(encoding="utf-8"))


def _norm(s: str) -> str:
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().lower()


def _find_sheet(xls: pd.ExcelFile, wanted: str) -> str:
    w = _norm(wanted)
    for name in xls.sheet_names:
        if _norm(name) == w:
            return name
    for name in xls.sheet_names:
        if w in _norm(name) or _norm(name) in w:
            return name
    raise KeyError(f"Aba '{wanted}' não encontrada. Abas: {xls.sheet_names}")


# ---------------------------------------------------------------- temperatura
def _thermal(xls, sheet, software, strings, lc):
    df = pd.read_excel(xls, sheet_name=_find_sheet(xls, sheet))
    cols = {_norm(c): c for c in df.columns}
    md_col = next(c for n, c in cols.items() if n.startswith("md"))
    geo_col = next((c for n, c in cols.items() if "geoterm" in n), None)
    rows = []
    for st in strings:
        m = _norm(st["match"])
        col = next((c for n, c in cols.items() if m in n), None)
        if col is None:
            continue
        sub = df[[md_col, col]].dropna()
        rows += [(lc, "temp", software, st["key"], float(a), float(b)) for a, b in sub.values]
    if geo_col is not None:
        sub = df[[md_col, geo_col]].dropna()
        rows += [(lc, "temp", software, "GEO", float(a), float(b)) for a, b in sub.values]
    return rows


# ------------------------------------------------------------------- pressão
def _press_pysafe(xls, sheet, pmap, lc):
    df = pd.read_excel(xls, sheet_name=_find_sheet(xls, sheet))
    cols = {_norm(c): c for c in df.columns}
    md_col = next(c for n, c in cols.items() if n.startswith("md"))
    rows = []
    # curvas brutas (para o slide de visão geral)
    for n, c in cols.items():
        if c == md_col:
            continue
        curve = c.split("-")[0].strip()
        sub = df[[md_col, c]].dropna()
        rows += [(lc, "press_raw", "PySafe", curve, float(a), float(b)) for a, b in sub.values]
    # recorte por revestimento segundo o mapa
    for key, spec in pmap.items():
        for si, seg in enumerate(spec["pysafe"]):
            c = next(c for n, c in cols.items() if n.startswith(_norm(seg["curve"])))
            sub = df[[md_col, c]].dropna()
            sub = sub[(sub[md_col] >= seg["md_from"] - 1e-6) & (sub[md_col] <= seg["md_to"] + 1e-6)]
            rows += [(lc, "press", "PySafe", key, float(a), float(b), si) for a, b in sub.values]
    return rows


def _press_wellcat(xls, sheet, pmap, lc):
    """WellCat exporta blocos [idx | MD | P] lado a lado, 2 linhas de cabeçalho."""
    raw = pd.read_excel(xls, sheet_name=_find_sheet(xls, sheet), header=None)
    rows = []
    hdr = {j: f"{raw.iat[0, j]} {raw.iat[1, j]}" for j in raw.columns}
    md_cols = [j for j in raw.columns if _norm(raw.iat[0, j]).startswith("md")]
    for j in md_cols:
        label = _norm(hdr[j + 1])
        key = next((k for k, s in pmap.items() if _norm(s["wellcat_match"]) in label), None)
        if key is None:
            continue
        sub = raw.iloc[2:, [j, j + 1]].apply(pd.to_numeric, errors="coerce").dropna()
        rows += [(lc, "press", "WellCat", key, float(a), float(b)) for a, b in sub.values]
    return rows


def get_arch(lc: dict, cfg: dict) -> dict:
    return lc.get("architecture") or cfg["architecture"]


def get_pressure_map(lc: dict, cfg: dict) -> dict:
    """Mapa de pressão interna derivado da arquitetura (ou explícito em pressure_map)."""
    from architecture import pressure_map
    return lc.get("pressure_map") or pressure_map(get_arch(lc, cfg))


def extract_loadcase(lc: dict, cfg: dict) -> pd.DataFrame:
    path = ROOT / lc["file"]
    xls = pd.ExcelFile(path)
    sh = lc["sheets"]
    pmap = get_pressure_map(lc, cfg)
    rows = []
    rows += _thermal(xls, sh["temp_pysafe"], "PySafe", cfg["strings"], lc["id"])
    rows += _thermal(xls, sh["temp_wellcat"], "WellCat", cfg["strings"], lc["id"])
    rows += _press_pysafe(xls, sh["press_pysafe"], pmap, lc["id"])
    rows += _press_wellcat(xls, sh["press_wellcat"], pmap, lc["id"])
    rows = [r if len(r) == 7 else (*r, 0) for r in rows]
    out = pd.DataFrame(rows, columns=["loadcase", "kind", "software", "string", "md", "value", "seg"])
    out["src"] = "medido"
    out = pd.concat([out, fill_temperature(out, cfg, lc)], ignore_index=True)
    return out.sort_values(["kind", "software", "string", "md"], kind="stable").reset_index(drop=True)


# ------------------------------------------------- preenchimento térmico
def _tail_fit(md, v, at, n):
    """Reta de mínimos quadrados nos n pontos da ponta mais próxima de `at`."""
    o = np.argsort(md)
    md, v = md[o], v[o]
    sl = slice(-n, None) if at > md.max() else slice(0, n)
    a, b = np.polyfit(md[sl], v[sl], 1)
    return a * np.asarray(at) + b


def fill_temperature(tidy: pd.DataFrame, cfg: dict, lc: dict) -> pd.DataFrame:
    """Gera kind='temp_fill': PySafe e WellCat no mesmo grid de MD de cada
    revestimento (trecho PySafe topo→sapata ∪ MDs do WellCat).

    src = medido | interpolado | estimado
      interpolado: reta entre dois pontos medidos vizinhos (lacunas internas)
      estimado   : fora da faixa medida do WellCat, conforme temp_fill.extend
         - linear_tail: reta pelos n pontos da ponta (só até max_len_m)
         - anchor     : reta do último ponto medido até o 1º ponto medido do
                        revestimento de referência (ref_string) e, a partir
                        dele, segue os valores medidos desse revestimento
    """
    fcfg = {**(cfg.get("temp_fill") or {}), **(lc.get("temp_fill") or {})}
    dflt = fcfg.get("extend_default", {"method": "linear_tail", "n": 2, "max_len_m": 300})
    ext_cfg = fcfg.get("extend", {})
    T = tidy[tidy.kind == "temp"]
    out = []
    for st in cfg["strings"]:
        k = st["key"]
        ps = T[(T.software == "PySafe") & (T.string == k)]
        wc = T[(T.software == "WellCat") & (T.string == k)]
        if ps.empty or wc.empty:
            continue
        lo, hi = ps.md.min(), ps.md.max()
        grid = np.unique(np.r_[ps.md.values, wc.md.values[(wc.md >= lo) & (wc.md <= hi)]])
        xp, yp = _monotonic(ps.md.values, ps.value.values)
        for m in grid:
            out.append((lc["id"], "temp_fill", "PySafe", k, m, float(np.interp(m, xp, yp)), 0,
                        "medido" if np.isclose(ps.md.values, m).any() else "interpolado"))
        xw, yw = _monotonic(wc.md.values, wc.value.values)
        meas = set(np.round(wc.md.values, 3))
        e = {**dflt, **ext_cfg.get(k, {})}
        for m in grid:
            if xw.min() <= m <= xw.max():
                v, src = float(np.interp(m, xw, yw)), ("medido" if round(m, 3) in meas else "interpolado")
            elif e["method"] == "anchor":
                r = T[(T.software == "WellCat") & (T.string == e["ref_string"])]
                xr, yr = _monotonic(r.md.values, r.value.values)
                if m > xw.max():
                    if m <= xr.min():   # reta: último medido → 1º ponto da referência
                        v = float(np.interp(m, [xw.max(), xr.min()], [yw[-1], yr[0]]))
                    else:               # segue a referência medida
                        v = float(np.interp(m, xr, yr))
                else:
                    continue
                src = "estimado"
            else:
                dist = (m - xw.max()) if m > xw.max() else (xw.min() - m)
                if dist > e.get("max_len_m", 300):
                    continue
                v, src = float(_tail_fit(xw, yw, m, e.get("n", 2))), "estimado"
            out.append((lc["id"], "temp_fill", "WellCat", k, m, v, 0, src))
    return pd.DataFrame(out, columns=["loadcase", "kind", "software", "string", "md", "value", "seg", "src"])


def export_filled_xlsx(tidy: pd.DataFrame, cfg: dict, lc: dict, path: Path) -> Path:
    """Tabela completa (MD × revestimento × software) com a origem de cada valor."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    F = tidy[tidy.kind == "temp_fill"]
    wb = Workbook()
    wb.remove(wb.active)
    fills = {"interpolado": PatternFill("solid", fgColor="FFF2CC"),
             "estimado": PatternFill("solid", fgColor="F8CBAD")}
    for st in cfg["strings"]:
        g = F[F.string == st["key"]]
        if g.empty:
            continue
        ws = wb.create_sheet(st["short"].replace('"', "in").replace("/", "-")[:31])
        ws.append(["MD (m)", "PySafe (°C)", "Origem PySafe", "WellCat (°C)", "Origem WellCat", "Δ WC−PS (°C)"])
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="0E2433")
            c.alignment = Alignment(horizontal="center", wrap_text=True)
        pv = g.pivot_table(index="md", columns="software", values="value", aggfunc="last")
        sv = g.pivot_table(index="md", columns="software", values="src", aggfunc="last")
        for i, (m, row) in enumerate(pv.iterrows(), start=2):
            ps, wcv = row.get("PySafe"), row.get("WellCat")
            ws.append([round(m, 2), ps, sv.at[m, "PySafe"] if "PySafe" in sv else None,
                       None if pd.isna(wcv) else wcv,
                       sv.at[m, "WellCat"] if "WellCat" in sv and not pd.isna(wcv) else None,
                       f"=IF(D{i}=\"\",\"\",D{i}-B{i})"])
            for col, srcv in ((2, ws.cell(i, 3).value), (4, ws.cell(i, 5).value)):
                if srcv in fills:
                    ws.cell(i, col).fill = fills[srcv]
            for col in (2, 4, 6):
                ws.cell(i, col).number_format = "0.00"
        for col, w in zip("ABCDEF", (10, 12, 14, 12, 15, 13)):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = "A2"
    leg = wb.create_sheet("Legenda", 0)
    for r in [["Preenchimento térmico – " + lc["name"]], [],
              ["medido", "valor da planilha original"],
              ["interpolado", "reta entre os dois pontos medidos vizinhos"],
              ["estimado", "fora da faixa medida do WellCat (ver método em loadcases.yaml → temp_fill)"]]:
        leg.append(r)
    leg["A1"].font = Font(bold=True, size=13)
    leg["A4"].fill, leg["A5"].fill = fills["interpolado"], fills["estimado"]
    leg.column_dimensions["A"].width, leg.column_dimensions["B"].width = 14, 80
    wb.save(path)
    return path


# ---------------------------------------------------------------- comparação
def _monotonic(md, v):
    """Remove MDs repetidos (mantém o último valor) para interpolação."""
    s = pd.Series(v, index=md).groupby(level=0).last().sort_index()
    return s.index.values, s.values


def compare(tidy: pd.DataFrame, kind: str, key: str, compare_to: float | None = None) -> dict:
    """Delta = WellCat − PySafe, avaliado nas profundidades do WellCat dentro da
    faixa comum (PySafe interpolado linearmente)."""
    a = tidy[(tidy.kind == kind) & (tidy.software == "PySafe") & (tidy.string == key)]
    b = tidy[(tidy.kind == kind) & (tidy.software == "WellCat") & (tidy.string == key)]
    if a.empty or b.empty:
        return {}
    # PySafe pode ser composto por trechos (ex.: Anular A acima do packer +
    # Fluido abaixo): interpola só dentro de cada trecho, nunca atravessando o degrau
    md_l, d_l = [], []
    for _, g in a.groupby("seg"):
        xa, ya = _monotonic(g.md.values, g.value.values)
        bb = b[(b.md >= xa.min()) & (b.md <= xa.max())]
        md_l.append(bb.md.values)
        d_l.append(bb.value.values - np.interp(bb.md.values, xa, ya))
    md_c, d = np.concatenate(md_l), np.concatenate(d_l)
    if d.size == 0:
        return {}
    o = np.argsort(md_c, kind="stable")
    md_c, d = md_c[o], d[o]
    lo, hi = md_c.min(), md_c.max()
    bb = pd.DataFrame({"md": md_c})
    i = int(np.argmax(np.abs(d)))
    pa, pb = a.value, b.value
    return {
        "pysafe_md": (a.md.min(), a.md.max()),
        "wellcat_md": (b.md.min(), b.md.max()),
        "overlap": (lo, hi),
        "n": len(d),
        "pysafe_max": pa.max(), "pysafe_md_at_max": a.md.values[int(np.argmax(pa.values))],
        "wellcat_max": pb.max(), "wellcat_md_at_max": b.md.values[int(np.argmax(pb.values))],
        # média ponderada pela profundidade (trapézios) — o grid não é uniforme
        "d_mean": float(np.trapezoid(d, md_c) / (md_c[-1] - md_c[0])) if md_c[-1] > md_c[0] else float(np.mean(d)),
        "d_absmax": float(d[i]), "d_absmax_md": float(bb.md.values[i]),
        "d_min": float(d.min()), "d_max": float(d.max()),
        "delta_md": bb.md.values, "delta": d,
    }


if __name__ == "__main__":
    cfg = load_config()
    for lc in cfg["loadcases"]:
        t = extract_loadcase(lc, cfg)
        out = ROOT / "out" / f"{lc['id']}_tidy.csv"
        out.parent.mkdir(exist_ok=True)
        t.to_csv(out, index=False)
        print(lc["id"], t.groupby(["kind", "software", "string"]).md.agg(["count", "min", "max"]))
        for kind in ("temp", "press"):
            for st in cfg["strings"]:
                r = compare(t, kind, st["key"])
                if r:
                    print(kind, st["key"], {k: (round(v, 2) if isinstance(v, float) else v)
                                            for k, v in r.items() if k not in ("delta", "delta_md")})
