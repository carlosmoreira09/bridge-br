"""Arquitetura do poço → correspondência PySafe (Fluido / Anular A / B / C…) ×
revestimento (WellCat), seguindo a convenção do SCORE-PySafe:

- Fluido     : interior da COP/COI; abaixo do packer, interior do revestimento mais interno.
- Anular A   : entre a COP/COI e o 1º revestimento presente naquela profundidade
               (de baixo para cima: liner 10 3/4" e depois 14").
- Anular B,C : entre revestimentos consecutivos presentes na profundidade, ou entre o
               revestimento e o poço aberto, SOMENTE acima do TOC do revestimento interno
               (abaixo do TOC é cimento → não há anular).

Assim a pressão interna de cada revestimento é a curva do espaço imediatamente
dentro dele, trecho a trecho.
"""
from __future__ import annotations

LETTERS = "ABCDEFG"


def _breaks(arch):
    pts = {arch["wellhead"], arch["packer"], arch["td"]}
    for c in arch["casings"]:
        pts |= {c["top"], c["shoe"], c.get("toc", c["top"])}
    return sorted(p for p in pts if arch["wellhead"] <= p <= arch["td"])


def bands(arch):
    """Faixas de profundidade com o conteúdo radial (de dentro para fora).

    Cada faixa: dict(top, base, present=[keys inner→outer], regions=[...])
    regions: lista de (label, inner_key, outer_key) — inner_key 'COP'/None,
    outer_key = revestimento ou 'OH' (poço aberto). label ∈ Fluido, Anular X, Cimento.
    """
    out = []
    b = _breaks(arch)
    order = [c["key"] for c in arch["casings"]]
    cas = {c["key"]: c for c in arch["casings"]}
    for top, base in zip(b[:-1], b[1:]):
        mid = (top + base) / 2
        present = [k for k in order if cas[k]["top"] <= mid <= cas[k]["shoe"]]
        regs = []
        cop = mid < arch["packer"]
        if cop:
            regs.append(("Fluido", None, "COP"))
            first_outer = present[0] if present else "OH"
            regs.append(("Anular A", "COP", first_outer))
        else:
            regs.append(("Fluido", None, present[0] if present else "OH"))
        nxt = 1  # "A" é sempre o anular da COP/COI; externos começam em B
        chain = present + ["OH"]
        for i, k in enumerate(present):
            outer = chain[i + 1]
            toc = cas[k].get("toc", cas[k]["top"])
            if mid > toc:   # abaixo do TOC (mais profundo): cimentado
                regs.append(("Cimento", k, outer))
            else:
                regs.append((f"Anular {LETTERS[nxt + i]}", k, outer))
            if outer == "OH":
                break
        out.append(dict(top=top, base=base, present=present, regions=regs))
    return out


def pressure_map(arch, inset=0.05):
    """Mapa {casing_key: {wellcat_match, pysafe: [segmentos]}} para a pressão interna.

    `inset` recua os limites internos entre trechos para não interpolar sobre o
    degrau (ex.: topo do liner, packer)."""
    segs = {}
    for bd in bands(arch):
        for label, inner, outer in bd["regions"]:
            if outer in ("COP", "OH") or label == "Cimento":
                continue
            segs.setdefault(outer, []).append([label, bd["top"], bd["base"]])
    out = {}
    for k, lst in segs.items():
        merged = []
        for lab, t, b in lst:
            if merged and merged[-1][0] == lab and abs(merged[-1][2] - t) < 1e-6:
                merged[-1][2] = b
            else:
                merged.append([lab, t, b])
        c = next(c for c in arch["casings"] if c["key"] == k)
        ends = {c["top"], c["shoe"], arch["wellhead"], arch["td"]}  # limites físicos: sem recuo
        res = []
        for lab, t, b in merged:
            t2 = t if any(abs(t - e) < 1e-6 for e in ends) else t + inset
            b2 = b if any(abs(b - e) < 1e-6 for e in ends) else b - inset
            res.append({"curve": lab, "md_from": round(t2, 3), "md_to": round(b2, 3)})
        out[k] = {"wellcat_match": c["wellcat_match"], "pysafe": res}
    return out


def describe(arch):
    """Linhas legíveis da correspondência (para o deck e para o terminal)."""
    rows = []
    for k, v in pressure_map(arch, inset=0).items():
        for s in v["pysafe"]:
            rows.append((k, s["curve"], s["md_from"], s["md_to"]))
    return rows


if __name__ == "__main__":
    import yaml
    from pathlib import Path
    cfg = yaml.safe_load((Path(__file__).resolve().parents[1] / "loadcases.yaml").read_text())
    arch = cfg["architecture"]
    for bd in bands(arch):
        print(f"{bd['top']:7.1f}–{bd['base']:7.1f}", bd["present"], [r[0] for r in bd["regions"]])
    for r in describe(arch):
        print(r)
