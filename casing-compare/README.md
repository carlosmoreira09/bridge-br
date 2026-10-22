# Comparativo SCORE-PySafe × WellCat

Pipeline para comparar temperatura e pressão interna por revestimento entre SCORE-PySafe e WellCat
e gerar o deck `.pptx`. Detalhes de método, convenções e como adicionar carregamentos: `CONTEXT.md`.

| Pasta | Conteúdo |
|---|---|
| `data/raw/` | Planilhas de entrada (uma por carregamento) |
| `out/` | Decks, planilhas preenchidas e CSVs gerados |
| `src/` | `extract.py`, `architecture.py`, `build_deck.py` |

```bash
pip install pandas numpy openpyxl pyyaml python-pptx
python3 src/build_deck.py
```
