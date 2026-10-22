# CONTEXT — Comparativo SCORE-PySafe × WellCat (Projeto de Revestimento)

Pipeline reprodutível que lê a planilha de cada carregamento (loadcase), compara
**WellCat × SCORE-PySafe** revestimento a revestimento (temperatura e pressão interna)
e gera um deck `.pptx` com gráficos nativos (editáveis no PowerPoint).

## 1. Estrutura

| Caminho | Conteúdo |
|---|---|
| `loadcases.yaml` | Registro dos carregamentos, revestimentos, mapa de pressão, textos de leitura |
| `data/raw/*.xlsx` | Planilhas de entrada (uma por carregamento) |
| `src/extract.py` | Leitura das 4 abas → tabela "tidy" + cálculo de Δ |
| `src/build_deck.py` | Gera o deck consolidado (todos os loadcases do YAML) |
| `out/` | `.pptx`, `<id>_tidy.csv` (dados normalizados), `<id>_stats.csv` (estatísticas), `<id>_temperaturas_preenchidas.xlsx` |

Dependências: `python3`, `pandas`, `numpy`, `openpyxl`, `pyyaml`, `python-pptx`.

## 2. Formato esperado da planilha (4 abas)

| Aba | Estrutura | Observação |
|---|---|---|
| Dados Térmicos SCORE-PySafe | `Cota`, `MD`, `Geotérmico`, `Fluido`, `Anular A/B`, `COP/COI`, `Liner 10 3/4"`, `14"`, `22"`, `36"` (x2) | Contínuo; cada coluna termina na sapata da fase |
| Dados Térmicos WellCat | Mesmo cabeçalho; MD comum, colunas esparsas | Cada coluna tem pontos em MDs próprios (NaN nos demais) |
| Dados Pressão PySafe | `MD`, `Fluido`, `Anular A`, `Anular B` | Curvas contínuas, sem rótulo de revestimento |
| Dados Pressão WellCat | Blocos lado a lado `[idx, MD, P]`, 2 linhas de cabeçalho | MD duplicado em degraus (ex.: 3024,96 / 3025,02) |

Nomes de aba são localizados por correspondência aproximada (ignora acento/espaço).
Colunas térmicas são achadas pelo trecho `match` de cada revestimento (`22"`, `14"`, `10 3/4"`, `COP/COI`).
Condutor 36" fica de fora (não está em `strings`).

## 3. Regras de comparação

| Item | Regra |
|---|---|
| Δ | `WellCat − PySafe`, avaliado nas profundidades do WellCat dentro da faixa comum |
| Interpolação | PySafe interpolado linearmente; MDs repetidos → mantém o último valor |
| Degraus | Interpolação nunca atravessa troca de curva PySafe (campo `seg`, ex.: packer) |
| Média | Δ médio ponderado pela profundidade (trapézios), pois o grid não é uniforme |
| Cobertura | % medido = faixa com dado WellCat original ÷ trecho PySafe do revestimento |
| Pressão interna | Derivada de `architecture` (ver abaixo); `pressure_map` explícito no loadcase sobrescreve |

### Convenção de nomes: PySafe × WellCat

O SCORE-PySafe nomeia **espaços**; o WellCat nomeia **revestimentos**. A pressão interna de cada
revestimento é a curva do espaço imediatamente dentro dele, trecho a trecho:

| Espaço PySafe | Definição |
|---|---|
| Fluido | Interior da COP/COI; abaixo do packer, interior do revestimento mais interno |
| Anular A | Entre a COP/COI e o 1º revestimento presente (de baixo para cima: liner 10 3/4" → 14") |
| Anular B | Entre o 1º e o 2º limite, **só acima do TOC** do revestimento interno: poço aberto (fluido de perfuração) até a sapata do 22", depois interior do 22" |
| Cimento | Abaixo do TOC — não há anular, não há comparação |

O mapa é gerado por `src/architecture.py` a partir de `architecture:` no YAML (cabeça, packer, TD,
e para cada revestimento: topo, sapata, TOC). Para outro poço/fase basta ajustar a arquitetura.
`python3 src/architecture.py` imprime as faixas e o mapa resultante.

Resultado (Orion Firme 01), conferido nos dados — a 2.500 m: 22" WellCat 6.721 psi × Anular B 6.750 psi × Anular A 10.774 psi:

| Revestimento | Curva PySafe | MD (m) |
|---|---|---|
| 22" | Anular B | 1.785–3.425 |
| 14" | Anular A | 1.785–6.050 (topo do liner) |
| Liner 10 3/4" | Anular A + Fluido | 6.050–6.317 (packer) + 6.317–6.732 |

Limites internos entre trechos (topo do liner, packer) recebem recuo de 0,05 m para não interpolar sobre o degrau.
Anular B também é a pressão externa do 14" — fora da comparação porque o WellCat exportado só traz pressão interna.
TOC do liner assumido no topo (6.050 m).

## 3.1 Preenchimento térmico (`temp_fill` no YAML)

A aba térmica do WellCat é esparsa. Antes de comparar, cada revestimento é levado a um grid único
(MDs do PySafe do topo à sapata ∪ MDs do WellCat), gerando `kind = temp_fill` com a coluna `src`:

| `src` | Regra |
|---|---|
| medido | Valor original da planilha |
| interpolado | Reta entre os dois pontos medidos vizinhos (lacuna interna) |
| estimado | Fora da faixa medida do WellCat, pelo método configurado |

| Método | Uso | Regra |
|---|---|---|
| `linear_tail` (padrão) | Extensões curtas (liner 10 3/4" de 6.050 a 6.079 m) | Reta pelos `n` pontos da ponta, só até `max_len_m` (300 m) |
| `anchor` | 14" e COP/COI (dado só até ~2.400 m) | Reta do último ponto medido até o 1º ponto medido de `ref_string` (liner 10 3/4"); abaixo disso segue o liner |

Por que não reta pura no 14" e na COP/COI: a tendência rasa extrapolada até 6.150–6.317 m dá 140–184 °C,
acima do fluido produzido (~113–118 °C), fisicamente impossível. Revisar `ref_string` se a arquitetura mudar.

Nos slides: WellCat medido em linha laranja com marcadores; estimado em pontilhado laranja-claro.
Indicadores mostram Δ médio só no trecho medido e no trecho completo. A planilha
`<id>_temperaturas_preenchidas.xlsx` traz uma aba por revestimento (MD, PySafe, WellCat, origem, Δ)
com células interpoladas em amarelo e estimadas em salmão.

## 4. Como adicionar um novo carregamento

1. Copiar a planilha para `data/raw/<nome>.xlsx`.
2. Acrescentar um bloco em `loadcases:` no `loadcases.yaml`:
   ```yaml
   - id: inj_agua_max            # sem espaços
     name: "Injeção de Água – Máxima"
     short: "Inj. Água Máx."
     file: "data/raw/<nome>.xlsx"
     sheets: {temp_pysafe: "...", temp_wellcat: "...", press_pysafe: "...", press_wellcat: "..."}
     # architecture: {...}       # só se a arquitetura (sapatas, TOC, packer) mudar
     observations: ["...", "...", "..."]   # opcional, até 3 (slide Resumo)
     next_steps: ["..."]                    # opcional (slide final)
   ```
3. Rodar `python3 src/extract.py` → conferir contagens/faixas impressas e `out/<id>_stats.csv`.
4. Escrever `observations` a partir de `out/<id>_stats.csv` (Δ médio, Δ máx e MD, cobertura).
5. Rodar `python3 src/build_deck.py` → deck único com todos os loadcases.
   - Com 2+ loadcases entra automaticamente o slide **Consolidado – Δ máximo por carregamento**.
   - `--only <id> [<id>...]` gera só os escolhidos; `--out caminho.pptx` muda o destino.

## 5. Estrutura do deck (por loadcase)

| # | Slide |
|---|---|
| 1 | Capa (poço + carregamentos) |
| 2 | Resumo da comparação (tabela Δ + leitura) |
| 3 | Escopo, cobertura de dados e premissas |
| 4 | Esquemático PySafe (espaços) × WellCat (revestimentos) |
| 5–11 | Temperatura: divisória, visão geral, 22", 14", liner 10 3/4", COP/COI, perfil ΔT |
| 12–17 | Pressão interna: divisória, visão geral, 22", 14", liner 10 3/4", perfil ΔP |
| fim | Consolidado (se 2+) e Pontos a confirmar |

Cada slide de revestimento: 3 gráficos (PySafe | WellCat | Sobreposição) com mesma escala,
MD no eixo vertical crescendo para baixo, e 4 indicadores (máx PySafe, máx WellCat, Δ médio, cobertura).
Notas do apresentador trazem a leitura automática e as faixas de MD.

## 6. Convenções visuais

| Elemento | Cor |
|---|---|
| SCORE-PySafe | `#0F7C8C` (linha cheia) |
| WellCat | `#D9822B` (tracejada + marcadores quando < 40 pontos) |
| Geotérmico | `#9AA5AE` tracejado |
| 22" / 14" / Liner / COP | `#34495E` / `#6C5B9E` / `#2E8B57` / `#B23A48` |

Fonte: Cambria (títulos), Calibri (corpo). Formato numérico pt-BR (1.785,0).

## 7. Cuidados conhecidos

- O LibreOffice descarta séries XY mais curtas que as seguintes; `depth_chart` repete o último ponto para igualar comprimentos (inócuo no PowerPoint).
- WellCat térmico deste loadcase cobre só o trecho raso do 14" (até 2.494 m) e da COP/COI (até 2.359 m); o restante é estimado (ver 3.1).
- Pressão WellCat do 14" abaixo de 6.050 m (topo do liner) não entra no Δ: PySafe não tem curva equivalente.
- Validar o `.pptx` com `validate.py` da skill pptx após mudanças no gerador.

## 8. Prompt sugerido para o Claude Code CLI

```
Leia CONTEXT.md. Adicionei data/raw/<arquivo>.xlsx (carregamento <nome>).
1) Registre o loadcase no loadcases.yaml (confira nomes das abas).
2) Rode src/extract.py e mostre contagens/faixas por revestimento.
3) Se a arquitetura mudou (sapatas, topo do liner, packer, TOC), ajuste architecture e rode src/architecture.py.
4) Escreva 3 observations factuais a partir de out/<id>_stats.csv.
5) Gere o deck consolidado com src/build_deck.py e valide.
```
