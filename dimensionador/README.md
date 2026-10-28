# Dimensionador de Revestimento (offline)

Um arquivo só, `dimensionador-revestimento.html`, com o mesmo motor do WellPlatform
(`packages/core/src/dimensionamento`, motor 1.0.0, commit `ff93a74`). Abre com
duplo clique no Edge ou no Chrome. Não precisa de internet e nenhum arquivo sai do computador.

## Uso

| Passo | O que fazer |
|---|---|
| 1 | Escolher o **projeto do SCORE** (`project*.json`) |
| 2 | Escolher os **carregamentos do SCORE** (`Carregamentos*.xlsx`, abas `F2_Hid - P.Int`…) |
| 3 | Conferir as unidades da planilha (padrão do SCORE: m, psi, lbf, °C) |
| 4 | Ajustar as regras (trecho mínimo, máximo de trechos, OD alternativo, proteção do sal) e, se quiser, os FS |
| 5 | **Dimensionar**: fase por fase, a coluna modelada × a recomendada pelo cardápio 27-29, o menor FS por modo, onde reprova e os ajustes sugeridos (pressão de teste máxima, colapso no sal) |
| 6 | **Baixar colunas recomendadas (.csv)**, se precisar levar para outra planilha |

## Limites

- Motor de poço vertical (avisa se a trajetória do projeto for inclinada).
- Casos de sobrevivência que não estiverem na planilha (shut-in e fluxo descontrolado do WCA) não entram.
- Confirmar a configuração final no SCORE antes de usar em projeto.

## Atualizar

Gerado a partir do WellPlatform. Para refazer: `bun build` de um ponto de entrada que expõe
`lerPasta`, `lerCarregamentosDoScore`, `paraUnidadeInterna`, `lerProjetoDoScore`, `dimensionarPoco` e
`CARDAPIO_27_29` como `window.Motor`, embutido no HTML.
