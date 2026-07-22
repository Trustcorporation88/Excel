# Trust Excel · Agente + 7 Ferramentas Premium

App: https://excel.trustcorp.com.br

## O que é
1. **Agente Excel** — 8 modos (insights, fórmulas, limpeza, comparação, etc.)
2. **7 ferramentas premium** — cada uma com cockpit executivo, gráficos nativos do Excel,
   listas com autocomplete (validação de dados), formatação condicional e motores de cálculo.

## As 7 ferramentas (`/public/planilhas/`)
| # | Arquivo | Inteligência |
|---|---------|--------------|
| 1 | `01_diagnostico_financeiro_360.xlsx` | Score de saúde 0–100, projeção de patrimônio, radar, doughnut, alertas de teto |
| 2 | `02_simulador_fluxo_caixa_90d.xlsx` | Cenários (pessimista/realista/otimista), runway em dias, ponto de ruptura |
| 3 | `03_painel_okr_habitos.xlsx` | Previsão de atingimento por ritmo, burn-up de OKRs, radar de hábitos |
| 4 | `04_regua_cobranca_inteligente.xlsx` | DSO, aging por faixa, recuperação esperada por risco, ação sugerida |
| 5 | `05_motor_precificacao.xlsx` | Valor/hora, ponto de equilíbrio (gráfico), margem de contribuição, 3 cenários |
| 6 | `06_cockpit_produtividade.xlsx` | Throughput, cycle time, WIP, burndown, alertas de tarefas envelhecidas |
| 7 | `07_bi_executivo_negocio.xlsx` | MRR, crescimento, churn, CAC, LTV, LTV/CAC, leitura executiva automática |

Todas abrem já com dados de exemplo e recalculam sozinhas conforme o usuário preenche as abas de entrada.

## Geração dos modelos
Motor Python (openpyxl) em `../build_v3/` — `builders.py` + `_style.py`.
Gera gráficos nativos, validações e formatação condicional (recursos que o SheetJS do navegador não escreve).

## Dev / Build
```bash
npm install
npm run dev
npm run build
```
