# Trust Excel · Agente + 8 Ferramentas Premium

App: https://excel.trustcorp.com.br

## O que é
1. **Agente Excel** — 8 modos (insights, fórmulas, limpeza, comparação, etc.)
2. **8 ferramentas premium** — cockpit executivo, gráficos nativos, autocomplete, formatação condicional e motores de cálculo

## IA (OpenAI principal + DeepSeek fallback)

### Variáveis no Railway
No serviço **Excel** (`excel.trustcorp.com.br`), configure:

```bash
OPENAI_API_KEY=sk-proj-...
OPENAI_MODEL=gpt-4o-mini

DEEPSEEK_API_KEY=sk-...
DEEPSEEK_MODEL=deepseek-chat

NODE_ENV=production
PORT=3000
```

Ordem de uso:
1. tenta **OpenAI**
2. se falhar/ausente, cai para **DeepSeek**

### Local
```bash
cp .env.example .env
# edite .env com suas chaves
npm install
npm run dev
```

Health check:
- local: `http://localhost:5173/api/health` (dev) ou `http://localhost:3000/api/health` (prod)
- prod: `https://excel.trustcorp.com.br/api/health`

## As 8 ferramentas (`/public/planilhas/`)
| # | Arquivo | Inteligência |
|---|---------|--------------|
| 1 | `01_diagnostico_financeiro_360.xlsx` | Score 0–100, projeção, radar, doughnut, alertas |
| 2 | `02_simulador_fluxo_caixa_90d.xlsx` | Cenários, runway, ruptura |
| 3 | `03_painel_okr_habitos.xlsx` | Previsão por ritmo, OKRs, hábitos |
| 4 | `04_regua_cobranca_inteligente.xlsx` | DSO, aging, recuperação esperada |
| 5 | `05_motor_precificacao.xlsx` | Valor/hora, break-even, 3 cenários |
| 6 | `06_cockpit_produtividade.xlsx` | Throughput, cycle time, WIP, burndown |
| 7 | `07_bi_executivo_negocio.xlsx` | MRR, churn, CAC, LTV/CAC |
| 8 | `08_analise_cohort_retencao.xlsx` | Heatmap de retenção e cascata de evasão |

## Dev / Build / Start
```bash
npm install
npm run dev      # desenvolvimento (proxy IA no Vite)
npm run build    # gera dist/
npm start        # produção: Express serve dist + /api/ai/chat
```

No Railway, o start command deve ser:
```bash
npm run build && npm start
```
ou (se o build já roda no pipeline):
```bash
npm start
```
