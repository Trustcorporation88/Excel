# Trust Excel · Agente + 7 Planilhas + Gerador

App em https://excel.trustcorp.com.br

## O que tem
1. **Agente Excel** com 8 modos (insights, fórmulas, limpeza, comparação, etc.)
2. **7 planilhas prontas (v2)** com dashboards visuais, KPIs, barras e gráficos
3. **Gerador sob demanda** no navegador (personaliza e baixa na hora)

## Downloads estáticos
Arquivos em `/public/planilhas/`:

1. `01_controle_financeiro_pessoal.xlsx`
2. `02_fluxo_de_caixa_simples.xlsx`
3. `03_metas_e_habitos.xlsx`
4. `04_controle_de_cobrancas.xlsx`
5. `05_precificacao.xlsx`
6. `06_gestao_de_tarefas_kanban.xlsx`
7. `07_relatorio_que_se_monta_sozinho.xlsx`

## Gerador
Módulo: `src/planilhaGenerator.js`  
Gera workbooks com SheetJS no cliente a partir dos parâmetros do formulário.

## Dev
```bash
npm install
npm run dev
```

## Build
```bash
npm run build
```
