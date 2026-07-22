import * as XLSX from "xlsx";

/** Gera planilhas Trust Excel sob demanda no navegador. */

function aoaToSheet(aoa, cols = {}) {
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  if (cols.widths) {
    ws["!cols"] = cols.widths.map((w) => ({ wch: w }));
  }
  if (cols.merges) ws["!merges"] = cols.merges;
  return ws;
}

function downloadWorkbook(wb, filename) {
  XLSX.writeFile(wb, filename, { bookType: "xlsx" });
}

function todayISO() {
  const d = new Date();
  return d.toISOString().slice(0, 10);
}

function addDays(iso, n) {
  const d = new Date(iso + "T12:00:00");
  d.setDate(d.getDate() + n);
  return d.toISOString().slice(0, 10);
}

function monthStart() {
  const d = new Date();
  return new Date(d.getFullYear(), d.getMonth(), 1).toISOString().slice(0, 10);
}

// ---------- builders ----------

function buildFinanceiro(opts = {}) {
  const nome = opts.nome || "Usuário Trust";
  const cats = opts.categorias || [
    "Moradia", "Alimentação", "Transporte", "Lazer", "Assinaturas",
    "Saúde", "Educação", "Salário", "Freelance", "Investimentos", "Outros",
  ];
  const wb = XLSX.utils.book_new();

  const catAoa = [
    ["TRUST EXCEL · CATEGORIAS"],
    ["Lista-mestre para validação"],
    [],
    ["Categoria"],
    ...cats.map((c) => [c]),
  ];
  XLSX.utils.book_append_sheet(wb, aoaToSheet(catAoa, { widths: [22] }), "CATEGORIAS");

  const base = monthStart();
  const samples = opts.lancamentos || [
    [addDays(base, 1), "Salário", "Salário mensal", "Entrada", opts.renda || 9200],
    [addDays(base, 2), "Moradia", "Aluguel + condomínio", "Saída", 2450],
    [addDays(base, 3), "Alimentação", "Mercado", "Saída", 780.4],
    [addDays(base, 4), "Assinaturas", "Softwares", "Saída", 219.9],
    [addDays(base, 5), "Transporte", "Combustível / app", "Saída", 360],
    [addDays(base, 6), "Freelance", "Projeto cliente", "Entrada", 1800],
    [addDays(base, 7), "Lazer", "Restaurante", "Saída", 240],
    [addDays(base, 8), "Saúde", "Farmácia", "Saída", 145.5],
    [addDays(base, 12), "Investimentos", "Aporte", "Saída", 800],
  ];
  const lanAoa = [
    ["TRUST EXCEL · LANÇAMENTOS"],
    [`Titular: ${nome}`],
    [],
    ["Data", "Categoria", "Descrição", "Tipo", "Valor"],
    ...samples,
    ...Array.from({ length: 20 }, () => ["", "", "", "", ""]),
  ];
  XLSX.utils.book_append_sheet(wb, aoaToSheet(lanAoa, { widths: [12, 16, 28, 12, 14] }), "LANCAMENTOS");

  const painelRows = [
    ["TRUST EXCEL · DASHBOARD FINANCEIRO"],
    ["Totais do mês · ranking · barras visuais"],
    [],
    ["Mês de referência", base],
    [],
    ["TOTAL ENTRADAS", '=SUMIF(LANCAMENTOS!D:D,"Entrada",LANCAMENTOS!E:E)'],
    ["TOTAL SAÍDAS", '=SUMIF(LANCAMENTOS!D:D,"Saída",LANCAMENTOS!E:E)'],
    ["SALDO DO MÊS", "=B6-B7"],
    ["TAXA DE GASTO", '=IFERROR(B7/B6,0)'],
    [],
    ["Categoria", "Total saídas", "% do gasto", "Barra"],
  ];
  cats.forEach((c, i) => {
    const r = 12 + i;
    painelRows.push([
      c,
      `=SUMIFS(LANCAMENTOS!E:E,LANCAMENTOS!B:B,A${r},LANCAMENTOS!D:D,"Saída")`,
      `=IFERROR(B${r}/$B$7,0)`,
      `=IF(C${r}=0,"",REPT("█",ROUNDDOWN(C${r}*20,0))&REPT("░",20-ROUNDDOWN(C${r}*20,0)))`,
    ]);
  });
  painelRows.push([]);
  painelRows.push([
    "Insight",
    '=CONCATENATE("Maior gasto: ",INDEX(A12:A22,MATCH(MAX(B12:B22),B12:B22,0))," (",TEXT(MAX(B12:B22),"R$ #,##0.00"),")")',
  ]);
  painelRows.push(["Automação TrustCorp: preencha LANCAMENTOS; o dashboard recalcula sozinho."]);
  XLSX.utils.book_append_sheet(wb, aoaToSheet(painelRows, { widths: [22, 18, 12, 28] }), "DASHBOARD");
  downloadWorkbook(wb, "01_controle_financeiro_pessoal.xlsx");
}

function buildFluxo(opts = {}) {
  const wb = XLSX.utils.book_new();
  const t = todayISO();
  const saldoInicial = opts.saldoInicial ?? 15000;
  const rows = opts.movimentos || [
    [addDays(t, -14), "Recebimento cliente", 4200, "", "Realizado"],
    [addDays(t, -10), "Despesa operacional", "", 980, "Realizado"],
    [addDays(t, -6), "Recebimento projeto", 3100, "", "Realizado"],
    [addDays(t, -2), "Folha / estrutura", "", 3900, "Realizado"],
    [addDays(t, 2), "Recebimento previsto", 4800, "", "Previsto"],
    [addDays(t, 5), "Aluguel", "", 2400, "Previsto"],
    [addDays(t, 8), "Impostos", "", 1250, "Previsto"],
    [addDays(t, 12), "Contrato novo", 2700, "", "Previsto"],
    [addDays(t, 18), "Fornecedor", "", 760, "Previsto"],
    [addDays(t, 25), "Entrada recorrente", 2200, "", "Previsto"],
  ];
  const mov = [
    ["TRUST EXCEL · MOVIMENTOS"],
    ["Realizado + previsto · saldo acumulado"],
    [],
    ["Data", "Descrição", "Entrada", "Saída", "Saldo acumulado", "Status"],
  ];
  rows.forEach((row, i) => {
    const r = 5 + i;
    const formula =
      i === 0
        ? `=${saldoInicial}+IF(C${r}="",0,C${r})-IF(D${r}="",0,D${r})`
        : `=E${r - 1}+IF(C${r}="",0,C${r})-IF(D${r}="",0,D${r})`;
    mov.push([row[0], row[1], row[2] || "", row[3] || "", formula, row[4]]);
  });
  for (let i = 0; i < 15; i++) mov.push(["", "", "", "", "", ""]);
  XLSX.utils.book_append_sheet(wb, aoaToSheet(mov, { widths: [12, 26, 12, 12, 16, 12] }), "MOVIMENTOS");

  const dash = [
    ["TRUST EXCEL · DASHBOARD FLUXO DE CAIXA"],
    ["Saldo · projeção 30 dias · alerta"],
    [],
    ["SALDO HOJE", '=LOOKUP(2,1/(MOVIMENTOS!F5:F200="Realizado"),MOVIMENTOS!E5:E200)'],
    ["ENTRADAS PREVISTAS 30D", '=SUMIFS(MOVIMENTOS!C:C,MOVIMENTOS!F:F,"Previsto")'],
    ["SAÍDAS PREVISTAS 30D", '=SUMIFS(MOVIMENTOS!D:D,MOVIMENTOS!F:F,"Previsto")'],
    ["SALDO PROJETADO 30D", "=B4+B5-B6"],
    [],
    ["ALERTA", '=IF(B7<0,"⚠ Caixa negativo em 30 dias","✓ Projeção saudável")'],
    [],
    ["Automação: lance realizados e previstos em MOVIMENTOS. O dashboard atualiza sozinho."],
  ];
  XLSX.utils.book_append_sheet(wb, aoaToSheet(dash, { widths: [28, 40] }), "DASHBOARD");
  downloadWorkbook(wb, "02_fluxo_de_caixa_simples.xlsx");
}

function buildMetas(opts = {}) {
  const wb = XLSX.utils.book_new();
  const metas = opts.metas || [
    ["Faturar R$ 30 mil/mês", 30000, 19200, "2026-12-31", "Liberdade financeira"],
    ["Reserva de emergência", 20000, 11000, "2026-10-01", "Segurança"],
    ["Reduzir assinaturas", 200, 95, "2026-08-31", "Estancar vazamento"],
    ["Ler 12 livros", 12, 6, "2026-12-31", "Clareza"],
    ["Fechar 8 clientes", 8, 4, "2026-12-31", "Crescimento"],
  ];
  const mAoa = [
    ["TRUST EXCEL · METAS"],
    ["Progresso automático com barra visual"],
    [],
    ["Meta", "Valor-alvo", "Realizado", "% Progresso", "Barra", "Prazo", "Por quê"],
  ];
  metas.forEach((m, i) => {
    const r = 5 + i;
    mAoa.push([
      m[0], m[1], m[2],
      `=IF(B${r}=0,0,C${r}/B${r})`,
      `=REPT("█",ROUNDDOWN(D${r}*12,0))&REPT("░",12-ROUNDDOWN(D${r}*12,0))`,
      m[3], m[4],
    ]);
  });
  XLSX.utils.book_append_sheet(wb, aoaToSheet(mAoa, { widths: [28, 12, 12, 12, 18, 12, 20] }), "METAS");

  const habits = opts.habitos || ["Treinar", "Ler 20 min", "Prospecção", "Registrar finanças", "Dormir 7h"];
  const hAoa = [
    ["TRUST EXCEL · HÁBITOS"],
    ["Marque x · consistência automática"],
    [],
    ["Hábito", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb", "Dom", "% Semana", "Status"],
  ];
  habits.forEach((h, i) => {
    const r = 5 + i;
    const marks = [1, 1, 0, 1, 1, 0, 1].map((x) => (x ? "x" : ""));
    hAoa.push([
      h, ...marks,
      `=COUNTIF(B${r}:H${r},"x")/7`,
      `=IF(I${r}>=0.8,"🔥 Consistente",IF(I${r}>=0.5,"Em progresso","⚠ Frágil"))`,
    ]);
  });
  XLSX.utils.book_append_sheet(wb, aoaToSheet(hAoa, { widths: [18, 5, 5, 5, 5, 5, 5, 5, 10, 16] }), "HABITOS");

  const dash = [
    ["TRUST EXCEL · DASHBOARD METAS & HÁBITOS"],
    [],
    ["METAS ATIVAS", "=COUNTA(METAS!A5:A50)"],
    ["PROGRESSO MÉDIO", "=IFERROR(AVERAGE(METAS!D5:D50),0)"],
    ["HÁBITOS", "=COUNTA(HABITOS!A5:A50)"],
    ["CONSISTÊNCIA", "=IFERROR(AVERAGE(HABITOS!I5:I50),0)"],
    [],
    ["Ritual: 10 minutos no domingo. Atualize Realizado e marque hábitos."],
  ];
  XLSX.utils.book_append_sheet(wb, aoaToSheet(dash, { widths: [20, 18] }), "DASHBOARD");
  downloadWorkbook(wb, "03_metas_e_habitos.xlsx");
}

function buildCobrancas(opts = {}) {
  const wb = XLSX.utils.book_new();
  const t = todayISO();
  const samples = opts.cobrancas || [
    ["Cliente Alpha", "Mensalidade SaaS", 1800, addDays(t, -12), "Em aberto"],
    ["Cliente Beta", "Projeto site", 5200, addDays(t, -4), "Em aberto"],
    ["Cliente Gama", "Consultoria", 2400, addDays(t, 2), "Em aberto"],
    ["Cliente Delta", "Setup", 980, addDays(t, 9), "Em aberto"],
    ["Cliente Épsilon", "Suporte", 650, addDays(t, -1), "Pago"],
    ["Cliente Zeta", "Marketing", 3600, addDays(t, -18), "Em aberto"],
    ["Cliente Nova", "Implementação", 7800, addDays(t, -7), "Em aberto"],
  ];
  const aoa = [
    ["TRUST EXCEL · COBRANÇAS"],
    ["Semáforo, atraso e ação sugerida"],
    [],
    ["Cliente", "Descrição", "Valor", "Vencimento", "Situação", "Status automático", "Dias atraso", "Ação sugerida"],
  ];
  samples.forEach((s, i) => {
    const r = 5 + i;
    aoa.push([
      s[0], s[1], s[2], s[3], s[4],
      `=IF(E${r}="Pago","Pago",IF(D${r}<TODAY(),"ATRASADO",IF(D${r}<=TODAY()+3,"Vence em breve","No prazo")))`,
      `=IF(OR(E${r}="Pago",D${r}>=TODAY()),0,TODAY()-D${r})`,
      `=IF(F${r}="ATRASADO","Cobrar hoje",IF(F${r}="Vence em breve","Lembrete amigável",IF(F${r}="Pago","—","Acompanhar")))`,
    ]);
  });
  for (let i = 0; i < 20; i++) aoa.push(["", "", "", "", "", "", "", ""]);
  XLSX.utils.book_append_sheet(wb, aoaToSheet(aoa, { widths: [16, 20, 12, 12, 12, 16, 12, 18] }), "COBRANCAS");

  const dash = [
    ["TRUST EXCEL · DASHBOARD COBRANÇAS"],
    [],
    ["EM ABERTO", '=SUMIF(COBRANCAS!E:E,"Em aberto",COBRANCAS!C:C)'],
    ["ATRASADO", '=SUMIFS(COBRANCAS!C:C,COBRANCAS!E:E,"Em aberto",COBRANCAS!D:D,"<"&TODAY())'],
    ["QTD ATRASADAS", '=COUNTIFS(COBRANCAS!E:E,"Em aberto",COBRANCAS!D:D,"<"&TODAY())'],
    ["A RECEBER 7 DIAS", '=SUMIFS(COBRANCAS!C:C,COBRANCAS!E:E,"Em aberto",COBRANCAS!D:D,">="&TODAY(),COBRANCAS!D:D,"<="&TODAY()+7)'],
    [],
    ["PRIORIDADE", '=CONCATENATE("Há ",B5," cobrança(s) atrasada(s) somando ",TEXT(B4,"R$ #,##0.00"),". Foque nos vermelhos.")'],
    ["Ritual: toda segunda, filtrar ATRASADO e cobrar."],
  ];
  XLSX.utils.book_append_sheet(wb, aoaToSheet(dash, { widths: [20, 70] }), "DASHBOARD");
  downloadWorkbook(wb, "04_controle_de_cobrancas.xlsx");
}

function buildPrecificacao(opts = {}) {
  const wb = XLSX.utils.book_new();
  const renda = opts.renda || 20000;
  const custos = opts.custosFixos || 4200;
  const horas = opts.horasMes || 110;
  const horasProjeto = opts.horasProjeto || 24;
  const custosVar = opts.custosVar || 520;
  const alvo = opts.precoAlvo || 8500;

  const tempo = [
    ["TRUST EXCEL · VALOR DA HORA"],
    [],
    ["Renda desejada / mês", renda],
    ["Custos fixos mensais", custos],
    ["Horas produtivas / mês", horas],
    [],
    ["VALOR DA HORA", "=(B3+B4)/B5"],
    [],
    ["Use 100–120h produtivas, não 160h."],
  ];
  XLSX.utils.book_append_sheet(wb, aoaToSheet(tempo, { widths: [28, 16] }), "MEU_TEMPO");

  const dash = [
    ["TRUST EXCEL · DASHBOARD PRECIFICAÇÃO"],
    ["3 cenários automáticos"],
    [],
    ["Serviço", opts.servico || "Consultoria / projeto sob demanda"],
    ["Horas estimadas", horasProjeto],
    ["Custos variáveis", custosVar],
    ["Valor da hora", "=MEU_TEMPO!B7"],
    ["Custo total", "=(B5*B7)+B6"],
    [],
    ["Cenário", "Margem", "Preço sugerido", "Quando usar", "Barra"],
    ["Mínimo", 0.15, "=$B$8/(1-B11)", "Entrar na conta", '=REPT("█",ROUND(B11*20,0))'],
    ["Ideal", 0.3, "=$B$8/(1-B12)", "Padrão saudável", '=REPT("█",ROUND(B12*20,0))'],
    ["Premium", 0.5, "=$B$8/(1-B13)", "Urgência / especialização", '=REPT("█",ROUND(B13*20,0))'],
    [],
    ["Preço-alvo do cliente", alvo],
    ["Horas máximas p/ fechar (margem 30%)", "=MAX(0,((B15*(1-0.3))-B6)/B7)"],
    [],
    ["Regra: não baixe preço — remova escopo."],
  ];
  XLSX.utils.book_append_sheet(wb, aoaToSheet(dash, { widths: [36, 14, 16, 28, 14] }), "DASHBOARD");
  downloadWorkbook(wb, "05_precificacao.xlsx");
}

function buildKanban(opts = {}) {
  const wb = XLSX.utils.book_new();
  const t = todayISO();
  const tasks = opts.tarefas || [
    ["Enviar proposta Cliente A", "Flávio", "Alta", "Fazendo", addDays(t, -1), ""],
    ["Fechar conciliação", "Ana", "Alta", "A fazer", addDays(t, 1), ""],
    ["Publicar campanha", "MKT", "Média", "A fazer", addDays(t, 3), ""],
    ["Revisar contratos", "Jurídico", "Alta", "Fazendo", t, ""],
    ["Organizar pipeline CRM", "Comercial", "Média", "Feito", addDays(t, -2), addDays(t, -2)],
    ["Preparar relatório", "Ops", "Baixa", "A fazer", addDays(t, 4), ""],
    ["Cobrar top 5", "Financeiro", "Alta", "A fazer", addDays(t, -3), ""],
  ];
  const aoa = [
    ["TRUST EXCEL · TAREFAS KANBAN"],
    [],
    ["Tarefa", "Responsável", "Prioridade", "Status", "Prazo", "Alerta", "Conclusão", "Notas"],
  ];
  tasks.forEach((task, i) => {
    const r = 4 + i;
    aoa.push([
      task[0], task[1], task[2], task[3], task[4],
      `=IF(A${r}="","",IF(D${r}="Feito","",IF(E${r}<TODAY(),"ATRASADA",IF(E${r}=TODAY(),"HOJE",""))))`,
      task[5] || "",
      "",
    ]);
  });
  for (let i = 0; i < 20; i++) aoa.push(["", "", "", "", "", "", "", ""]);
  XLSX.utils.book_append_sheet(wb, aoaToSheet(aoa, { widths: [28, 14, 12, 12, 12, 12, 12, 20] }), "TAREFAS");

  const dash = [
    ["TRUST EXCEL · DASHBOARD KANBAN"],
    [],
    ["A FAZER", '=COUNTIF(TAREFAS!D:D,"A fazer")'],
    ["FAZENDO", '=COUNTIF(TAREFAS!D:D,"Fazendo")'],
    ["FEITO", '=COUNTIF(TAREFAS!D:D,"Feito")'],
    ["ATRASADAS", '=COUNTIFS(TAREFAS!D:D,"<>Feito",TAREFAS!E:E,"<"&TODAY())'],
    ["WIP STATUS", '=IF(B4>3,"⚠ Acima do limite","✓ Dentro do limite")'],
    [],
    ["Regra de ouro: no máximo 3 em Fazendo."],
  ];
  XLSX.utils.book_append_sheet(wb, aoaToSheet(dash, { widths: [16, 28] }), "DASHBOARD");
  downloadWorkbook(wb, "06_gestao_de_tarefas_kanban.xlsx");
}

function buildRelatorio(opts = {}) {
  const wb = XLSX.utils.book_new();
  const t = todayISO();
  const dados = [
    ["TRUST EXCEL · DADOS DO DIA"],
    ["2 minutos por dia · o resumo se escreve sozinho"],
    [],
    ["Data", "Categoria/Cliente", "Métrica", "Valor"],
  ];
  const clients = opts.clientes || ["Cliente A", "Cliente B", "Cliente C", "Inbound"];
  let seed = 11;
  const rnd = () => {
    seed = (seed * 1103515245 + 12345) & 0x7fffffff;
    return seed / 0x7fffffff;
  };
  for (let daysAgo = 14; daysAgo >= 0; daysAgo--) {
    const d = addDays(t, -daysAgo);
    const dt = new Date(d + "T12:00:00");
    if (dt.getDay() === 0 || dt.getDay() === 6) continue;
    const client = clients[Math.floor(rnd() * clients.length)];
    dados.push([d, client, "Vendas", Math.round(900 + rnd() * 4300)]);
    dados.push([d, client, "Atendimentos", Math.round(3 + rnd() * 17)]);
    dados.push([d, "Operação", "Entregas", Math.round(1 + rnd() * 7)]);
  }
  for (let i = 0; i < 20; i++) dados.push(["", "", "", ""]);
  XLSX.utils.book_append_sheet(wb, aoaToSheet(dados, { widths: [12, 18, 14, 12] }), "DADOS");

  const dash = [
    ["TRUST EXCEL · RELATÓRIO AUTOMÁTICO"],
    ["100% fórmulas · copie e envie em 1 minuto"],
    [],
    ["VENDAS SEMANA", '=SUMIFS(DADOS!D:D,DADOS!C:C,"Vendas",DADOS!A:A,">="&TODAY()-WEEKDAY(TODAY(),2)+1)'],
    ["SEMANA ANTERIOR", '=SUMIFS(DADOS!D:D,DADOS!C:C,"Vendas",DADOS!A:A,">="&TODAY()-WEEKDAY(TODAY(),2)+1-7,DADOS!A:A,"<"&TODAY()-WEEKDAY(TODAY(),2)+1)'],
    ["VARIAÇÃO", "=IFERROR((B4-B5)/B5,0)"],
    ["ATENDIMENTOS", '=SUMIFS(DADOS!D:D,DADOS!C:C,"Atendimentos",DADOS!A:A,">="&TODAY()-WEEKDAY(TODAY(),2)+1)'],
    ["ENTREGAS", '=SUMIFS(DADOS!D:D,DADOS!C:C,"Entregas",DADOS!A:A,">="&TODAY()-WEEKDAY(TODAY(),2)+1)'],
    ["DESTAQUE", '=IFERROR(INDEX(DADOS!B:B,MATCH(MAXIFS(DADOS!D:D,DADOS!C:C,"Vendas",DADOS!A:A,">="&TODAY()-WEEKDAY(TODAY(),2)+1),DADOS!D:D,0)),"—")'],
    [],
    ["TEXTO PRONTO"],
    ['=CONCATENATE("Semana fechou em R$ ",TEXT(B4,"#,##0.00")," (",TEXT(B6,"+0.0%;-0.0%")," vs anterior). Atendimentos: ",TEXT(B7,"0")," · Entregas: ",TEXT(B8,"0"),". Destaque: ",B9,".")'],
    [],
    ["Ritual sexta 16h: abrir, conferir, copiar o texto, enviar."],
  ];
  XLSX.utils.book_append_sheet(wb, aoaToSheet(dash, { widths: [22, 90] }), "DASHBOARD");
  downloadWorkbook(wb, "07_relatorio_que_se_monta_sozinho.xlsx");
}

export const GENERATORS = {
  "financeiro-pessoal": buildFinanceiro,
  "fluxo-caixa": buildFluxo,
  "metas-habitos": buildMetas,
  cobrancas: buildCobrancas,
  precificacao: buildPrecificacao,
  "tarefas-kanban": buildKanban,
  "relatorio-automatico": buildRelatorio,
};

export function generatePlanilha(id, opts = {}) {
  const fn = GENERATORS[id];
  if (!fn) throw new Error(`Gerador não encontrado: ${id}`);
  fn(opts);
  return true;
}
