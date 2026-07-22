import { useState, useRef } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, Cell,
  XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid, Legend,
} from "recharts";

// ---------- design tokens ----------
const C = {
  bg: "#0C0A07",
  surface: "#161209",
  surface2: "#1E1810",
  line: "#2E2415",
  orange: "#FF7A1A",
  orangeDeep: "#E85D04",
  green: "#21A366",
  greenBright: "#7BE495",
  cream: "#F5EFE4",
  muted: "#A89C88",
};
const CHART_COLORS = ["#7BE495", "#FF7A1A", "#21A366", "#FFB25E", "#4ADE80", "#E85D04", "#A7F3D0", "#FDBA74"];


// ---------- 7 planilhas automatizadas (Trust Excel Services) ----------
const PLANILHA_SERVICES = [
  {
    id: "financeiro-pessoal",
    nome: "Controle Financeiro Pessoal",
    tag: "PLANILHA 1",
    desc: "Entradas, saídas e saldo do mês com painel e gráfico por categoria.",
    file: "/planilhas/01_controle_financeiro_pessoal.xlsx",
    tabs: "PAINEL · LANCAMENTOS · CATEGORIAS",
    auto: "Listas suspensas, totais do mês, gasto por categoria",
  },
  {
    id: "fluxo-caixa",
    nome: "Fluxo de Caixa Simples",
    tag: "PLANILHA 2",
    desc: "Saldo hoje + projeção de 30 dias com realizados e previstos.",
    file: "/planilhas/02_fluxo_de_caixa_simples.xlsx",
    tabs: "RESUMO · MOVIMENTOS",
    auto: "Saldo acumulado, projeção 30 dias, alerta de caixa",
  },
  {
    id: "metas-habitos",
    nome: "Metas e Hábitos do Ano",
    tag: "PLANILHA 3",
    desc: "Progresso de metas e consistência semanal de hábitos.",
    file: "/planilhas/03_metas_e_habitos.xlsx",
    tabs: "PAINEL · METAS · HABITOS",
    auto: "% progresso, barra visual, heatmap de hábitos",
  },
  {
    id: "cobrancas",
    nome: "Controle de Cobranças",
    tag: "PLANILHA 4",
    desc: "Semáforo de atraso, dias vencidos e fila de cobrança.",
    file: "/planilhas/04_controle_de_cobrancas.xlsx",
    tabs: "RESUMO · COBRANCAS",
    auto: "Status automático, dias de atraso, ação sugerida",
  },
  {
    id: "precificacao",
    nome: "Precificação",
    tag: "PLANILHA 5",
    desc: "Valor da hora e cenários mínimo, ideal e premium.",
    file: "/planilhas/05_precificacao.xlsx",
    tabs: "PRECO · MEU_TEMPO",
    auto: "Custo da entrega, margens e horas para preço-alvo",
  },
  {
    id: "tarefas-kanban",
    nome: "Gestão de Tarefas (Kanban)",
    tag: "PLANILHA 6",
    desc: "A fazer, fazendo e feito com alerta de atraso.",
    file: "/planilhas/06_gestao_de_tarefas_kanban.xlsx",
    tabs: "RESUMO · TAREFAS",
    auto: "Alertas, contadores kanban e limite WIP",
  },
  {
    id: "relatorio-automatico",
    nome: "Relatório que se Monta Sozinho",
    tag: "PLANILHA 7",
    desc: "Resumo semanal automático com texto pronto para copiar.",
    file: "/planilhas/07_relatorio_que_se_monta_sozinho.xlsx",
    tabs: "RESUMO · DADOS",
    auto: "Totais da semana, variação e destaque automático",
  },
];

// ---------- os 8 modos ----------
const MODES = [
  { id: 1, nome: "5 Insights", tag: "PROMPT 1", desc: "Padrões, anomalias e oportunidades em segundos.", icon: "⚡", dados: true },
  { id: 2, nome: "Fórmula Pronta", tag: "PROMPT 2", desc: "PROCV, SOMASES, datas — sem tentativa e erro.", icon: "ƒx", objetivo: true, placeholder: "Ex.: buscar o preço de um produto pelo código na aba Produtos" },
  { id: 3, nome: "Limpeza de Dados", tag: "PROMPT 3", desc: "Duplicados, vazios e erros de formato resolvidos.", icon: "✦", dados: true },
  { id: 4, nome: "Comparar 2 Planilhas", tag: "PROMPT 4", desc: "Divergências, ausências e diferenças de valor.", icon: "≠", dados: true, dados2: true },
  { id: 5, nome: "Explique a Planilha", tag: "PROMPT 5", desc: "Decifra a base herdada como se você fosse novo.", icon: "?", dados: true },
  { id: 6, nome: "Resumo Executivo", tag: "PROMPT 6", desc: "5 linhas, 3 achados, 2 riscos, 3 próximos passos.", icon: "▤", dados: true },
  { id: 7, nome: "Dashboard", tag: "PROMPT 7", desc: "KPIs, gráficos e layout antes do Power BI.", icon: "▦", dados: true },
  { id: 8, nome: "Automação", tag: "PROMPT 8", desc: "VBA, Power Query ou passo a passo pronto.", icon: "⟳", objetivo: true, dadosOpc: true, placeholder: "Ex.: todo mês eu junto 12 planilhas de filiais em uma só e removo duplicados" },
];

const DUVIDA_FLOWS = [
  { id: "sugerir", nome: "Sugerir o modo", desc: "Lê a dúvida e indica o melhor dos 8 prompts." },
  { id: "auto", nome: "Escolher e rodar", desc: "Escolhe o modo sozinho e já executa." },
  { id: "livre", nome: "Dúvida livre", desc: "Responde direto, sem obrigar um dos 8." },
];

const JSON_BASE = `Responda APENAS com um objeto JSON válido, sem markdown, sem crases, sem texto antes ou depois. Use exatamente esta estrutura (omita campos que não se aplicam, mas SEMPRE inclua "titulo" e "resumo"):
{
  "titulo": "título curto do resultado",
  "resumo": "resumo em linguagem simples, 3 a 5 frases, para quem não entende nada de Excel",
  "kpis": [{"label": "nome da métrica", "valor": "valor formatado", "detalhe": "contexto curto"}],
  "insights": ["insight em linguagem simples"],
  "riscos": ["risco ou alerta"],
  "proximos_passos": ["ação recomendada"],
  "graficos": [{"tipo": "bar|line|pie", "titulo": "título do gráfico", "dados": [{"nome": "categoria", "valor": 123}]}],
  "formulas": [{"objetivo": "o que faz", "formula_pv": "=FORMULA(A;B)", "formula_virgula": "=FORMULA(A,B)", "explicacao": "explicação simples"}],
  "codigo": {"linguagem": "VBA ou Power Query ou Passo a passo", "conteudo": "código ou passos"},
  "tabela": {"titulo": "título", "colunas": ["col1","col2"], "linhas": [["a","b"]]}
}
Gere no máximo 4 kpis e no máximo 3 gráficos, sempre com dados numéricos reais extraídos da base quando houver. Todos os textos em português do Brasil, simples e diretos.`;

function parseJsonLoose(text) {
  const clean = String(text || "").replace(/```json|```/g, "").trim();
  const start = clean.indexOf("{");
  const end = clean.lastIndexOf("}");
  if (start < 0 || end < 0) throw new Error("Resposta sem JSON");
  return JSON.parse(clean.slice(start, end + 1));
}

async function callDeepSeek(messages, maxTokens = 4000) {
  const r = await fetch("/api/deepseek/chat/completions", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "deepseek-chat",
      max_tokens: maxTokens,
      temperature: 0.3,
      messages,
    }),
  });
  const data = await r.json();
  if (!r.ok) {
    const msg = data?.error?.message || `Erro da API (${r.status})`;
    throw new Error(msg);
  }
  return data?.choices?.[0]?.message?.content || "";
}

function modeCatalog() {
  return MODES.map((m) => `${m.id}. ${m.nome} — ${m.desc}`).join("\n");
}

async function classificarModo(duvida, temDados, temDados2) {
  const text = await callDeepSeek([
    {
      role: "system",
      content: "Você classifica pedidos de Excel. Responda só com JSON válido.",
    },
    {
      role: "user",
      content: `Escolha o melhor modo (1 a 8) para este pedido.

MODOS:
${modeCatalog()}

REGRAS:
- Se pedir fórmula/PROCV/SOMASES → 2
- Se pedir limpar duplicados/formatos → 3
- Se pedir comparar duas bases → 4
- Se pedir explicar a planilha → 5
- Se pedir resumo pra gestor → 6
- Se pedir dashboard/KPIs/gráficos de layout → 7
- Se pedir VBA/Power Query/automatizar rotina → 8
- Caso contrário (insights, padrões, oportunidades) → 1
- Tem dados da planilha 1? ${temDados ? "sim" : "não"}
- Tem planilha 2? ${temDados2 ? "sim" : "não"}
- Se o pedido exigir comparar 2 planilhas mas não houver planilha 2, prefira 1 ou 5.

PEDIDO DO USUÁRIO:
${duvida}

Responda APENAS:
{"mode_id": 1, "motivo": "frase curta em português"}`,
    },
  ], 300);
  const parsed = parseJsonLoose(text);
  const mode = MODES.find((m) => m.id === Number(parsed.mode_id)) || MODES[0];
  return { mode, motivo: parsed.motivo || `Melhor encaixe: ${mode.nome}` };
}

// ---------- prompts ----------
function buildPrompt(mode, csv, csv2, objetivo, duvida) {
  const dataBlock = csv ? `\n\nDADOS DA PLANILHA (CSV):\n${csv}` : "";
  const data2Block = csv2 ? `\n\nSEGUNDA PLANILHA (CSV):\n${csv2}` : "";
  const objBlock = objetivo ? `\n\nOBJETIVO DO USUÁRIO: ${objetivo}` : "";
  const duvidaBlock = duvida ? `\n\nPERGUNTA / DÚVIDA DO USUÁRIO (priorize responder isso com base nos dados): ${duvida}` : "";

  const tasks = {
    1: `Analise esta planilha e me entregue: 5 insights importantes, padrões e anomalias, e oportunidades de ação. Explique em linguagem simples. Gere gráficos que mostrem os principais padrões encontrados e KPIs com os números mais relevantes.`,
    2: `Crie a fórmula do Excel para o objetivo descrito. Me entregue no campo "formulas": a fórmula pronta na versão com ponto e vírgula (formula_pv) e na versão com vírgula (formula_virgula), com explicação simples de cada parte. Se houver dados, adapte a fórmula às colunas reais. Sugira até 3 fórmulas alternativas ou complementares.`,
    3: `Padronize nomes, datas, telefones e categorias desta base. Aponte nos "insights": duplicados encontrados, campos vazios e erros de formato. Em "tabela", devolva a base limpa e padronizada (todas as linhas corrigidas). Em "kpis", mostre: linhas originais, duplicados removidos, campos corrigidos.`,
    4: `Compare estas duas planilhas e mostre: divergências de valor entre registros iguais, registros que existem em uma e não na outra, e diferenças de valor. Em "tabela", monte a auditoria com colunas: ID/Chave, Situação (Divergente/Só na 1ª/Só na 2ª), Valor Planilha 1, Valor Planilha 2, Diferença. Gere um gráfico comparativo e KPIs com totais de divergências.`,
    5: `Explique esta planilha como se eu fosse novo na empresa. Mostre no resumo: o objetivo provável da base. Em "insights": as colunas/áreas importantes e o que cada uma significa. Em "kpis": as métricas críticas com valores atuais. Em "riscos": dados inconsistentes, campos duplicados, possíveis problemas. Gere gráficos das métricas principais.`,
    6: `Transforme esta base em um resumo executivo pronto pro gestor: "resumo" com exatamente 5 linhas, "insights" com exatamente 3 achados, "riscos" com exatamente 2 riscos, "proximos_passos" com exatamente 3 ações. KPIs com os 4 números que o gestor precisa ver. Gráficos que sustentem os achados.`,
    7: `Com base nesta base, projete um dashboard: em "kpis", os 4 KPIs principais com valores calculados dos dados. Em "graficos", os melhores gráficos (calcule os dados reais). Em "insights", o layout sugerido do dashboard (o que vai em cima, no meio, embaixo). Em "proximos_passos", o storytelling executivo: a narrativa que o dashboard deve contar.`,
    8: `Crie uma solução para automatizar a tarefa descrita. Em "codigo", entregue a solução completa (VBA ou Power Query, o que for melhor — ou passo a passo se for mais simples). Em "insights", explique como instalar/usar em linguagem de quem nunca mexeu com isso. Em "proximos_passos", melhorias futuras possíveis.`,
  };

  return `${tasks[mode.id]}${objBlock}${duvidaBlock}${dataBlock}${data2Block}\n\n${JSON_BASE}`;
}

function buildFreePrompt(csv, csv2, duvida) {
  const dataBlock = csv ? `\n\nDADOS DA PLANILHA (CSV):\n${csv}` : "";
  const data2Block = csv2 ? `\n\nSEGUNDA PLANILHA (CSV):\n${csv2}` : "";
  return `Responda à dúvida do usuário de forma direta e útil.
Não se force a seguir um dos 8 modos prontos — escolha o formato de resposta que melhor resolve a pergunta (insights, fórmula, limpeza, comparação, explicação, resumo, dashboard ou automação).
Se faltarem dados para uma parte, diga no resumo e avance com o que for possível.

DÚVIDA DO USUÁRIO:
${duvida}${dataBlock}${data2Block}

${JSON_BASE}`;
}

function validarModo(m, dadosFinais, csv2, objetivo) {
  if (m.objetivo && String(objetivo || "").trim().length <= 3) {
    return "Descreva melhor o que você quer (objetivo/dúvida).";
  }
  if (m.dados && !dadosFinais) return "Envie a planilha (ou cole os dados) para este modo.";
  if (m.dados2 && !csv2) return "Este modo precisa da Planilha 2.";
  return "";
}

// ---------- leitura de arquivo ----------
async function fileToCSV(file) {
  const buf = await file.arrayBuffer();
  if (/\.(xlsx|xls|xlsm)$/i.test(file.name)) {
    const wb = XLSX.read(buf, { type: "array" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_csv(sheet);
  }
  return new TextDecoder("utf-8").decode(buf);
}
function limitCSV(csv, maxRows = 180, maxChars = 24000) {
  if (!csv) return "";
  const lines = csv.trim().split(/\r?\n/).slice(0, maxRows);
  let out = lines.join("\n");
  if (out.length > maxChars) out = out.slice(0, maxChars);
  return out;
}
function previewRows(csv, n = 6) {
  if (!csv) return null;
  const lines = csv.trim().split(/\r?\n/).slice(0, n);
  return lines.map((l) => l.split(",").slice(0, 6));
}

// ---------- componentes de resultado ----------
function KpiCards({ kpis }) {
  if (!kpis?.length) return null;
  return (
    <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
      {kpis.slice(0, 4).map((k, i) => (
        <div key={i} className="rounded-xl p-4" style={{ background: C.surface2, border: `1px solid ${C.line}` }}>
          <div className="text-xs uppercase tracking-widest mb-1" style={{ color: C.muted }}>{k.label}</div>
          <div className="text-2xl font-black" style={{ color: C.greenBright, fontFamily: "'Sora', sans-serif" }}>{k.valor}</div>
          {k.detalhe && <div className="text-xs mt-1" style={{ color: C.muted }}>{k.detalhe}</div>}
        </div>
      ))}
    </div>
  );
}

function Chart({ g }) {
  const data = (g.dados || []).map((d) => ({ nome: String(d.nome), valor: Number(d.valor) || 0 }));
  if (!data.length) return null;
  const common = { fontSize: 11, fill: C.muted };
  return (
    <div className="rounded-xl p-4" style={{ background: C.surface2, border: `1px solid ${C.line}` }}>
      <div className="text-sm font-bold mb-3" style={{ color: C.cream }}>{g.titulo}</div>
      <div style={{ width: "100%", height: 240 }}>
        <ResponsiveContainer>
          {g.tipo === "pie" ? (
            <PieChart>
              <Pie data={data} dataKey="valor" nameKey="nome" outerRadius={85} innerRadius={45} paddingAngle={2}>
                {data.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}
              </Pie>
              <Tooltip contentStyle={{ background: C.surface, border: `1px solid ${C.line}`, borderRadius: 8, color: C.cream }} />
              <Legend wrapperStyle={{ fontSize: 11, color: C.muted }} />
            </PieChart>
          ) : g.tipo === "line" ? (
            <LineChart data={data}>
              <CartesianGrid stroke={C.line} strokeDasharray="3 3" />
              <XAxis dataKey="nome" tick={common} />
              <YAxis tick={common} width={44} />
              <Tooltip contentStyle={{ background: C.surface, border: `1px solid ${C.line}`, borderRadius: 8, color: C.cream }} />
              <Line type="monotone" dataKey="valor" stroke={C.greenBright} strokeWidth={2.5} dot={{ fill: C.orange, r: 3.5 }} />
            </LineChart>
          ) : (
            <BarChart data={data}>
              <CartesianGrid stroke={C.line} strokeDasharray="3 3" vertical={false} />
              <XAxis dataKey="nome" tick={common} />
              <YAxis tick={common} width={44} />
              <Tooltip cursor={{ fill: "#ffffff08" }} contentStyle={{ background: C.surface, border: `1px solid ${C.line}`, borderRadius: 8, color: C.cream }} />
              <Bar dataKey="valor" radius={[6, 6, 0, 0]}>
                {data.map((_, i) => <Cell key={i} fill={CHART_COLORS[i % CHART_COLORS.length]} />)}
              </Bar>
            </BarChart>
          )}
        </ResponsiveContainer>
      </div>
    </div>
  );
}

function ListBlock({ titulo, itens, cor, marcador }) {
  if (!itens?.length) return null;
  return (
    <div className="rounded-xl p-4" style={{ background: C.surface2, border: `1px solid ${C.line}` }}>
      <div className="text-xs uppercase tracking-widest font-bold mb-3" style={{ color: cor }}>{titulo}</div>
      <ul className="space-y-2">
        {itens.map((t, i) => (
          <li key={i} className="flex gap-2 text-sm leading-relaxed" style={{ color: C.cream }}>
            <span style={{ color: cor }}>{marcador}</span><span>{t}</span>
          </li>
        ))}
      </ul>
    </div>
  );
}

function FormulaBlock({ formulas }) {
  if (!formulas?.length) return null;
  return (
    <div className="space-y-3">
      {formulas.map((f, i) => (
        <div key={i} className="rounded-xl p-4" style={{ background: C.surface2, border: `1px solid ${C.orange}44` }}>
          <div className="text-sm font-bold mb-2" style={{ color: C.orange }}>{f.objetivo}</div>
          {f.formula_pv && (
            <CopyLine label="Versão com ;" text={f.formula_pv} />
          )}
          {f.formula_virgula && (
            <CopyLine label="Versão com ," text={f.formula_virgula} />
          )}
          {f.explicacao && <p className="text-xs mt-2 leading-relaxed" style={{ color: C.muted }}>{f.explicacao}</p>}
        </div>
      ))}
    </div>
  );
}

function CopyLine({ label, text }) {
  const [ok, setOk] = useState(false);
  return (
    <div className="mb-2">
      <div className="text-[10px] uppercase tracking-widest mb-1" style={{ color: C.muted }}>{label}</div>
      <div className="flex items-center gap-2">
        <code className="flex-1 text-xs px-3 py-2 rounded-lg overflow-x-auto whitespace-nowrap" style={{ background: C.bg, color: C.greenBright, border: `1px solid ${C.line}` }}>{text}</code>
        <button
          onClick={() => { navigator.clipboard?.writeText(text); setOk(true); setTimeout(() => setOk(false), 1500); }}
          className="text-xs px-3 py-2 rounded-lg font-bold shrink-0"
          style={{ background: ok ? C.green : C.orange, color: "#0C0A07" }}
        >{ok ? "Copiado" : "Copiar"}</button>
      </div>
    </div>
  );
}

function CodeBlock({ codigo }) {
  const [ok, setOk] = useState(false);
  if (!codigo?.conteudo) return null;
  return (
    <div className="rounded-xl overflow-hidden" style={{ border: `1px solid ${C.orange}44` }}>
      <div className="flex items-center justify-between px-4 py-2" style={{ background: C.surface2 }}>
        <span className="text-xs font-bold uppercase tracking-widest" style={{ color: C.orange }}>{codigo.linguagem}</span>
        <button
          onClick={() => { navigator.clipboard?.writeText(codigo.conteudo); setOk(true); setTimeout(() => setOk(false), 1500); }}
          className="text-xs px-3 py-1 rounded-lg font-bold"
          style={{ background: ok ? C.green : C.orange, color: "#0C0A07" }}
        >{ok ? "Copiado" : "Copiar código"}</button>
      </div>
      <pre className="text-xs p-4 overflow-x-auto leading-relaxed" style={{ background: C.bg, color: C.greenBright, maxHeight: 420 }}>{codigo.conteudo}</pre>
    </div>
  );
}

function TableBlock({ tabela }) {
  if (!tabela?.linhas?.length) return null;
  return (
    <div className="rounded-xl overflow-hidden" style={{ border: `1px solid ${C.line}` }}>
      <div className="px-4 py-2 text-sm font-bold" style={{ background: C.surface2, color: C.cream }}>{tabela.titulo}</div>
      <div className="overflow-x-auto" style={{ maxHeight: 380 }}>
        <table className="w-full text-xs">
          <thead>
            <tr style={{ background: C.green + "22" }}>
              {tabela.colunas?.map((c, i) => (
                <th key={i} className="text-left px-3 py-2 font-bold whitespace-nowrap" style={{ color: C.greenBright }}>{c}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {tabela.linhas.map((r, i) => (
              <tr key={i} style={{ borderTop: `1px solid ${C.line}`, background: i % 2 ? C.surface : "transparent" }}>
                {r.map((cell, j) => (
                  <td key={j} className="px-3 py-2 whitespace-nowrap" style={{ color: C.cream }}>{String(cell)}</td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ---------- upload ----------
function FileDrop({ label, csv, onFile, onClear }) {
  const ref = useRef(null);
  const rows = previewRows(csv);
  return (
    <div>
      <div className="text-xs uppercase tracking-widest font-bold mb-2" style={{ color: C.muted }}>{label}</div>
      {!csv ? (
        <button
          onClick={() => ref.current?.click()}
          className="w-full rounded-xl py-6 px-4 text-sm font-semibold transition-transform hover:scale-[1.01]"
          style={{ background: C.surface2, border: `2px dashed ${C.green}66`, color: C.greenBright }}
        >
          ⬆ Enviar arquivo .xlsx ou .csv
          <div className="text-xs font-normal mt-1" style={{ color: C.muted }}>ou cole os dados no campo abaixo</div>
        </button>
      ) : (
        <div className="rounded-xl p-3" style={{ background: C.surface2, border: `1px solid ${C.green}66` }}>
          <div className="flex items-center justify-between mb-2">
            <span className="text-xs font-bold" style={{ color: C.greenBright }}>✓ Dados carregados ({csv.trim().split(/\r?\n/).length} linhas)</span>
            <button onClick={onClear} className="text-xs px-2 py-1 rounded" style={{ color: C.orange }}>Trocar</button>
          </div>
          {rows && (
            <div className="overflow-x-auto rounded-lg" style={{ background: C.bg }}>
              <table className="text-[10px] w-full">
                <tbody>
                  {rows.map((r, i) => (
                    <tr key={i} style={{ borderBottom: `1px solid ${C.line}` }}>
                      {r.map((c, j) => (
                        <td key={j} className="px-2 py-1 whitespace-nowrap" style={{ color: i === 0 ? C.greenBright : C.muted, fontWeight: i === 0 ? 700 : 400 }}>{c}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}
      <input
        ref={ref} type="file" accept=".xlsx,.xls,.xlsm,.csv,.txt" className="hidden"
        onChange={async (e) => {
          const f = e.target.files?.[0];
          if (f) onFile(limitCSV(await fileToCSV(f)));
          e.target.value = "";
        }}
      />
    </div>
  );
}

// ---------- app ----------
export default function AgenteExcel() {
  const [mode, setMode] = useState(MODES[0]);
  const [csv, setCsv] = useState("");
  const [csv2, setCsv2] = useState("");
  const [paste, setPaste] = useState("");
  const [objetivo, setObjetivo] = useState("");
  const [duvida, setDuvida] = useState("");
  const [fluxoDuvida, setFluxoDuvida] = useState("sugerir");
  const [sugestao, setSugestao] = useState(null);
  const [loading, setLoading] = useState(false);
  const [statusMsg, setStatusMsg] = useState("");
  const [erro, setErro] = useState("");
  const [res, setRes] = useState(null);

  const dadosFinais = csv || limitCSV(paste);
  const temDuvida = duvida.trim().length > 3;

  const podeAnalisar = (() => {
    if (loading) return false;
    // Com dúvida: as 3 opções (sugerir / auto / livre)
    if (temDuvida) return true;
    // Sem dúvida: fluxo clássico pelo modo clicado
    const falta = validarModo(mode, dadosFinais, csv2, objetivo);
    return !falta;
  })();

  function mapErro(e) {
    const msg = e?.message || "";
    if (msg.includes("API key") || msg.includes("Unauthorized") || msg.includes("401")) {
      return "Chave DeepSeek inválida ou ausente. Coloque DEEPSEEK_API_KEY no arquivo .env e reinicie o servidor.";
    }
    return msg || "O agente não conseguiu processar. Tente com menos linhas ou verifique se os dados têm cabeçalho na primeira linha.";
  }

  async function rodarModo(modeAlvo, { forcarObjetivo } = {}) {
    const obj = forcarObjetivo || objetivo || duvida.trim();
    const falta = validarModo(modeAlvo, dadosFinais, csv2, obj);
    if (falta) throw new Error(falta);
    const prompt = buildPrompt(
      modeAlvo,
      modeAlvo.objetivo && !modeAlvo.dadosOpc ? "" : dadosFinais,
      csv2,
      obj,
      duvida.trim(),
    );
    const text = await callDeepSeek([
      {
        role: "system",
        content: "Você é um analista de Excel. Responda somente com JSON válido, sem markdown.",
      },
      { role: "user", content: prompt },
    ]);
    return parseJsonLoose(text);
  }

  async function analisar() {
    setLoading(true); setErro(""); setRes(null); setSugestao(null); setStatusMsg("");
    try {
      // Sem dúvida → modo clicado (clássico)
      if (!temDuvida) {
        setStatusMsg(`Rodando ${mode.nome}…`);
        setRes(await rodarModo(mode));
        return;
      }

      // ------- dúvida livre -------
      if (fluxoDuvida === "livre") {
        setStatusMsg("Respondendo sua dúvida em modo livre…");
        const text = await callDeepSeek([
          {
            role: "system",
            content: "Você é um analista de Excel. Responda somente com JSON válido, sem markdown.",
          },
          { role: "user", content: buildFreePrompt(dadosFinais, csv2, duvida.trim()) },
        ]);
        setRes(parseJsonLoose(text));
        return;
      }

      // ------- sugerir ou auto -------
      setStatusMsg("Lendo sua dúvida e sugerindo o melhor modo…");
      const { mode: melhor, motivo } = await classificarModo(duvida.trim(), !!dadosFinais, !!csv2);
      setMode(melhor);
      setSugestao({ mode: melhor, motivo });

      if (fluxoDuvida === "sugerir") {
        setStatusMsg("");
        return; // só sugere; usuário confirma e roda
      }

      setStatusMsg(`Modo escolhido: ${melhor.nome}. Rodando…`);
      setRes(await rodarModo(melhor, { forcarObjetivo: duvida.trim() }));
    } catch (e) {
      setErro(mapErro(e));
    } finally {
      setLoading(false);
      setStatusMsg("");
    }
  }

  async function confirmarSugestao() {
    if (!sugestao?.mode) return;
    setLoading(true); setErro(""); setRes(null); setStatusMsg(`Rodando ${sugestao.mode.nome}…`);
    try {
      setMode(sugestao.mode);
      const parsed = await rodarModo(sugestao.mode, { forcarObjetivo: duvida.trim() });
      setRes(parsed);
    } catch (e) {
      setErro(mapErro(e));
    } finally {
      setLoading(false);
      setStatusMsg("");
    }
  }

  const botaoLabel = (() => {
    if (loading) return statusMsg || "Agente trabalhando…";
    if (!temDuvida) return `Rodar ${mode.nome}`;
    if (fluxoDuvida === "sugerir") return "Sugerir o melhor modo";
    if (fluxoDuvida === "auto") return "Escolher modo e rodar";
    return "Responder dúvida livre";
  })();

  return (
    <div className="min-h-screen" style={{ background: C.bg, fontFamily: "'Inter', system-ui, sans-serif" }}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Sora:wght@700;800&family=Inter:wght@400;600;700&display=swap');
        @keyframes pulse-glow { 0%,100%{opacity:.5} 50%{opacity:1} }
        .glow { animation: pulse-glow 1.4s ease-in-out infinite }`}</style>

      {/* header */}
      <header className="max-w-5xl mx-auto px-4 pt-8 pb-6">
        <div className="flex items-center gap-3">
          <div
            className="w-12 h-12 rounded-xl flex items-center justify-center text-2xl shrink-0"
            style={{ background: `linear-gradient(135deg, ${C.orange}, ${C.orangeDeep})`, boxShadow: `0 0 24px ${C.orange}55` }}
          >✳</div>
          <div>
            <h1 className="text-2xl sm:text-3xl font-extrabold leading-none" style={{ fontFamily: "'Sora', sans-serif", color: C.cream }}>
              AGENTE <span style={{ color: C.orange }}>EXCEL</span>
            </h1>
            <p className="text-xs mt-1" style={{ color: C.muted }}>Joga a planilha. Conta o que precisa. Recebe pronto.</p>
          </div>
        </div>
      </header>

      <main className="max-w-5xl mx-auto px-4 pb-16 space-y-6">
        
        {/* 0 · 7 planilhas prontas e automatizadas */}
        <section className="rounded-2xl p-4 md:p-5 border" style={{ background: C.surface, borderColor: C.line }}>
          <div className="flex flex-col md:flex-row md:items-end md:justify-between gap-2 mb-4">
            <div>
              <div className="text-[11px] tracking-[0.2em] uppercase mb-1" style={{ color: C.muted }}>0 · Serviços automatizados</div>
              <h2 className="text-lg font-semibold" style={{ color: C.cream }}>7 planilhas prontas da Trust Excel</h2>
              <p className="text-sm mt-1" style={{ color: C.muted }}>
                Baixe, preencha e deixe as fórmulas trabalharem. Cada arquivo já vem com painel, validações e automações.
              </p>
            </div>
            <a
              href="https://trustcorp.com.br"
              target="_blank"
              rel="noreferrer"
              className="text-xs underline underline-offset-4"
              style={{ color: C.orange }}
            >
              Ver no hub TrustCorp
            </a>
          </div>
          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
            {PLANILHA_SERVICES.map((s) => (
              <article key={s.id} className="rounded-xl p-3 border flex flex-col gap-2" style={{ background: C.surface2, borderColor: C.line }}>
                <div className="flex items-center justify-between gap-2">
                  <span className="text-[10px] tracking-[0.16em] uppercase" style={{ color: C.orange }}>{s.tag}</span>
                  <span className="text-[10px]" style={{ color: C.muted }}>{s.tabs}</span>
                </div>
                <h3 className="text-sm font-semibold" style={{ color: C.cream }}>{s.nome}</h3>
                <p className="text-xs leading-relaxed" style={{ color: C.muted }}>{s.desc}</p>
                <p className="text-[11px]" style={{ color: C.greenBright }}>Automação: {s.auto}</p>
                <div className="mt-auto pt-1">
                  <a
                    href={s.file}
                    download
                    className="inline-flex items-center justify-center rounded-lg px-3 py-2 text-xs font-semibold"
                    style={{ background: C.orange, color: C.bg }}
                  >
                    Baixar planilha
                  </a>
                </div>
              </article>
            ))}
          </div>
        </section>

{/* modos */}
        <section>
          <div className="text-xs uppercase tracking-widest font-bold mb-3" style={{ color: C.muted }}>1 · Os 8 modos</div>
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
            {MODES.map((m) => {
              const on = m.id === mode.id;
              const sugerido = sugestao?.mode?.id === m.id;
              return (
                <button
                  key={m.id}
                  onClick={() => { setMode(m); setRes(null); setErro(""); setSugestao(null); }}
                  className="text-left rounded-xl p-3 transition-all"
                  style={{
                    background: on || sugerido ? `linear-gradient(135deg, ${C.orange}22, ${C.surface2})` : C.surface,
                    border: `1px solid ${on || sugerido ? C.orange : C.line}`,
                    boxShadow: on || sugerido ? `0 0 16px ${C.orange}33` : "none",
                  }}
                >
                  <div className="text-[10px] font-bold tracking-widest mb-1" style={{ color: on || sugerido ? C.orange : C.muted }}>
                    {m.tag}{sugerido ? " · SUGERIDO" : ""}
                  </div>
                  <div className="text-sm font-bold leading-tight" style={{ color: C.cream }}>{m.icon} {m.nome}</div>
                  <div className="text-[10px] mt-1 leading-snug hidden sm:block" style={{ color: C.muted }}>{m.desc}</div>
                </button>
              );
            })}
          </div>
        </section>

        {/* entrada */}
        <section className="space-y-4">
          <div className="text-xs uppercase tracking-widest font-bold" style={{ color: C.muted }}>2 · Entregue os dados</div>

          {mode.objetivo && !temDuvida && (
            <div>
              <div className="text-xs uppercase tracking-widest font-bold mb-2" style={{ color: C.muted }}>
                {mode.id === 2 ? "O que a fórmula precisa fazer?" : "Qual tarefa você quer automatizar?"}
              </div>
              <textarea
                value={objetivo}
                onChange={(e) => setObjetivo(e.target.value)}
                placeholder={mode.placeholder}
                rows={3}
                className="w-full rounded-xl p-3 text-sm outline-none"
                style={{ background: C.surface2, border: `1px solid ${C.line}`, color: C.cream, resize: "vertical" }}
              />
            </div>
          )}

          <FileDrop
            label={mode.dados2 ? "Planilha 1" : "Sua planilha"}
            csv={csv}
            onFile={setCsv}
            onClear={() => setCsv("")}
          />
          {!csv && (
            <textarea
              value={paste}
              onChange={(e) => setPaste(e.target.value)}
              placeholder={"Ou cole aqui direto do Excel (Ctrl+C na planilha, Ctrl+V aqui)\nExemplo:\nVendas,Custo,Região,Mês\n12540,7230,Sudeste,Jan\n9870,6120,Sul,Jan"}
              rows={5}
              className="w-full rounded-xl p-3 text-xs outline-none font-mono"
              style={{ background: C.surface2, border: `1px solid ${C.line}`, color: C.greenBright, resize: "vertical" }}
            />
          )}

          <FileDrop label="Planilha 2 (se for comparar)" csv={csv2} onFile={setCsv2} onClear={() => setCsv2("")} />
        </section>

        {/* dúvidas */}
        <section className="space-y-3">
          <div className="text-xs uppercase tracking-widest font-bold" style={{ color: C.muted }}>
            3 · Dúvidas
          </div>
          <div
            className="rounded-xl p-4 space-y-4"
            style={{ background: C.surface, border: `1px solid ${C.line}` }}
          >
            <div>
              <label className="block text-sm font-bold mb-2" style={{ color: C.cream }}>
                O que você quer fazer hoje com seus números ou dados?
              </label>
              <textarea
                value={duvida}
                onChange={(e) => { setDuvida(e.target.value); setSugestao(null); }}
                placeholder={"Ex.: Quais produtos caíram de margem este mês? Quero priorizar reposição das 10 filiais com mais ruptura."}
                rows={4}
                className="w-full rounded-xl p-3 text-sm outline-none"
                style={{ background: C.surface2, border: `1px solid ${C.line}`, color: C.cream, resize: "vertical" }}
              />
            </div>

            <div>
              <div className="text-xs uppercase tracking-widest font-bold mb-2" style={{ color: C.muted }}>
                Como o agente deve usar essa dúvida?
              </div>
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
                {DUVIDA_FLOWS.map((f) => {
                  const on = fluxoDuvida === f.id;
                  return (
                    <button
                      key={f.id}
                      type="button"
                      onClick={() => { setFluxoDuvida(f.id); setSugestao(null); setErro(""); }}
                      className="text-left rounded-xl p-3 transition-all"
                      style={{
                        background: on ? `linear-gradient(135deg, ${C.green}22, ${C.surface2})` : C.surface2,
                        border: `1px solid ${on ? C.greenBright : C.line}`,
                      }}
                    >
                      <div className="text-sm font-bold" style={{ color: on ? C.greenBright : C.cream }}>{f.nome}</div>
                      <div className="text-[10px] mt-1 leading-snug" style={{ color: C.muted }}>{f.desc}</div>
                    </button>
                  );
                })}
              </div>
            </div>

            {sugestao && (
              <div
                className="rounded-xl p-4"
                style={{ background: `${C.orange}18`, border: `1px solid ${C.orange}66` }}
              >
                <div className="text-xs uppercase tracking-widest font-bold mb-1" style={{ color: C.orange }}>Sugestão</div>
                <div className="text-sm font-bold mb-1" style={{ color: C.cream }}>
                  {sugestao.mode.icon} {sugestao.mode.nome} ({sugestao.mode.tag})
                </div>
                <p className="text-xs mb-3" style={{ color: C.muted }}>{sugestao.motivo}</p>
                {fluxoDuvida === "sugerir" && (
                  <button
                    type="button"
                    onClick={confirmarSugestao}
                    disabled={loading}
                    className="w-full py-3 rounded-xl font-extrabold text-sm"
                    style={{
                      fontFamily: "'Sora', sans-serif",
                      background: `linear-gradient(135deg, ${C.orange}, ${C.orangeDeep})`,
                      color: "#0C0A07",
                      cursor: loading ? "not-allowed" : "pointer",
                    }}
                  >
                    {loading ? <span className="glow">✳ Rodando…</span> : `Confirmar e rodar ${sugestao.mode.nome}`}
                  </button>
                )}
              </div>
            )}
          </div>
        </section>

        <section className="space-y-4">
          <button
            onClick={analisar}
            disabled={!podeAnalisar}
            className="w-full py-4 rounded-xl font-extrabold text-base tracking-wide transition-all"
            style={{
              fontFamily: "'Sora', sans-serif",
              background: podeAnalisar ? `linear-gradient(135deg, ${C.orange}, ${C.orangeDeep})` : C.surface2,
              color: podeAnalisar ? "#0C0A07" : C.muted,
              boxShadow: podeAnalisar ? `0 0 28px ${C.orange}44` : "none",
              cursor: podeAnalisar ? "pointer" : "not-allowed",
            }}
          >
            {loading ? <span className="glow">✳ {botaoLabel}</span> : `⚡ ${botaoLabel}`}
          </button>

          {erro && (
            <div className="rounded-xl p-4 text-sm" style={{ background: "#3a120a", border: "1px solid #7c2d12", color: "#fca5a5" }}>
              {erro}
            </div>
          )}
        </section>

        {/* resultado */}
        {res && (
          <section className="space-y-4">
            <div className="flex items-center gap-2">
              <div className="h-px flex-1" style={{ background: C.line }} />
              <div className="text-xs uppercase tracking-widest font-bold" style={{ color: C.green }}>Resultado pronto</div>
              <div className="h-px flex-1" style={{ background: C.line }} />
            </div>

            <div className="rounded-xl p-5" style={{ background: `linear-gradient(135deg, ${C.green}18, ${C.surface})`, border: `1px solid ${C.green}55` }}>
              <h2 className="text-xl font-extrabold mb-2" style={{ fontFamily: "'Sora', sans-serif", color: C.greenBright }}>{res.titulo}</h2>
              <p className="text-sm leading-relaxed" style={{ color: C.cream }}>{res.resumo}</p>
            </div>

            <KpiCards kpis={res.kpis} />

            {res.graficos?.length > 0 && (
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-3">
                {res.graficos.slice(0, 3).map((g, i) => <Chart key={i} g={g} />)}
              </div>
            )}

            <FormulaBlock formulas={res.formulas} />
            <CodeBlock codigo={res.codigo} />
            <TableBlock tabela={res.tabela} />

            <div className="grid grid-cols-1 lg:grid-cols-3 gap-3">
              <ListBlock titulo="Achados" itens={res.insights} cor={C.greenBright} marcador="●" />
              <ListBlock titulo="Riscos" itens={res.riscos} cor={C.orange} marcador="⚠" />
              <ListBlock titulo="Próximos passos" itens={res.proximos_passos} cor={C.cream} marcador="→" />
            </div>
          </section>
        )}

        {!res && !loading && (
          <div className="text-center text-xs pt-4" style={{ color: C.muted }}>
            Sai da análise e vai pra decisão. · Os dados enviados são limitados às primeiras ~180 linhas.
          </div>
        )}
      </main>
    </div>
  );
}
