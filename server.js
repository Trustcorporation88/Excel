import express from "express";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;

const OPENAI_API_KEY = (process.env.OPENAI_API_KEY || "").trim();
const DEEPSEEK_API_KEY = (process.env.DEEPSEEK_API_KEY || "").trim();
const OPENAI_MODEL = process.env.OPENAI_MODEL || "gpt-4o-mini";
const DEEPSEEK_MODEL = process.env.DEEPSEEK_MODEL || "deepseek-chat";

app.use(express.json({ limit: "2mb" }));

async function callProvider({ name, url, apiKey, model, messages, max_tokens, temperature }) {
  if (!apiKey) {
    const err = new Error(`${name}_missing_key`);
    err.status = 401;
    err.provider = name;
    throw err;
  }

  const r = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model,
      messages,
      max_tokens,
      temperature,
    }),
  });

  let data = {};
  try {
    data = await r.json();
  } catch {
    data = {};
  }

  if (!r.ok) {
    const msg =
      data?.error?.message ||
      data?.message ||
      data?.error ||
      `${name} HTTP ${r.status}`;
    const err = new Error(typeof msg === "string" ? msg : JSON.stringify(msg));
    err.status = r.status;
    err.provider = name;
    err.payload = data;
    throw err;
  }

  const content = data?.choices?.[0]?.message?.content || "";
  return { content, provider: name, model, raw: data };
}

app.get("/api/health", (_req, res) => {
  res.json({
    ok: true,
    openai: Boolean(OPENAI_API_KEY),
    deepseek: Boolean(DEEPSEEK_API_KEY),
    primary: OPENAI_API_KEY ? "openai" : DEEPSEEK_API_KEY ? "deepseek" : null,
  });
});

// Shared core logic: messages -> OpenAI -> DeepSeek fallback.
// Returns { status, body } instead of writing to a response, so it can be
// reused by both /api/ai/chat and /api/deepseek/chat/completions without
// making an internal HTTP call back into this same server.
async function runAiChat({ messages, max_tokens, temperature }) {
  if (!Array.isArray(messages) || messages.length === 0) {
    return { status: 400, body: { error: "messages[] é obrigatório" } };
  }

  if (!OPENAI_API_KEY && !DEEPSEEK_API_KEY) {
    return {
      status: 401,
      body: {
        error:
          "Nenhuma chave configurada. Defina OPENAI_API_KEY e/ou DEEPSEEK_API_KEY no Railway.",
      },
    };
  }

  const attempts = [];
  // 1) OpenAI first
  if (OPENAI_API_KEY) {
    try {
      const out = await callProvider({
        name: "openai",
        url: "https://api.openai.com/v1/chat/completions",
        apiKey: OPENAI_API_KEY,
        model: OPENAI_MODEL,
        messages,
        max_tokens,
        temperature,
      });
      return { status: 200, body: out };
    } catch (e) {
      attempts.push({ provider: "openai", error: e.message, status: e.status || 500 });
    }
  }

  // 2) DeepSeek fallback
  if (DEEPSEEK_API_KEY) {
    try {
      const out = await callProvider({
        name: "deepseek",
        url: "https://api.deepseek.com/chat/completions",
        apiKey: DEEPSEEK_API_KEY,
        model: DEEPSEEK_MODEL,
        messages,
        max_tokens,
        temperature,
      });
      return {
        status: 200,
        body: { ...out, fallback_from: attempts[0]?.provider || null, attempts },
      };
    } catch (e) {
      attempts.push({ provider: "deepseek", error: e.message, status: e.status || 500 });
    }
  }

  return {
    status: 502,
    body: { error: "Falha em todos os provedores de IA", attempts },
  };
}

app.post("/api/ai/chat", async (req, res) => {
  try {
    const messages = req.body?.messages;
    const max_tokens = Number(req.body?.max_tokens || 4000);
    const temperature = Number(req.body?.temperature ?? 0.3);

    const result = await runAiChat({ messages, max_tokens, temperature });
    return res.status(result.status).json(result.body);
  } catch (e) {
    return res.status(500).json({ error: e.message || "Erro interno" });
  }
});

// Compat: old DeepSeek-only path still works, but also tries OpenAI first.
// Calls the shared logic directly instead of making an internal HTTP fetch
// back into this same server, which could deadlock while it is starting up.
app.post("/api/deepseek/chat/completions", async (req, res) => {
  try {
    const messages = req.body?.messages || [];
    const max_tokens = Number(req.body?.max_tokens || 4000);
    const temperature = Number(req.body?.temperature ?? 0.3);

    const result = await runAiChat({ messages, max_tokens, temperature });
    if (result.status !== 200) {
      return res.status(result.status).json(result.body);
    }

    // Translate to OpenAI-compatible response for old frontend
    const data = result.body;
    return res.json({
      choices: [{ message: { role: "assistant", content: data.content || "" } }],
      provider: data.provider,
      model: data.model,
    });
  } catch (e) {
    return res.status(500).json({ error: { message: e.message || "proxy error" } });
  }
});

// Static frontend
const distDir = path.join(__dirname, "dist");
app.use(express.static(distDir));
app.use((req, res) => {
  // don't swallow API 404s
  if (req.path.startsWith("/api/")) return res.status(404).json({ error: "not found" });
  res.sendFile(path.join(distDir, "index.html"));
});

app.listen(PORT, () => {
  console.log(`Trust Excel listening on :${PORT}`);
  console.log(
    `AI providers => openai:${OPENAI_API_KEY ? "yes" : "no"} deepseek:${DEEPSEEK_API_KEY ? "yes" : "no"}`
  );
});
