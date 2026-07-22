import { defineConfig, loadEnv } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'

export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), '')
  const openaiKey = env.OPENAI_API_KEY || ''
  const deepseekKey = env.DEEPSEEK_API_KEY || ''
  const openaiModel = env.OPENAI_MODEL || 'gpt-4o-mini'
  const deepseekModel = env.DEEPSEEK_MODEL || 'deepseek-chat'

  async function callProvider({ name, url, apiKey, model, body }) {
    if (!apiKey) {
      const err = new Error(`${name}_missing_key`)
      err.status = 401
      throw err
    }
    const r = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        model,
        messages: body.messages,
        max_tokens: body.max_tokens ?? 4000,
        temperature: body.temperature ?? 0.3,
      }),
    })
    const data = await r.json().catch(() => ({}))
    if (!r.ok) {
      const msg = data?.error?.message || data?.message || data?.error || `${name} HTTP ${r.status}`
      const err = new Error(typeof msg === 'string' ? msg : JSON.stringify(msg))
      err.status = r.status
      throw err
    }
    return {
      content: data?.choices?.[0]?.message?.content || '',
      provider: name,
      model,
    }
  }

  function aiMiddleware() {
    return {
      name: 'trust-excel-ai-middleware',
      configureServer(server) {
        server.middlewares.use(async (req, res, next) => {
          if (req.method === 'GET' && req.url?.startsWith('/api/health')) {
            res.setHeader('Content-Type', 'application/json')
            res.end(JSON.stringify({
              ok: true,
              openai: Boolean(openaiKey),
              deepseek: Boolean(deepseekKey),
              primary: openaiKey ? 'openai' : deepseekKey ? 'deepseek' : null,
            }))
            return
          }

          if (req.method === 'POST' && req.url?.startsWith('/api/ai/chat')) {
            try {
              const chunks = []
              for await (const c of req) chunks.push(c)
              const body = JSON.parse(Buffer.concat(chunks).toString('utf8') || '{}')
              if (!Array.isArray(body.messages) || !body.messages.length) {
                res.statusCode = 400
                res.setHeader('Content-Type', 'application/json')
                res.end(JSON.stringify({ error: 'messages[] é obrigatório' }))
                return
              }
              if (!openaiKey && !deepseekKey) {
                res.statusCode = 401
                res.setHeader('Content-Type', 'application/json')
                res.end(JSON.stringify({
                  error: 'Nenhuma chave configurada. Defina OPENAI_API_KEY e/ou DEEPSEEK_API_KEY.',
                }))
                return
              }

              const attempts = []
              if (openaiKey) {
                try {
                  const out = await callProvider({
                    name: 'openai',
                    url: 'https://api.openai.com/v1/chat/completions',
                    apiKey: openaiKey,
                    model: openaiModel,
                    body,
                  })
                  res.setHeader('Content-Type', 'application/json')
                  res.end(JSON.stringify(out))
                  return
                } catch (e) {
                  attempts.push({ provider: 'openai', error: e.message, status: e.status || 500 })
                }
              }
              if (deepseekKey) {
                try {
                  const out = await callProvider({
                    name: 'deepseek',
                    url: 'https://api.deepseek.com/chat/completions',
                    apiKey: deepseekKey,
                    model: deepseekModel,
                    body,
                  })
                  res.setHeader('Content-Type', 'application/json')
                  res.end(JSON.stringify({ ...out, fallback_from: attempts[0]?.provider || null, attempts }))
                  return
                } catch (e) {
                  attempts.push({ provider: 'deepseek', error: e.message, status: e.status || 500 })
                }
              }
              res.statusCode = 502
              res.setHeader('Content-Type', 'application/json')
              res.end(JSON.stringify({ error: 'Falha em todos os provedores de IA', attempts }))
            } catch (e) {
              res.statusCode = 500
              res.setHeader('Content-Type', 'application/json')
              res.end(JSON.stringify({ error: e.message || 'Erro interno' }))
            }
            return
          }

          next()
        })
      },
    }
  }

  return {
    plugins: [react(), tailwindcss(), aiMiddleware()],
  }
})
