import { test } from 'node:test'
import assert from 'node:assert/strict'
import { deepseekChatCompletion } from '../src/deepseek.ts'

test('deepseekChatCompletion posts chat completions and parses content', async () => {
  const originalFetch = globalThis.fetch
  let captured: { url: string; init?: RequestInit } | undefined
  globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
    captured = { url: String(input), init }
    return new Response(JSON.stringify({
      choices: [{ message: { content: '{"repairs":[]}' } }],
    }), { status: 200, headers: { 'Content-Type': 'application/json' } })
  }) as typeof fetch
  try {
    const content = await deepseekChatCompletion({
      apiKey: 'sk-test',
      model: 'deepseek-chat',
      messages: [{ role: 'user', content: 'hi' }],
    })
    assert.equal(content, '{"repairs":[]}')
    assert.equal(captured!.url, 'https://api.deepseek.com/chat/completions')
    assert.equal(captured!.init!.headers!['Authorization'], 'Bearer sk-test')
    const body = JSON.parse(String(captured!.init!.body)) as { model: string; stream: boolean }
    assert.equal(body.model, 'deepseek-chat')
    assert.equal(body.stream, false)
  } finally {
    globalThis.fetch = originalFetch
  }
})

test('deepseekChatCompletion surfaces API errors', async () => {
  const originalFetch = globalThis.fetch
  globalThis.fetch = (async () => new Response('bad key', { status: 401 })) as typeof fetch
  try {
    await assert.rejects(
      () => deepseekChatCompletion({ apiKey: 'sk-bad', model: 'deepseek-chat', messages: [] }),
      /DeepSeek API 401: bad key/,
    )
  } finally {
    globalThis.fetch = originalFetch
  }
})

test('deepseekLlmTextFromEnv falls back to BAI_API with api.b.ai base url', async () => {
  const originalFetch = globalThis.fetch
  const original = { deepseek: process.env.DEEPSEEK_API_KEY, bai: process.env.BAI_API, baiModel: process.env.BAI_MODEL, provider: process.env.LLM_PROVIDER }
  process.env.DEEPSEEK_API_KEY = ''
  delete process.env.DEEPSEEK_API_KEY
  process.env.BAI_API = 'bai-key'
  process.env.BAI_MODEL = 'glm-5.3-flash'
  delete process.env.LLM_PROVIDER
  let captured: { url: string; init?: RequestInit } | undefined
  globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
    captured = { url: String(input), init }
    return new Response(JSON.stringify({
      choices: [{ message: { content: 'ok', reasoning_content: 'thinking...' } }],
    }), { status: 200, headers: { 'Content-Type': 'application/json' } })
  }) as typeof fetch
  try {
    const { deepseekLlmTextFromEnv } = await import('../src/deepseek.ts')
    const llm = deepseekLlmTextFromEnv()
    const content = await llm('plan this')
    assert.equal(content, 'ok')
    assert.equal(captured!.url, 'https://api.b.ai/v1/chat/completions')
    assert.equal(captured!.init!.headers!['Authorization'], 'Bearer bai-key')
    const body = JSON.parse(String(captured!.init!.body)) as { model: string; max_tokens: number }
    assert.equal(body.model, 'glm-5.3-flash')
    assert.equal(body.max_tokens, 4000)
  } finally {
    globalThis.fetch = originalFetch
    if (original.deepseek === undefined) delete process.env.DEEPSEEK_API_KEY
    else process.env.DEEPSEEK_API_KEY = original.deepseek
    if (original.bai === undefined) delete process.env.BAI_API
    else process.env.BAI_API = original.bai
    if (original.baiModel === undefined) delete process.env.BAI_MODEL
    else process.env.BAI_MODEL = original.baiModel
    if (original.provider === undefined) delete process.env.LLM_PROVIDER
    else process.env.LLM_PROVIDER = original.provider
  }
})

test('deepseekLlmTextFromEnv honors LLM_PROVIDER=bai even when deepseek key exists', async () => {
  const originalFetch = globalThis.fetch
  const original = { deepseek: process.env.DEEPSEEK_API_KEY, bai: process.env.BAI_API, provider: process.env.LLM_PROVIDER }
  process.env.DEEPSEEK_API_KEY = 'sk-deepseek'
  process.env.BAI_API = 'bai-key'
  process.env.LLM_PROVIDER = 'bai'
  let captured: { url: string } | undefined
  globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
    captured = { url: String(input) }
    return new Response(JSON.stringify({ choices: [{ message: { content: 'ok' } }] }), { status: 200 })
  }) as typeof fetch
  try {
    const { deepseekLlmTextFromEnv } = await import('../src/deepseek.ts')
    await deepseekLlmTextFromEnv()('x')
    assert.equal(captured!.url, 'https://api.b.ai/v1/chat/completions')
  } finally {
    globalThis.fetch = originalFetch
    if (original.deepseek === undefined) delete process.env.DEEPSEEK_API_KEY
    else process.env.DEEPSEEK_API_KEY = original.deepseek
    if (original.bai === undefined) delete process.env.BAI_API
    else process.env.BAI_API = original.bai
    if (original.provider === undefined) delete process.env.LLM_PROVIDER
    else process.env.LLM_PROVIDER = original.provider
  }
})

test('deepseekLlmTextFromEnv requires one of the keys', async () => {
  const original = { deepseek: process.env.DEEPSEEK_API_KEY, bai: process.env.BAI_API, provider: process.env.LLM_PROVIDER }
  delete process.env.DEEPSEEK_API_KEY
  delete process.env.BAI_API
  delete process.env.LLM_PROVIDER
  try {
    const { deepseekLlmTextFromEnv } = await import('../src/deepseek.ts')
    assert.throws(() => deepseekLlmTextFromEnv(), /API key/)
  } finally {
    if (original.deepseek === undefined) delete process.env.DEEPSEEK_API_KEY
    else process.env.DEEPSEEK_API_KEY = original.deepseek
    if (original.bai === undefined) delete process.env.BAI_API
    else process.env.BAI_API = original.bai
    if (original.provider === undefined) delete process.env.LLM_PROVIDER
    else process.env.LLM_PROVIDER = original.provider
  }
})
