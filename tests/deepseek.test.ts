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
      /LLM API .*returned 401: bad key/,
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
    // 8192, not 4000: a reasoning model bills hidden reasoning to the same budget as the
    // answer, and 2000 was spent entirely on thinking for a real planning prompt.
    assert.equal(body.max_tokens, 8192)
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

test('an empty completion says the budget ran out, not that the gateway returned nothing', async () => {
  // A reasoning model bills its hidden reasoning to the same output budget as the answer.
  // Against Step 5 Preview a ~9,200-character planning prompt spent the whole 2,000-token
  // budget thinking and returned an empty completion — 55 of 80 benchmark failures were
  // this. The old message ("returned no content") pointed at the gateway, and the gateway
  // was fine; the budget was not. A diagnosis that names the wrong component costs more
  // than no diagnosis.
  const originalFetch = globalThis.fetch
  globalThis.fetch = (async () =>
    new Response(
      JSON.stringify({
        choices: [{ message: { content: '' }, finish_reason: 'length' }],
        usage: { completion_tokens: 2000, completion_tokens_details: { reasoning_tokens: 2000 } },
      }),
      { status: 200, headers: { 'Content-Type': 'application/json' } },
    )) as typeof fetch
  try {
    await assert.rejects(
      () => deepseekChatCompletion({ apiKey: 'sk-x', model: 'step-5-preview-free', messages: [], maxTokens: 2000 }),
      (error: Error) => {
        assert.match(error.message, /ran out of output tokens/)
        assert.match(error.message, /reasoning=2000/)
        assert.match(error.message, /LLM_MAX_TOKENS/)
        return true
      },
    )
  } finally {
    globalThis.fetch = originalFetch
  }
})

test('an empty completion with a normal finish reason still reports what it saw', async () => {
  const originalFetch = globalThis.fetch
  globalThis.fetch = (async () =>
    new Response(
      JSON.stringify({ choices: [{ message: { content: null }, finish_reason: 'stop' }] }),
      { status: 200, headers: { 'Content-Type': 'application/json' } },
    )) as typeof fetch
  try {
    await assert.rejects(
      () => deepseekChatCompletion({ apiKey: 'sk-x', model: 'm', messages: [] }),
      /returned no content \(finish_reason=stop\)/,
    )
  } finally {
    globalThis.fetch = originalFetch
  }
})
