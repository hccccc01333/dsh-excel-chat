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
