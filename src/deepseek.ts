import type { LlmText } from './advisor.ts'

export interface DeepSeekToolCall {
  id: string
  type: 'function'
  function: { name: string; arguments: string }
}

export interface DeepSeekMessage {
  role: 'system' | 'user' | 'assistant' | 'tool'
  content: string | null
  tool_call_id?: string
  tool_calls?: DeepSeekToolCall[]
}

export interface DeepSeekChatOptions {
  apiKey: string
  baseUrl?: string
  model: string
  messages: DeepSeekMessage[]
  temperature?: number
  maxTokens?: number
  signal?: AbortSignal
}

export async function deepseekChatCompletion(options: DeepSeekChatOptions): Promise<string> {
  const baseUrl = (options.baseUrl ?? 'https://api.deepseek.com').replace(/\/+$/, '')
  const response = await fetch(`${baseUrl}/chat/completions`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${options.apiKey}`,
    },
    body: JSON.stringify({
      model: options.model,
      messages: options.messages,
      temperature: options.temperature ?? 0,
      max_tokens: options.maxTokens ?? 2000,
      stream: false,
    }),
    signal: options.signal,
  })
  if (!response.ok) {
    throw new Error(`DeepSeek API ${response.status}: ${(await response.text()).slice(0, 500)}`)
  }
  const data = (await response.json()) as { choices?: Array<{ message?: { content?: string } }> }
  const content = data.choices?.[0]?.message?.content
  if (typeof content !== 'string') {
    throw new Error('DeepSeek API returned no content')
  }
  return content
}

/**
 * Build a repair-advisor LlmText from the environment: DEEPSEEK_API_KEY is
 * required; DEEPSEEK_BASE_URL and DEEPSEEK_MODEL are optional overrides.
 */
export function deepseekLlmTextFromEnv(model = process.env.DEEPSEEK_MODEL ?? 'deepseek-chat'): LlmText {
  const apiKey = process.env.DEEPSEEK_API_KEY
  if (!apiKey) throw new Error('DEEPSEEK_API_KEY is not set')
  const baseUrl = process.env.DEEPSEEK_BASE_URL
  return async (prompt, signal) => deepseekChatCompletion({
    apiKey,
    baseUrl,
    model,
    messages: [{ role: 'user', content: prompt }],
    temperature: 0,
    maxTokens: 2000,
    signal,
  })
}

export interface DeepSeekChatReply {
  content: string | null
  toolCalls: DeepSeekToolCall[]
}

/** Chat completion with native function-calling tools (DeepSeek tool_calls). */
export async function deepseekChatWithTools(
  options: Omit<DeepSeekChatOptions, 'messages'> & { messages: DeepSeekMessage[]; tools: unknown[] },
): Promise<DeepSeekChatReply> {
  const baseUrl = (options.baseUrl ?? 'https://api.deepseek.com').replace(/\/+$/, '')
  const response = await fetch(`${baseUrl}/chat/completions`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${options.apiKey}`,
    },
    body: JSON.stringify({
      model: options.model,
      messages: options.messages,
      temperature: options.temperature ?? 0,
      max_tokens: options.maxTokens ?? 4000,
      tools: options.tools,
      stream: false,
    }),
    signal: options.signal,
  })
  if (!response.ok) {
    throw new Error(`DeepSeek API ${response.status}: ${(await response.text()).slice(0, 500)}`)
  }
  const data = (await response.json()) as {
    choices?: Array<{ message?: { content?: string | null; tool_calls?: DeepSeekToolCall[] } }>
  }
  const message = data.choices?.[0]?.message
  if (!message) {
    throw new Error('DeepSeek API returned no message')
  }
  return { content: message.content ?? null, toolCalls: message.tool_calls ?? [] }
}
