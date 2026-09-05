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
 * Build a repair-advisor LlmText from the environment. Two provider shapes:
 * 1. DEEPSEEK_API_KEY (+ optional DEEPSEEK_BASE_URL / DEEPSEEK_MODEL)
 * 2. BAI_API against https://api.b.ai/v1 (+ optional BAI_MODEL,
 *    e.g. glm-5.3-flash or qwen3.8-flash); BAI_BASE_URL overrides the host.
 * LLM_PROVIDER=bai|deepseek forces a provider; otherwise DEEPSEEK wins when
 * both keys are set.
 */
export function deepseekLlmTextFromEnv(model = process.env.DEEPSEEK_MODEL ?? 'deepseek-chat'): LlmText {
  const deepseekKey = process.env.DEEPSEEK_API_KEY
  const baiKey = process.env.BAI_API
  const provider = process.env.LLM_PROVIDER === 'bai' || process.env.LLM_PROVIDER === 'deepseek'
    ? process.env.LLM_PROVIDER
    : (deepseekKey ? 'deepseek' : 'bai')
  if ((provider === 'deepseek' && !deepseekKey) || (provider === 'bai' && !baiKey)) {
    throw new Error(`LLM provider "${provider}" selected but its API key (DEEPSEEK_API_KEY / BAI_API) is not set`)
  }
  const apiKey = provider === 'deepseek' ? deepseekKey! : baiKey!
  const baseUrl = provider === 'deepseek'
    ? process.env.DEEPSEEK_BASE_URL
    : (process.env.BAI_BASE_URL ?? 'https://api.b.ai/v1')
  const resolvedModel = provider === 'deepseek' ? model : (process.env.BAI_MODEL ?? model)
  // Reasoning models (glm/qwen flash) spend completion tokens on hidden
  // reasoning; allow a larger budget when calling them.
  const maxTokens = provider === 'deepseek' ? 2000 : Number(process.env.BAI_MAX_TOKENS ?? 4000)
  return async (prompt, signal) => deepseekChatCompletion({
    apiKey,
    baseUrl,
    model: resolvedModel,
    messages: [{ role: 'user', content: prompt }],
    temperature: 0,
    maxTokens,
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
