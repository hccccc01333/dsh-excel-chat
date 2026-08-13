import { test } from 'node:test'
import assert from 'node:assert/strict'
import { Context } from '@deepseek-ai/cordis'
import {
  AttachmentId,
  AttachmentStore,
  type ImageAttachmentRef,
  type SaveImageAttachment,
  type StoredImageAttachment,
} from '@deepseek-ai/dsh-attachment'
import { LlmAdapter, LlmRuntime, type GenerateOptions, type StreamChunk } from '@deepseek-ai/dsh-llm'
import SystemPrompt from '@deepseek-ai/dsh-system-prompt'
import ToolRuntime from '@deepseek-ai/dsh-tools'
import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createVisionCritic } from '../src/chart-visual.ts'
import { visionTextFromContext } from '../src/vision.ts'

class FakeAttachmentStore extends AttachmentStore {
  readonly imageLimits = {
    maxImageBytes: 10_000_000,
    maxImagesPerMessage: 1,
    maxMessageImageBytes: 10_000_000,
    maxImagePixels: 50_000_000,
    mediaTypes: ['image/png'],
  } as const

  async validateImage(): Promise<void> {}

  async saveImage(input: SaveImageAttachment): Promise<ImageAttachmentRef> {
    return {
      attachmentId: AttachmentId('fake-attachment-1'),
      mediaType: input.mediaType,
      bytes: input.data.length,
      width: 1,
      height: 1,
      name: input.name,
    }
  }

  async readImage(ref: ImageAttachmentRef): Promise<StoredImageAttachment> {
    return { ref, data: new Uint8Array() }
  }
}

class RecordingLlmAdapter extends LlmAdapter {
  lastMessage: { content: Array<{ type: string; attachment?: unknown }> } | undefined

  stream(options: GenerateOptions): AsyncIterable<StreamChunk> {
    this.lastMessage = options.messages[0] as { content: Array<{ type: string; attachment?: unknown }> }
    const reply = JSON.stringify({ issues: [{ kind: 'legend', severity: 'warning', description: 'overlap' }] })
    return (async function* () {
      yield { type: 'block-start', index: 0, blockType: 'text' }
      yield { type: 'text-delta', index: 0, text: reply }
      yield { type: 'block-end', index: 0, block: { type: 'text', text: reply } }
      yield { type: 'finish', reason: { kind: 'stop' } }
    })()
  }
}

test('visionTextFromContext uploads the image and streams a vision reply', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'vera-vision-'))
  const imagePath = join(dir, 'chart-1.png')
  await writeFile(imagePath, Buffer.from('fake-png-bytes'))

  const ctx = new Context()
  await ctx.plugin(SystemPrompt)
  await ctx.plugin(ToolRuntime)
  await ctx.plugin(LlmRuntime)
  new FakeAttachmentStore(ctx)
  const adapter = new RecordingLlmAdapter()
  ctx.llm.registerAdapter(['fake-vision'], adapter)

  const text = await visionTextFromContext(ctx, 'fake-vision', 'vision-model')(imagePath, 'check this chart')
  const reply = JSON.parse(text) as { issues: Array<{ kind: string }> }
  assert.equal(reply.issues[0]!.kind, 'legend')
  const imageBlock = adapter.lastMessage!.content.find((block) => block.type === 'image')
  assert.ok(imageBlock)
  assert.equal((imageBlock!.attachment as { mediaType: string }).mediaType, 'image/png')
  await ctx.fiber.dispose()
})

test('vision wiring feeds createVisionCritic end to end', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'vera-vision-'))
  const imagePath = join(dir, 'chart-2.png')
  await writeFile(imagePath, Buffer.from('fake-png-bytes'))

  const ctx = new Context()
  await ctx.plugin(SystemPrompt)
  await ctx.plugin(ToolRuntime)
  await ctx.plugin(LlmRuntime)
  new FakeAttachmentStore(ctx)
  ctx.llm.registerAdapter(['fake-vision'], new RecordingLlmAdapter())

  const report = await createVisionCritic(visionTextFromContext(ctx, 'fake-vision', 'vision-model'))(imagePath)
  assert.equal(report.issues.length, 1)
  assert.equal(report.issues[0]!.severity, 'warning')
  await ctx.fiber.dispose()
})

test('vision wiring rejects missing services', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'vera-vision-'))
  const imagePath = join(dir, 'chart-3.png')
  await writeFile(imagePath, Buffer.from('fake-png-bytes'))
  const ctx = new Context()
  await assert.rejects(
    () => visionTextFromContext(ctx, 'fake-vision', 'vision-model')(imagePath, 'check'),
    /llm service is not mounted/,
  )
})
