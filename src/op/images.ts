/**
 * Embedding images, including reading their intrinsic size from the bytes.
 *
 * The image parts are written at the XML layer rather than through ExcelJS, so the
 * extent maths lives here: a caller may give a pixel width, a height, both or
 * neither, and the aspect ratio has to survive all four cases. Split out of
 * `operations.ts` unchanged.
 */
import ExcelJS from 'exceljs'
import { readFile } from 'node:fs/promises'
import { columnToNumber, parseCellId } from '../formula.ts'
import { findSheet } from './core.ts'
import type { ExcelOperation } from '../operation-types.ts'
import { t } from '../i18n.ts'

const IMAGE_EXTENSIONS = new Set(['png', 'jpeg', 'jpg', 'gif'])

const MAX_IMAGE_BYTES = 20 * 1024 * 1024


export async function insertImageIntoSheet(
  workbook: ExcelJS.Workbook,
  options: Extract<ExcelOperation, { op: 'insertImage' }>,
): Promise<void> {
  const parsed = parseCellId(options.cell)
  const sheet = findSheet(workbook, parsed.sheet)
  if (!sheet) throw new Error(`sheet not found: ${parsed.sheet}`)
  if (options.file && options.base64) throw new Error('insertImage takes either file or base64, not both')
  if (!options.file && !options.base64) throw new Error('insertImage requires file or base64')

  const { buffer, extension } = await readImageSource(options)
  if (buffer.byteLength > MAX_IMAGE_BYTES) {
    throw new Error(`insertImage refuses images over ${MAX_IMAGE_BYTES / 1024 / 1024}MB (got ${Math.round(buffer.byteLength / 1024 / 1024)}MB)`)
  }

  // exceljs ships an older Buffer typing that Node 22's generic Buffer no longer
  // satisfies; the bytes are identical, so narrow it at the boundary.
  const imageId = workbook.addImage({ buffer, extension } as unknown as Parameters<ExcelJS.Workbook['addImage']>[0])
  // exceljs requires `ext` both in its types and at render time — omitting it
  // throws while writing the drawing XML.
  const ext = resolveImageExtent(options, buffer)
  if (!ext) {
    throw new Error('insertImage could not read the image size; pass width and height explicitly')
  }
  sheet.addImage(imageId, {
    tl: { col: columnToNumber(parsed.column) - 1, row: parsed.row - 1 },
    ext,
  })
}


async function readImageSource(
  options: Extract<ExcelOperation, { op: 'insertImage' }>,
): Promise<{ buffer: Buffer; extension: 'png' | 'jpeg' | 'gif' }> {
  if (options.file) {
    // Check the extension first so a typo'd format fails fast, before touching disk.
    const extension = toImageExtension(options.file)
    // readFile yields Buffer<ArrayBufferLike> while exceljs types want the narrower
    // Buffer. Identical bytes, so a cast beats copying up to 20MB for nothing.
    return { buffer: await readFile(options.file) as Buffer, extension }
  }
  const raw = options.base64!
  const dataUri = /^data:image\/([a-z0-9]+);base64,(.*)$/is.exec(raw)
  if (dataUri) {
    return {
      buffer: Buffer.from(dataUri[2]!.replace(/\s+/g, ''), 'base64'),
      extension: toImageExtension(dataUri[1]!),
    }
  }
  // Bare base64 carries no format hint; PNG is the safe default (and the caller
  // can pass a data URI whenever the payload is actually jpeg/gif).
  return { buffer: Buffer.from(raw.replace(/\s+/g, ''), 'base64'), extension: 'png' }
}

function toImageExtension(hint: string): 'png' | 'jpeg' | 'gif' {
  const extension = (/\.([a-z0-9]+)$/i.exec(hint)?.[1] ?? hint).toLowerCase()
  if (!IMAGE_EXTENSIONS.has(extension)) {
    throw new Error(`insertImage supports png/jpeg/gif, got "${extension}"`)
  }
  return extension === 'jpg' ? 'jpeg' : (extension as 'png' | 'jpeg' | 'gif')
}

/**
 * Intrinsic pixel size from the image header. PNG and GIF keep it at a fixed
 * offset; JPEG needs a walk to the first SOF segment. Returns null for anything
 * unrecognised, so the caller asks for an explicit size instead of guessing.
 */

function resolveImageExtent(
  options: { width?: number; height?: number },
  buffer: Buffer,
): { width: number; height: number } | null {
  const { width, height } = options
  if (width !== undefined && height !== undefined) return { width, height }
  const intrinsic = readImageSize(buffer)
  if (width === undefined && height === undefined) return intrinsic
  if (!intrinsic) return null
  if (width !== undefined) {
    return { width, height: scaleDimension(width, intrinsic.height, intrinsic.width) }
  }
  return { width: scaleDimension(height!, intrinsic.width, intrinsic.height), height: height! }
}


function scaleDimension(given: number, numerator: number, denominator: number): number {
  if (denominator <= 0) return given
  return Math.max(1, Math.round((given * numerator) / denominator))
}

function readImageSize(buffer: Buffer): { width: number; height: number } | null {
  // PNG: IHDR chunk holds big-endian u32 width/height at offsets 16 and 20.
  if (buffer.length >= 24 && buffer.readUInt32BE(0) === 0x89504e47) {
    return { width: buffer.readUInt32BE(16), height: buffer.readUInt32BE(20) }
  }
  // GIF: logical screen descriptor, little-endian u16 at offsets 6 and 8.
  if (buffer.length >= 10 && buffer.toString('latin1', 0, 3) === 'GIF') {
    return { width: buffer.readUInt16LE(6), height: buffer.readUInt16LE(8) }
  }
  // JPEG: walk marker segments to the frame header (SOF0-SOF15, minus DHT/JPG/DAC).
  if (buffer.length >= 4 && buffer[0] === 0xff && buffer[1] === 0xd8) {
    let offset = 2
    while (offset + 9 <= buffer.length) {
      if (buffer[offset] !== 0xff) { offset++; continue }
      const marker = buffer[offset + 1]!
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        return { width: buffer.readUInt16BE(offset + 7), height: buffer.readUInt16BE(offset + 5) }
      }
      const segmentLength = buffer.readUInt16BE(offset + 2)
      if (segmentLength < 2) break
      offset += 2 + segmentLength
    }
  }
  return null
}
