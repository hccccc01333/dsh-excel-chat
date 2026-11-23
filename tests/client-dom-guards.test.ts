import { test } from 'node:test'
import assert from 'node:assert/strict'
import { JSDOM } from 'jsdom'
import {
  captureWindowListeners,
  keepGridFocusInPlace,
  removeCapturedListeners,
  scopeGlobalTypingToGrid,
  type WindowListener,
} from '../src/client/dom-guards.ts'

const PAGE = `<!doctype html><html><body>
  <div id="grid">
    <textarea id="editor"></textarea>
    <input id="hidden-input" style="width:0">
  </div>
  <input id="composer">
</body></html>`

/**
 * The guards read `window` / `Element` at call time, not at import time, so a
 * fresh realm per test is enough — no module reloading needed.
 */
function install(): { window: JSDOM['window']; document: Document } {
  const dom = new JSDOM(PAGE)
  const shared = globalThis as unknown as Record<string, unknown>
  shared.window = dom.window
  shared.document = dom.window.document
  shared.Element = dom.window.Element
  shared.HTMLTextAreaElement = dom.window.HTMLTextAreaElement
  shared.Event = dom.window.Event
  shared.KeyboardEvent = dom.window.KeyboardEvent
  return { window: dom.window, document: dom.window.document as Document }
}

test('captureWindowListeners records the window handlers and lets go of them', () => {
  const { window: win } = install()
  let calls = 0
  const handler = (): void => { calls += 1 }
  const built = captureWindowListeners(() => {
    win.addEventListener('resize', handler)
    win.addEventListener('keydown', handler)
    return 'grid'
  })
  assert.equal(built.value, 'grid')
  assert.deepEqual(built.listeners.map(([type]) => type), ['resize', 'keydown'])
  win.dispatchEvent(new win.Event('resize'))
  assert.equal(calls, 1)
  for (const [type, listener, options] of built.listeners) win.removeEventListener(type, listener, options)
  win.dispatchEvent(new win.Event('resize'))
  assert.equal(calls, 1)
})

test('body handlers are captured and removed from the body, not from window', () => {
  // x-data-spreadsheet binds `xclickoutside` on document.body. The old capture
  // only patched window, so unmount removed a non-existent window handler and
  // left the body one firing on a canvas that was no longer in the page.
  const { window: win, document: doc } = install()
  let calls = 0
  const handler = (): void => { calls += 1 }
  const built = captureWindowListeners(() => {
    doc.body.addEventListener('click', handler)
    return 'grid'
  })
  assert.equal(built.listeners[0]![3], 'body')
  doc.body.dispatchEvent(new win.Event('click'))
  assert.equal(calls, 1)
  removeCapturedListeners(built.listeners)
  doc.body.dispatchEvent(new win.Event('click'))
  assert.equal(calls, 1)
})

test('a build that throws leaves no window handler behind', () => {
  const { window: win } = install()
  let calls = 0
  assert.throws(() => captureWindowListeners(() => {
    win.addEventListener('resize', () => { calls += 1 })
    throw new Error('boom')
  }), /boom/)
  win.dispatchEvent(new win.Event('resize'))
  assert.equal(calls, 0)
})

test('the grid stops handling keys aimed at a text field outside it', () => {
  const { window: win, document: doc } = install()
  const grid = doc.getElementById('grid') as HTMLElement
  const composer = doc.getElementById('composer') as HTMLInputElement
  const seen: string[] = []
  const handler = (event: Event): void => {
    seen.push((event.target as HTMLElement).id)
    event.preventDefault()
  }
  const listeners: WindowListener[] = [['keydown', handler, undefined]]
  scopeGlobalTypingToGrid(listeners, grid)
  const press = (target: EventTarget): Event => {
    const event = new win.KeyboardEvent('keydown', { bubbles: true, cancelable: true })
    target.dispatchEvent(event)
    return event
  }

  const inComposer = press(composer)
  assert.deepEqual(seen, [])
  assert.equal(inComposer.defaultPrevented, false)

  const inEditor = press(doc.getElementById('editor') as HTMLElement)
  assert.deepEqual(seen, ['editor'])
  assert.equal(inEditor.defaultPrevented, true)

  // Plain page chrome keeps the library's shortcuts (Ctrl+Z, arrows) working.
  press(doc.body)
  assert.deepEqual(seen, ['editor', ''])
})

test('paste is scoped the same way, and other window handlers are left alone', () => {
  const { window: win, document: doc } = install()
  const grid = doc.getElementById('grid') as HTMLElement
  const composer = doc.getElementById('composer') as HTMLInputElement
  const seen: string[] = []
  const paste = (event: Event): void => { seen.push(`paste:${(event.target as HTMLElement).id}`) }
  const resize = (): void => { seen.push('resize') }
  // What the sheet constructor does: bind on window, then hand us the list.
  win.addEventListener('paste', paste)
  win.addEventListener('resize', resize)
  const listeners: WindowListener[] = [
    ['paste', paste, undefined],
    ['resize', resize, undefined],
  ]
  scopeGlobalTypingToGrid(listeners, grid)

  composer.dispatchEvent(new win.Event('paste', { bubbles: true }))
  doc.getElementById('editor')!.dispatchEvent(new win.Event('paste', { bubbles: true }))
  win.dispatchEvent(new win.Event('resize'))

  assert.deepEqual(seen, ['paste:editor', 'resize'])
  // The resize entry must still be the original function, not a wrapper.
  assert.equal(listeners[1]![1], resize)
})

test('the editor and the hidden selection input focus without scrolling ancestors', () => {
  const { document: doc } = install()
  const grid = doc.getElementById('grid') as HTMLElement
  const editor = doc.getElementById('editor') as HTMLTextAreaElement
  const hidden = doc.getElementById('hidden-input') as HTMLInputElement
  const composer = doc.getElementById('composer') as HTMLInputElement
  const calls: Array<FocusOptions | undefined> = []
  for (const el of [editor, hidden, composer]) {
    el.focus = ((options?: FocusOptions): void => { calls.push(options) }) as typeof el.focus
  }
  keepGridFocusInPlace(grid)
  editor.focus()
  // x-data-spreadsheet's Selector.setAreaOffset focuses this on every selection.
  hidden.focus()
  assert.deepEqual(calls, [{ preventScroll: true }, { preventScroll: true }])

  // Focusables outside the grid are not patched.
  composer.focus()
  assert.deepEqual(calls, [{ preventScroll: true }, { preventScroll: true }, undefined])
})

test('a card without a focusable element does not throw', () => {
  const { document: doc } = install()
  const empty = doc.createElement('div')
  keepGridFocusInPlace(empty)
})
