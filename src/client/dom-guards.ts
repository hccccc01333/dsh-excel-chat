/**
 * DOM guards for the spreadsheet card.
 *
 * x-data-spreadsheet (1.1.9) takes over the page in three ways that are fine on
 * a full-page demo and wrong inside a chat column: it binds global listeners it
 * never releases, it treats "a click landed here" as permanent ownership of the
 * keyboard, and it focuses its editor with a plain `focus()` that scrolls every
 * ancestor. These helpers keep each of those local to the grid.
 *
 * No React in here on purpose: the behaviour is testable with plain jsdom.
 */

/**
 * The fourth slot records where the handler was registered. x-data-spreadsheet
 * binds `xclickoutside` on `document.body`, not on `window`; recording only the
 * type/listener made unmount remove it from the wrong target and leak it.
 */
export type WindowListener = [
  string,
  EventListenerOrEventListenerObject,
  boolean | AddEventListenerOptions | undefined,
  ('window' | 'body')?,
]

/** Anything a person types into — the grid's global shortcuts must stay out of it. */
export const TYPING_SELECTOR = 'input, textarea, select, [contenteditable]:not([contenteditable="false"])'

/**
 * Record the window listeners added while `build()` runs, so unmount can remove
 * them. The library has no `destroy()`; `el.innerHTML = ''` leaves
 * resize/click/paste/keydown registered, reloading a canvas that is no longer
 * in the document.
 */
export function captureWindowListeners<T>(build: () => T): { value: T; listeners: WindowListener[] } {
  const listeners: WindowListener[] = []
  const windowNative = window.addEventListener
  const bodyNative = document.body.addEventListener
  const patch = <T extends EventTarget>(target: T, native: T['addEventListener'], source: 'window' | 'body'): void => {
    ;(target as { addEventListener: unknown }).addEventListener = function (
      this: T,
      type: string,
      listener: EventListenerOrEventListenerObject,
      options?: boolean | AddEventListenerOptions,
    ): void {
      listeners.push([type, listener, options, source])
      native.call(this, type, listener, options)
    }
  }
  patch(window, windowNative, 'window')
  patch(document.body, bodyNative, 'body')
  try {
    return { value: build(), listeners }
  } catch (error) {
    // A half-built grid still managed to register handlers; do not leave them.
    removeCapturedListeners(listeners)
    throw error
  } finally {
    window.addEventListener = windowNative
    document.body.addEventListener = bodyNative
  }
}

/** Remove captured handlers from the target they were actually added to. */
export function removeCapturedListeners(listeners: WindowListener[]): void {
  for (const [type, listener, options, source] of listeners) {
    const target: EventTarget = source === 'body' ? document.body : window
    target.removeEventListener(type, listener, options)
  }
}

/**
 * The library decides that "the keyboard belongs to the grid" from the last
 * window click, then handles arrows / ctrl+z / paste globally and calls
 * `preventDefault()`. One click on a cell was enough to stop the caret keys
 * from working in the composer — the card looked like it kept the focus.
 *
 * Keep the handlers registered (the click flag still needs updating) but skip
 * them when the event is aimed at a text field outside the grid. Inside the
 * grid, or on plain page chrome, they behave exactly as before.
 */
export function scopeGlobalTypingToGrid(listeners: WindowListener[], root: HTMLElement): void {
  for (let index = 0; index < listeners.length; index += 1) {
    const [type, listener, options, source] = listeners[index]!
    // Body handlers (clickoutside) stay where they are; only global keyboard
    // handlers on window need the text-field guard.
    if (source === 'body' || (type !== 'keydown' && type !== 'paste')) continue
    window.removeEventListener(type, listener, options)
    const guarded = (event: Event): void => {
      const target = event.target
      if (target instanceof Element && !root.contains(target) && target.closest(TYPING_SELECTOR) !== null) return
      if (typeof listener === 'function') listener.call(window, event)
      else listener.handleEvent(event)
    }
    window.addEventListener(type, guarded, options)
    listeners[index] = [type, guarded, options]
  }
}

/**
 * Every focusable element the grid owns must keep focus local: a plain `focus()`
 * scrolls all ancestors to reveal the target, which yanks the conversation
 * column around whenever a cell is selected or edited.
 *
 * The editor textarea is only half of it. The library also focuses a 0-width
 * hidden input on **every** cell selection (`Selector.setAreaOffset`), so
 * clicking around a card repeatedly pulls focus into it and can drag the
 * surrounding scroll container. Patch both, and leave focusables outside the
 * grid (the chat composer) untouched.
 */
export function keepGridFocusInPlace(root: HTMLElement): void {
  const focusables = root.querySelectorAll<HTMLElement>(TYPING_SELECTOR)
  for (const focusable of focusables) {
    const nativeFocus = focusable.focus.bind(focusable)
    focusable.focus = (options?: FocusOptions): void => nativeFocus({ preventScroll: true, ...options })
  }
}
