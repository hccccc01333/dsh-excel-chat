import assert from 'node:assert/strict'
import { test } from 'node:test'
import { apply } from '../src/index.ts'
import { guardedContext, registrationSummary } from '../src/registration.ts'

/**
 * A host stub that behaves like the real one where it matters: `effect` runs its
 * callback synchronously, which is what cordis does and what makes a throw
 * inside a registration escape `apply`.
 */
function fakeHost(options: { reject?: string[]; onError?: string[] } = {}) {
  const registered: string[] = []
  const reject = new Set(options.reject ?? [])
  const host = {
    logger: {
      info: () => {},
      warn: () => {},
      error: (line: string) => options.onError?.push(line),
    },
    get: () => undefined,
    effect: (callback: () => unknown) => {
      const result = callback()
      return typeof result === 'function' ? result : () => {}
    },
    tools: {
      register: (definition: { name?: string }) => {
        const name = String(definition?.name)
        if (reject.has(name)) throw new Error(`the host refused ${name}`)
        registered.push(name)
        return () => {}
      },
    },
    systemPrompt: { section: () => () => {} },
  }
  return { host, registered }
}

/** Run `body` with console.error captured, so a test can assert what was reported. */
function captureConsoleError(body: () => void): string[] {
  const lines: string[] = []
  const original = console.error
  console.error = (...args: unknown[]) => {
    lines.push(args.map(String).join(' '))
  }
  try {
    body()
  } finally {
    console.error = original
  }
  return lines
}

test('a failing registration does not stop the ones after it', () => {
  const { host, registered } = fakeHost()
  const { ctx, report } = guardedContext(host as never)

  captureConsoleError(() => {
    ctx.effect(() => ctx.tools.register({ name: 'excel_a' } as never), 'tool:excel_a')
    ctx.effect(() => {
      throw new Error('malformed definition')
    }, 'tool:excel_b')
    ctx.effect(() => ctx.tools.register({ name: 'excel_c' } as never), 'tool:excel_c')
  })

  assert.deepEqual(registered, ['excel_a', 'excel_c'])
  assert.deepEqual(
    report.registered,
    ['excel_a', 'excel_c'],
    'a rejected tool must not be counted as registered',
  )
  assert.equal(report.failures.length, 1)
  assert.equal(report.failures[0].label, 'tool:excel_b')
  assert.match(report.failures[0].detail, /malformed definition/)
})

test('a throwing registration never escapes ctx.effect', () => {
  const { host } = fakeHost()
  const { ctx } = guardedContext(host as never)

  captureConsoleError(() => {
    // Without the guard this throw would leave `ctx.effect`, abort `apply`, and
    // make cordis dispose every effect the plugin registered.
    assert.doesNotThrow(() =>
      ctx.effect(() => {
        throw new Error('boom')
      }, 'tool:excel_b'),
    )
  })
})

test('the failure is reported on the host logger as well as the console', () => {
  const logged: string[] = []
  const { host } = fakeHost({ onError: logged })
  const { ctx } = guardedContext(host as never)

  const consoleLines = captureConsoleError(() => {
    ctx.effect(() => {
      throw new Error('host rejected the schema')
    }, 'tool:excel_trace')
  })

  assert.equal(logged.length, 1)
  assert.match(logged[0], /tool:excel_trace failed to register/)
  assert.match(logged[0], /host rejected the schema/)
  assert.deepEqual(consoleLines, logged, 'the console fallback carries the same line')
})

test('registrationSummary distinguishes a full load from a partial one', () => {
  assert.match(registrationSummary({ registered: ['a', 'b'], failures: [] }), /ready: 2 tools/)
  const partial = registrationSummary({
    registered: ['a'],
    failures: [{ label: 'tool:b', detail: 'nope' }],
  })
  assert.match(partial, /PARTIAL/)
  assert.match(partial, /1 tools registered, 1 failed -> tool:b/)
})

test('apply still brings up every other tool when the host rejects one', () => {
  const logged: string[] = []
  const { host, registered } = fakeHost({ reject: ['excel_trace'], onError: logged })

  const consoleLines = captureConsoleError(() => {
    apply(host as never)
  })

  assert.equal(registered.length, 24, 'the other tools must survive one rejection')
  assert.ok(!registered.includes('excel_trace'))
  assert.ok(registered.includes('excel_find_errors'), 'tools after the rejected one still register')
  assert.ok(registered.includes('excel_operate'))
  assert.ok(
    [...logged, ...consoleLines].some((line) => /PARTIAL/.test(line)),
    'a partial load must say so instead of looking like a clean start',
  )
})
