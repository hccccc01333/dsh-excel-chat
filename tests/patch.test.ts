import { test } from 'node:test'
import assert from 'node:assert/strict'
import { applyPatches, revertPatches, type CellPatch } from '../src/patch.ts'

const patch: CellPatch = {
  id: 'D4',
  kind: 'formula',
  oldValue: '=B4-C3',
  newValue: '=B4-C4',
}

test('applyPatches and revertPatches round-trip', () => {
  const cells = { D4: '=B4-C3' }
  const applied = applyPatches(cells, [patch])
  assert.equal(applied.D4, '=B4-C4')
  assert.deepEqual(revertPatches(applied, [patch]), cells)
})

test('applyPatches rejects stale preconditions', () => {
  assert.throws(
    () => applyPatches({ D4: '=X1' }, [patch]),
    /precondition failed for D4/,
  )
})
