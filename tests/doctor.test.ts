import { test } from 'node:test'
import assert from 'node:assert/strict'
import { mkdir, mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { runDoctorChecks } from '../src/doctor.ts'

test('runDoctorChecks flags profiles with host packages in dependencies', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'vera-doctor-'))
  const bad = join(dir, 'bad')
  const good = join(dir, 'good')
  await mkdir(bad, { recursive: true })
  await mkdir(good, { recursive: true })
  await writeFile(join(bad, 'package.json'), JSON.stringify({
    dependencies: { '@deepseek-ai/dsh-tools': '0.0.1-rc.1', 'dsh-excel-chat': '0.34.0' },
  }))
  await writeFile(join(good, 'package.json'), JSON.stringify({
    peerDependencies: { '@deepseek-ai/dsh-tools': '^0.1.0-rc.6' },
  }))

  const checks = await runDoctorChecks({ profileDirs: [bad, good] })
  const badCheck = checks.find((check) => check.name.includes('bad'))
  const goodCheck = checks.find((check) => check.name.includes('good'))
  assert.ok(badCheck)
  assert.equal(badCheck.ok, false)
  assert.match(badCheck.detail, /@deepseek-ai\/dsh-tools/)
  assert.ok(goodCheck)
  assert.equal(goodCheck.ok, true)
})

test('runDoctorChecks passes the engine smoke test', async () => {
  const checks = await runDoctorChecks({ profileDirs: [] })
  const engine = checks.find((check) => check.name === 'engine-smoke')
  assert.ok(engine)
  assert.equal(engine.ok, true)
  assert.match(engine.detail, /公式 1 个/)
})
