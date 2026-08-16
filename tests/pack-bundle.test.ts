import { test } from 'node:test'
import assert from 'node:assert/strict'
import { execFile } from 'node:child_process'
import { readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { promisify } from 'node:util'

const execFileAsync = promisify(execFile)
const bundleDir = fileURLToPath(new URL('../bundle', import.meta.url))

const HOST_PACKAGES = [
  '@deepseek-ai/dsh-attachment',
  '@deepseek-ai/dsh-llm',
  '@deepseek-ai/dsh-system-prompt',
  '@deepseek-ai/dsh-tools',
]

test('DSH host packages are peerDependencies, never bundle dependencies', async () => {
  const manifest = JSON.parse(await readFile(join(bundleDir, 'package.json'), 'utf8')) as {
    dependencies?: Record<string, string>
    peerDependencies?: Record<string, string>
  }
  for (const name of HOST_PACKAGES) {
    assert.equal(manifest.dependencies?.[name], undefined, `${name} must not be a bundle dependency`)
    assert.ok(manifest.peerDependencies?.[name], `${name} must be a peerDependency`)
  }
  for (const [name, range] of Object.entries(manifest.dependencies ?? {})) {
    assert.ok(!name.startsWith('@deepseek-ai/dsh-'), `bundle dependency ${name} must not be a DSH host package`)
  }
})

test('npm pack dry-run includes dist, patch, README, and LICENSE', async () => {
  const manifest = JSON.parse(await readFile(join(bundleDir, 'package.json'), 'utf8')) as {
    name: string
    version: string
  }
  const npmCommand = process.platform === 'win32' ? 'npm.cmd' : 'npm'
  const { stdout } = await execFileAsync(npmCommand, ['pack', '--dry-run', '--json'], {
    cwd: bundleDir,
    shell: true,
  })
  const [result] = JSON.parse(stdout) as Array<{
    name: string
    version: string
    files: Array<{ path: string }>
  }>
  const paths = result.files.map((file) => file.path)
  for (const expected of ['dist/index.js', 'dist/index.d.ts', 'cordis.patch.yml', 'README.md', 'LICENSE']) {
    assert.ok(paths.includes(expected), `tarball missing ${expected}`)
  }
  assert.equal(result.name, manifest.name)
  assert.equal(result.version, manifest.version)
})
