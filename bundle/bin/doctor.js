#!/usr/bin/env node
import { readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { join, resolve as resolvePath } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { runDoctorChecks } from '../dist/doctor.js'

const root = resolvePath(fileURLToPath(new URL('..', import.meta.url)))
const args = process.argv.slice(2)
const profileIndex = args.indexOf('--profile')
const profileDirs = profileIndex >= 0 && args[profileIndex + 1] ? [args[profileIndex + 1]] : []

const checks = await runDoctorChecks({ profileDirs })
for (const check of checks) {
  const label = check.ok ? 'PASS' : 'FAIL'
  console.log(`[${label}] ${check.name}: ${check.detail}`)
}

let entryOk = false
let entry
try {
  entry = await import('../dist/index.js')
  entryOk = typeof entry.apply === 'function' && typeof entry.name === 'string'
  console.log(`[${entryOk ? 'PASS' : 'FAIL'}] bundle-entry: 插件入口可加载（${entry.name ?? 'unknown'}）`)
  if (!entryOk) process.exitCode = 1
} catch (error) {
  console.log(`[FAIL] bundle-entry: 插件入口加载失败：${error instanceof Error ? error.message : String(error)}`)
  process.exitCode = 1
}

// Registration self-test.
//
// The host owns `apply()`, so when a host refuses to bring the plugin up there is
// usually nothing left to inspect: no tools, no error, no trace. Running the same
// `apply()` against a recording stub here turns that into a fact on the command
// line — how many tools register, and the exact error for any that do not. This
// is the one channel that still works when the plugin never comes up in the host.
if (entryOk) {
  const registered = []
  const problems = []
  const host = {
    logger: {
      info: () => {},
      warn: (line) => problems.push(String(line)),
      error: (line) => problems.push(String(line)),
    },
    get: () => undefined,
    effect: (callback) => {
      const result = callback()
      return typeof result === 'function' ? result : () => {}
    },
    tools: {
      register: (definition) => {
        registered.push(definition?.name ?? '<unnamed>')
        return () => {}
      },
    },
    systemPrompt: { section: () => () => {} },
  }
  let threw = null
  try {
    entry.apply(host)
  } catch (error) {
    threw = error
  }
  const ok = registered.length > 0 && threw === null && problems.length === 0
  console.log(
    `[${ok ? 'PASS' : 'FAIL'}] tool-registration: 注册 ${registered.length} 个工具` +
      (problems.length > 0 ? `，${problems.length} 个注册失败` : ''),
  )
  for (const line of problems) console.log(`       ${line}`)
  if (threw) {
    console.log(`       apply() 抛出：${threw instanceof Error ? (threw.stack ?? threw.message) : String(threw)}`)
  }
  if (!ok) process.exitCode = 1
}

// Which copy of the kernel tool API the plugin actually resolves.
//
// A peer range that does not cover the running kernel makes a package manager
// install a second, nested copy of `@deepseek-ai/dsh-tools`, so the plugin and
// the host can end up building and registering tools through different copies.
// Printing both versions makes that mismatch visible instead of invisible.
function toolsApiVersion(fromDirectory) {
  try {
    const base = fromDirectory ? pathToFileURL(join(fromDirectory, 'noop.js')) : import.meta.url
    const require = createRequire(base)
    return require(require.resolve('@deepseek-ai/dsh-tools/package.json')).version
  } catch {
    return null
  }
}

const ownTools = toolsApiVersion(null)
console.log(`[INFO] kernel-api: 插件解析到 @deepseek-ai/dsh-tools@${ownTools ?? '（无法解析）'}`)
for (const dir of profileDirs) {
  const profileTools = toolsApiVersion(dir)
  if (profileTools === null) continue // a profile without the kernel has nothing to compare
  const same = profileTools === ownTools
  console.log(
    `[${same ? 'PASS' : 'WARN'}] kernel-api: 配置目录解析到 @${profileTools}` +
      (same ? '（与插件一致）' : '（与插件不同：插件与宿主可能不是同一份工具 API）'),
  )
}

// Will dsh actually mount this bundle?
//
// Since 0.2.0 a launcher checks every `@deepseek-ai/dsh*` peer in the bundle's
// package.json against the running dsh version, and a bundle that fails is
// dropped into the profile's `skippedBundles` — the layer never applies, the
// rows never mount, and **nothing is printed**. The plugin then looks like it
// simply is not there, which is exactly what issue #5 reported.
//
// Rather than reimplement that rule, ask the host: `evaluatePluginCompatibility`
// is exported, so the doctor can run the identical check and say in advance
// whether dsh will silently skip this bundle.
const ownManifest = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'))
let compatibilityChecked = 0
let compatibilityUnavailable = false
for (const dir of profileDirs) {
  let boot
  try {
    const require = createRequire(pathToFileURL(join(dir, 'noop.js')))
    boot = await import(pathToFileURL(require.resolve('@deepseek-ai/dsh-app-boot')).href)
  } catch {
    continue // no kernel in this profile, so there is nothing to be checked against
  }
  if (typeof boot.evaluatePluginCompatibility !== 'function' || typeof boot.getDshRuntimeVersion !== 'function') {
    // Kernels before 0.2.0 have no such check, so there is nothing to predict.
    compatibilityUnavailable = true
    continue
  }
  try {
    const runtime = boot.getDshRuntimeVersion()
    const exemptions = boot.readProfileCompatibility(dir).exemptions ?? {}
    const issue = boot.evaluatePluginCompatibility(ownManifest, exemptions, runtime)
    compatibilityChecked += 1
    if (issue === undefined) {
      console.log(`[PASS] bundle-compatibility: dsh ${runtime} 会挂载本插件（peer 区间覆盖该内核）`)
      continue
    }
    const exempted = issue.exempted ? '已豁免' : '未豁免'
    console.log(`[FAIL] bundle-compatibility: dsh ${runtime} 会【静默跳过】本插件（豁免：${exempted}）`)
    console.log(`       ${boot.pluginCompatibilityWarning(issue).split('.')[0]}`)
    if (!issue.exempted) process.exitCode = 1
  } catch (error) {
    console.log(`[WARN] bundle-compatibility: 无法完成检查：${error instanceof Error ? error.message : String(error)}`)
  }
}
if (compatibilityUnavailable && compatibilityChecked === 0) {
  console.log('[INFO] bundle-compatibility: 这些配置目录的内核早于 0.2.0，没有 bundle 兼容性检查，跳过')
}

if (checks.some((check) => !check.ok)) process.exitCode = 1
const passed = checks.filter((check) => check.ok).length + (entryOk ? 1 : 0)
console.log(`检查完成：${passed}/${checks.length + 1} 项通过（根目录 ${root}）`)
