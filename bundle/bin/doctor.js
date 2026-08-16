#!/usr/bin/env node
import { resolve as resolvePath } from 'node:path'
import { fileURLToPath } from 'node:url'
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
try {
  const entry = await import('../dist/index.js')
  entryOk = typeof entry.apply === 'function' && typeof entry.name === 'string'
  console.log(`[${entryOk ? 'PASS' : 'FAIL'}] bundle-entry: 插件入口可加载（${entry.name ?? 'unknown'}）`)
  if (!entryOk) process.exitCode = 1
} catch (error) {
  console.log(`[FAIL] bundle-entry: 插件入口加载失败：${error instanceof Error ? error.message : String(error)}`)
  process.exitCode = 1
}

if (checks.some((check) => !check.ok)) process.exitCode = 1
const passed = checks.filter((check) => check.ok).length + (entryOk ? 1 : 0)
console.log(`检查完成：${passed}/${checks.length + 1} 项通过（根目录 ${root}）`)
