import { access, mkdtemp, readFile, readdir } from 'node:fs/promises'
import { homedir, tmpdir } from 'node:os'
import { join } from 'node:path'
import ExcelJS from 'exceljs'
import { validate } from './validator.ts'
import { readWorkbookCells } from './workbook.ts'

export interface DoctorCheck {
  name: string
  ok: boolean
  detail: string
}

/** Host packages that must stay peer/dev-only; a profile installing them as
 * dependencies shadows the harness host and breaks tool dispatch (Issue #1). */
export const HOST_PACKAGES = [
  '@deepseek-ai/cordis',
  '@deepseek-ai/dsh-tools',
  '@deepseek-ai/dsh-llm',
  '@deepseek-ai/dsh-system-prompt',
  '@deepseek-ai/dsh-attachment',
]

/**
 * Install self-check (doctor): environment, host-package isolation, and a
 * real engine smoke test that creates a temp workbook and reads/validates it.
 * The smoke test intentionally bypasses the harness scheduler, so it works
 * even when host Symbol identity is broken.
 */
export async function runDoctorChecks(options: { profileDirs?: string[] } = {}): Promise<DoctorCheck[]> {
  const checks: DoctorCheck[] = []

  const nodeMajor = Number(process.versions.node.split('.')[0] ?? 0)
  checks.push({
    name: 'node-version',
    ok: nodeMajor >= 22,
    detail: `node ${process.versions.node}（要求 ^22.19 || >=24）`,
  })

  const profileDirs = options.profileDirs ?? []
  if (profileDirs.length === 0) {
    const root = join(homedir(), '.dsh', 'profiles')
    try {
      const entries = await readdir(root, { withFileTypes: true })
      profileDirs.push(...entries.filter((entry) => entry.isDirectory()).map((entry) => join(root, entry.name)))
    } catch {
      // no profiles on this machine; the isolation check below is skipped
    }
  }
  if (profileDirs.length === 0) {
    checks.push({ name: 'profile-host-isolation', ok: true, detail: '未发现 dsh profile，跳过宿主包隔离检查' })
  } else {
    for (const dir of profileDirs) {
      const packagePath = join(dir, 'package.json')
      try {
        await access(packagePath)
      } catch {
        continue
      }
      const manifest = JSON.parse(await readFile(packagePath, 'utf8')) as { dependencies?: Record<string, string> }
      const deps = manifest.dependencies ?? {}
      const violations = HOST_PACKAGES.filter((name) => deps[name] !== undefined)
      checks.push({
        name: `profile-host-isolation:${dir}`,
        ok: violations.length === 0,
        detail: violations.length === 0
          ? `${dir}：宿主包未出现在 dependencies（符合隔离要求）`
          : `${dir}：宿主包被安装为 dependencies：${violations.join('、')}（会导致所有工具调用失败）`,
      })
    }
  }

  try {
    const dir = await mkdtemp(join(tmpdir(), 'dsh-doctor-'))
    const path = join(dir, 'smoke.xlsx')
    const workbook = new ExcelJS.Workbook()
    const sheet = workbook.addWorksheet('订单')
    sheet.getCell('A1').value = '数量'
    sheet.getCell('B1').value = '金额'
    sheet.getCell('A2').value = 2
    sheet.getCell('B2').value = { formula: 'A2*10' }
    await workbook.xlsx.writeFile(path)
    const cells = await readWorkbookCells(await readFile(path))
    const result = validate(cells)
    const hasFormula = Object.values(cells).some((value) => String(value).startsWith('='))
    checks.push({
      name: 'engine-smoke',
      ok: hasFormula && typeof result.formulaCount === 'number' && result.formulaCount > 0,
      detail: `临时工作簿读取+公式体检正常（公式 ${result.formulaCount} 个，异常 ${result.anomalies.length} 个）`,
    })
  } catch (error) {
    checks.push({
      name: 'engine-smoke',
      ok: false,
      detail: `引擎冒烟失败：${error instanceof Error ? error.message : String(error)}`,
    })
  }

  return checks
}
