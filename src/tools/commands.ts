/**
 * The three slash commands.
 *
 * They act on a local file in place rather than producing a new one, which is why
 * they are commands and not tools: the user typed a path and expects that file to
 * change. Split out of `index.ts` unchanged.
 */
import type { Context } from '@deepseek-ai/cordis'
import { t } from '../i18n.ts'
import { applyInPlaceEdit, revertInPlaceEdit } from '../live-edit.ts'
import { runDoctorChecks } from '../doctor.ts'

export function registerCommands(ctx: Context): void {
const commands = ctx.get('commands') as {
  register(definition: {
    name: string
    description: string
    input?: { hint?: string }
    handler: (invocation: { rawInput: string }) => Promise<{ kind: 'success' | 'error'; text: string }> | { kind: 'success' | 'error'; text: string }
  }): () => void
} | undefined
if (commands) {
  ctx.effect(() => commands.register({
    name: 'excel-set',
    description: t('就地修改本地 Excel 文件的一个单元格（自动备份 .bak + 审计日志 + 公式体检）'),
    input: { hint: '{"path":"D:\\\\x.xlsx","cell":"Sheet1!A1","value":"..."}' },
    handler: async (invocation) => {
      try {
        const payload = JSON.parse(invocation.rawInput.trim()) as { path?: string; cell?: string; value?: string | number }
        if (!payload.path || !payload.cell || payload.value === undefined) {
          return { kind: 'error', text: t('需要 {"path","cell","value"}') }
        }
        const result = await applyInPlaceEdit(payload.path, payload.cell, payload.value)
        return {
          kind: 'success',
          text: t('已就地保存 {cell} = {value}（备份 {backup}，公式异常 {anomalies}）', { cell: result.cell, value: result.value, backup: result.backupPath, anomalies: result.anomalies }),
        }
      } catch (error) {
        return { kind: 'error', text: error instanceof Error ? error.message : String(error) }
      }
    },
  }), 'command:excel-set')
  ctx.effect(() => commands.register({
    name: 'excel-undo',
    description: t('回滚本地 Excel 文件最近一次就地修改'),
    input: { hint: '{"path":"D:\\\\x.xlsx"}' },
    handler: async (invocation) => {
      try {
        const payload = JSON.parse(invocation.rawInput.trim()) as { path?: string }
        if (!payload.path) return { kind: 'error', text: t('需要 {"path"}') }
        const outcome = await revertInPlaceEdit(payload.path)
        return { kind: outcome.restored ? 'success' : 'error', text: outcome.message }
      } catch (error) {
        return { kind: 'error', text: error instanceof Error ? error.message : String(error) }
      }
    },
  }), 'command:excel-undo')
  ctx.effect(() => commands.register({
    name: 'excel-doctor',
    description: t('安装自检：检查宿主包隔离、Node 版本与 Excel 引擎冒烟'),
    input: { hint: t('可选：{"profile":"D:\\\\profile-dir"}') },
    handler: async (invocation) => {
      try {
        const payload = JSON.parse(invocation.rawInput.trim() || '{}') as { profile?: string }
        const checks = await runDoctorChecks({ profileDirs: payload.profile ? [payload.profile] : [] })
        const lines = checks.map((check) => `${check.ok ? '✅' : '❌'} ${check.name}: ${check.detail}`)
        const failed = checks.filter((check) => !check.ok).length
        return { kind: failed === 0 ? 'success' : 'error', text: lines.join('\n') }
      } catch (error) {
        return { kind: 'error', text: error instanceof Error ? error.message : String(error) }
      }
    },
  }), 'command:excel-doctor')
}
}
