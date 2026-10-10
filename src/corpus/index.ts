import type { FileBenchmarkTask } from '../file-benchmark.ts'
import { editingTasks } from './editing.ts'
import { structureTasks } from './structure.ts'
import { crossSheetTasks } from './cross-sheet.ts'
import { missedEditTasks } from './missed-edits.ts'
import { analysisTasks } from './analysis.ts'
import { formulaTasks } from './formula.ts'
import { workflowTasks } from './workflow.ts'

/** Realistic offline task corpus (ExcelBench lite). */
export const corpusTasks: FileBenchmarkTask[] = [
  ...editingTasks,
  ...structureTasks,
  ...crossSheetTasks,
  ...missedEditTasks,
  ...analysisTasks,
  ...formulaTasks,
  ...workflowTasks,
]
