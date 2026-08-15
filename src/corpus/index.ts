import type { FileBenchmarkTask } from '../file-benchmark.ts'
import { editingTasks } from './editing.ts'
import { analysisTasks } from './analysis.ts'
import { formulaTasks } from './formula.ts'
import { workflowTasks } from './workflow.ts'

/** Realistic offline task corpus (ExcelBench lite): 100 file-based scenarios. */
export const corpusTasks: FileBenchmarkTask[] = [
  ...editingTasks,
  ...analysisTasks,
  ...formulaTasks,
  ...workflowTasks,
]
