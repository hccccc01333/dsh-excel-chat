#!/bin/bash
# Run ONLY the given corpus offsets (one task each); no pacing for skipped ids.
# Usage: bash scripts/bench-run-missing.sh 7 23 28 38 52 90 95
set -u
cd "$(dirname "$0")/.."
BAI_API=$("$SYSTEMROOT/System32/WindowsPowerShell/v1.0/powershell.exe" -NoProfile -Command "[Environment]::GetEnvironmentVariable('BAI_API','Machine')" | tr -d '\r\n')
if [ -z "$BAI_API" ]; then echo "BAI_API key not found" >&2; exit 1; fi
export BAI_API LLM_PROVIDER=bai BAI_MODEL=glm-5.3-flash LLM_BENCH_OUT=benchmarks/glm-5.3-flash-100.jsonl
echo "key loaded (len ${#BAI_API})" >&2
for offset in "$@"; do
  echo "=== task offset=$offset (checkpoint: $(wc -l < benchmarks/glm-5.3-flash-100.jsonl)/100) $(date +%H:%M:%S) ===" >&2
  LLM_BENCH_OFFSET=$offset LLM_BENCH_SAMPLE=1 \
    node tests/invoke-llm-benchmark.ts > "benchmarks/single-$offset.json" 2> "benchmarks/single-$offset.log"
  echo "offset $offset exit=$? $(date +%H:%M:%S)" >&2
  sleep 30
done
echo "MISSING DONE"
