#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
python3 "$ROOT/scripts/build_benchmark_xlsx.py" \
  --output "$ROOT/examples/benchmark-zh.xlsx" \
  --brief "这是一个 AI meal planning 产品需求，目标用户为忙碌上班族。" \
  --lang auto \
  --lang-source both

echo "generated: $ROOT/examples/benchmark-zh.xlsx"
