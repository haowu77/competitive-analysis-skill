# competitive-analysis-skill

Build industry-standard competitor benchmark workbooks (`.xlsx`) for software products.

Output includes 5 sheets:
- `Summary`
- `Benchmark`
- `Feature-Matrix`
- `Pricing-GTM`
- `Sources`

Supports multilingual output (`en`, `zh`, `ja`, `ko`, `es`, `fr`, `de`) with automatic language detection.

## Quick start

```bash
python3 scripts/build_benchmark_xlsx.py \
  --output examples/benchmark-zh.xlsx \
  --brief "这是一个AI meal planning产品需求"
```

Force English output:

```bash
python3 scripts/build_benchmark_xlsx.py \
  --output examples/benchmark-en.xlsx \
  --brief "这是一个AI meal planning产品需求" \
  --lang en
```

## CLI

```bash
python3 scripts/build_benchmark_xlsx.py \
  --output /abs/path/competitive-benchmark-YYYY-MM-DD.xlsx \
  --brief "your requirement brief" \
  --project-path /abs/path/to/repo \
  --region global \
  --top-n 8 \
  --period-months 24 \
  --lang auto \
  --lang-source both
```

## Project layout
- `SKILL.md`: skill contract and workflow
- `agents/openai.yaml`: UI metadata
- `scripts/`: benchmark generator
- `references/`: methodology and quality standards
- `examples/`: runnable commands
- `tests/`: regression tests

## Quality gates

```bash
python3 -m unittest discover -s tests -p 'test_*.py'
```

## License
MIT
