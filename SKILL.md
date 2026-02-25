---
name: competitive-analysis
description: Build industry-standard competitor benchmark outputs for software products. Use when users ask for competitor analysis, benchmarking, product comparison matrices, or “best alternatives” research for a project or requirement. Discover and classify competitors (direct/adjacent/substitute), score them with weighted dimensions, attach evidence links, and export decision-ready XLSX tables.
---

# Competitive Analysis Skill

## Workflow
1. Define scope from one input source:
- Repo context (`project-path`) or
- User requirement brief (`brief`).
2. Extract core JTBD, user segment, pricing context, and platform shape.
3. Discover candidates from multiple source types and classify into:
- `Direct`
- `Adjacent`
- `Substitute`
4. Keep the most relevant `top-n` competitors (default 8).
5. Fill scoring dimensions (1-5) with evidence-backed judgments.
6. Generate the benchmark workbook via script.

## Standard deliverable
Generate one XLSX file with exactly 5 sheets:
- `Summary` (default 3-row template: Market Definition / Positioning Statement / Strategic Implications)
- `Benchmark`
- `Feature-Matrix`
- `Pricing-GTM`
- `Sources`

## Multilingual output
The script supports language-aware output:
- If user input is Chinese, output workbook labels in Chinese.
- If user input is English, output workbook labels in English.
- If user input is Japanese, output workbook labels in Japanese.
- Also supports: `ko`, `es`, `fr`, `de`.

Language behavior:
- `--lang auto` (default): detect from input text.
- `--lang <code>`: force output language.
- `--lang-source brief|input|both`: choose language-detection source.

## Script
Run:

```bash
python3 scripts/build_benchmark_xlsx.py \
  --output /absolute/path/competitive-benchmark-YYYY-MM-DD.xlsx \
  --brief "your product requirement" \
  --project-path /abs/path/to/repo \
  --region global \
  --top-n 8 \
  --period-months 24 \
  --lang auto \
  --lang-source both
```

Force Chinese output:

```bash
python3 scripts/build_benchmark_xlsx.py \
  --output /absolute/path/competitive-benchmark-zh.xlsx \
  --brief "这是一个AI产品需求" \
  --lang zh
```

With prepared structured input:

```bash
python3 scripts/build_benchmark_xlsx.py \
  --output /absolute/path/competitive-benchmark.xlsx \
  --input-json /absolute/path/benchmark-input.json \
  --lang auto
```

## Input JSON schema
Top-level keys (all optional):
- `summary`: list of summary rows
- `benchmark`: list of competitor rows
- `feature_matrix`: list of feature rows
- `pricing_gtm`: list of pricing/GTM rows
- `sources`: list of evidence rows

The script maps keys by normalized aliases, so minor naming differences are tolerated (including localized header names).

## Quality rules
- Each competitor should have at least 3 sources.
- Each competitor should include at least one official source and one third-party source.
- Use absolute dates in `Published Date` and `Access Date`.
- Do not produce high threat conclusions without high-confidence evidence.

## References
- Methodology: `references/methodology.md`
- Scoring rubric: `references/scoring-rubric.md`
- Source-quality standards: `references/source-quality.md`
- i18n behavior: `references/i18n.md`
