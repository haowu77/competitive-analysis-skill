# competitive-analysis-skill

[![CI](https://github.com/haowu77/competitive-analysis-skill/actions/workflows/ci.yml/badge.svg)](https://github.com/haowu77/competitive-analysis-skill/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE)
[![Release](https://img.shields.io/github/v/release/haowu77/competitive-analysis-skill)](https://github.com/haowu77/competitive-analysis-skill/releases)

## 中文简介
`competitive-analysis-skill` 用于软件产品的标准化竞品分析输出，生成一份决策可用的 5-sheet 基准分析工作簿（`.xlsx`）：
- 摘要（Summary）
- 竞品基准（Benchmark）
- 功能矩阵（Feature-Matrix）
- 定价与GTM（Pricing-GTM）
- 证据来源（Sources）

核心方法：
- 竞品分层（Direct / Adjacent / Substitute）
- 加权评分（多维 1-5 + 总分）
- 证据可信度约束（官方源 + 第三方源）

## English Overview
`competitive-analysis-skill` generates an industry-standard competitor benchmark workbook (`.xlsx`) with 5 sheets:
- Summary
- Benchmark
- Feature-Matrix
- Pricing-GTM
- Sources

Method highlights:
- competitor classification (Direct / Adjacent / Substitute)
- weighted scoring (dimension scores + weighted total)
- evidence quality checks (official + third-party sources)

## Install / 安装

```bash
git clone https://github.com/haowu77/competitive-analysis-skill.git ~/.codex/skills/competitive-analysis-skill
```

## Quick Start / 快速开始
中文需求自动识别输出语言：

```bash
python3 scripts/build_benchmark_xlsx.py \
  --output examples/benchmark-zh.xlsx \
  --brief "这是一个AI meal planning产品需求"
```

强制英文输出：

```bash
python3 scripts/build_benchmark_xlsx.py \
  --output examples/benchmark-en.xlsx \
  --brief "这是一个AI meal planning产品需求" \
  --lang en
```

## Input / Output
输入来源可二选一：
- `--brief`（自然语言需求）
- `--input-json`（结构化数据）

输出固定 5 个 sheet（顺序固定）：
1. `Summary`
2. `Benchmark`
3. `Feature-Matrix`
4. `Pricing-GTM`
5. `Sources`

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

## Quality & Trust / 质量与可信度
- 每个竞品建议至少 3 条来源
- 建议同时包含官方来源与第三方来源
- 语言本地化：`en/zh/ja/ko/es/fr/de`
- 回归测试：`python3 -m unittest discover -s tests -p 'test_*.py'`
- CI：GitHub Actions 自动验证

## Use Cases / 适用场景
- 新产品立项竞品扫描
- 路线图优先级对位
- 投资/评审前的竞争格局梳理

## Project Layout
- `SKILL.md`: workflow and methodology contract
- `agents/openai.yaml`: UI metadata
- `scripts/`: benchmark generator
- `references/`: methodology, rubric, source-quality, i18n
- `examples/`: runnable quickstart
- `tests/`: regression tests

## License
MIT
