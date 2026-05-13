# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Automated **银行客户洞察报告 (Bank Client Insight Report)** generation system for 北京易诚互动网络技术股份有限公司 (易诚互动), a bank IT solution provider. Generates structured 12-chapter Word (.docx) reports with 79 tables covering bank client analysis, strategic insights, procurement data, and business opportunity identification.

## Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Generate a full report
python scripts/generate_report.py \
  --output-file "./output.docx" \
  --bank-name "宁波银行" \
  --period "2025年度" \
  --json-content "filled_content.json"

# Parse sample reports (for understanding DOCX structure)
python scripts/parse_sample_report.py
python scripts/parse_nbcb_report.py
python scripts/extract_tables_detail.py
```

## Project Structure

```
/
├── SKILL.md                          # Complete skill definition — the single source of truth
├── scripts/
│   ├── generate_report.py            # BankReportGenerator class — main .docx generator
│   ├── parse_sample_report.py        # DOCX structure parser (prototyping)
│   ├── parse_nbcb_report.py          # Ningbo Bank report parser (chapters 1-2)
│   └── extract_tables_detail.py      # Tables 1-8 detailed extraction
├── references/
│   ├── report_structure.md           # Definitive chapter/section/table specs (79 tables)
│   ├── banking_knowledge.md          # Domain knowledge reference
│   └── data_sources.md               # Data collection sources & quality standards
├── assets/templates/
│   └── chapter{1-12}_template.json   # Per-chapter JSON templates with fixed table structures
├── requirements.txt
└── CLAUDE.md
```

## Report Generation Workflow

1. **Data collection** — Gather data from public sources: bank annual reports, official websites, China Bidding platform, regulatory disclosures, financial media
2. **Template filling** — Populate chapter-specific JSON templates from `assets/templates/` with collected data (each template has strict structural requirements — fixed row counts, column headers, etc.)
3. **Document generation** — Run `scripts/generate_report.py` to convert filled JSON into a .docx file

## Key Architecture

- **generate_report.py**: `BankReportGenerator` class handles all document formatting (cover page, headings, tables with alternating row colors, cell backgrounds via XML). Input JSON format supports `chapters[]` with nested `subsections[]`, each containing `content[]` items of type `heading`/`paragraph`/`table`.
- **Templates (JSON)**: 12 chapter templates with fixed table structures. Chapter 2 (基础信息) has 21 tables, Chapter 3 (业务与战略) has ~18 tables, Chapter 8 (采购全景) has ~33 tables. Row counts and column headers are rigid — **do not alter the fixed structure**.
- **SKILL.md**: The complete operational guide. Contains step-by-step generation instructions, quality checklists for every chapter, directory structure rules (e.g., Chapter 3 has 5 fixed L2 headings, business sections under L2-(二) have 7 fixed L4 headings per block), and strict rules about data source attribution.

## Critical Rules for the Agent

- **Tables are strictly fixed** — Row count, column headers, and column order must match the template exactly. Never delete or add rows/columns to fixed tables.
- **所有数据必须标注来源** (all data must cite sources) — Format: `来源：[source name] [URL] [date]`. Estimation data must be explicitly labeled.
- **No fabricated data** — Unobtainable data must be marked as "未披露" (undisclosed), never fabricated.
- **Report is from 易诚互动's perspective** — Opportunity analysis (Chapter 9) must objectively assess 易诚互动's actual capabilities (strong in channel systems, weak in core banking/fintech platforms).
- **Chapter title placeholders**: `【银行名称】` must be replaced with the actual bank name in chapters 5-7 and 9.
- **Data priority**: Annual reports > official website > regulatory disclosures > Wind/financial terminals > broker reports > financial media.
