# Ozon SKU Presentation Skill

This repository contains a shareable Codex skill for generating a fixed-format Ozon SKU PowerPoint deck from:

- a merged Excel workbook (`.xlsx`)
- a PowerPoint template (`.pptx`)

## Contents

- `SKILL.md` - skill entrypoint and usage instructions
- `agents/openai.yaml` - Codex UI metadata
- `scripts/generate_ozon_sku_presentation.py` - deck generator

## Install

Install this skill from GitHub into your local Codex skills directory.

Example:

```powershell
python <path-to-install-script> --repo robertofabiani2003-crypto/LITEC --path .
```

Or copy this folder into:

```text
~/.codex/skills/ozon-sku-presentation
```

## Python dependencies

The generator script expects these packages:

```powershell
pip install openpyxl python-pptx
```

## Usage

Run the skill explicitly with:

```text
$ozon-sku-presentation
```

Then provide:

- input workbook path
- template presentation path
- output presentation path
