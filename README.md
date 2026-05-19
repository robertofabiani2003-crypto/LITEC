# LITEC Codex Skills

This repository is a collection of shareable Codex skills.

## Structure

Each skill lives in its own folder under `skills/`.

```text
skills/
  ozon-sku-presentation/
    SKILL.md
    agents/openai.yaml
    scripts/generate_ozon_sku_presentation.py
```

## Available skills

### `ozon-sku-presentation`

Generates a fixed-format Ozon SKU PowerPoint deck from:

- a merged Excel workbook (`.xlsx`)
- a PowerPoint template (`.pptx`)

Skill path in this repository:

```text
skills/ozon-sku-presentation
```

## Install

Ask Codex to install the skill from this repository and specify the skill folder path:

```text
Install the Codex skill from https://github.com/robertofabiani2003-crypto/LITEC using the path skills/ozon-sku-presentation
```

After installation, restart Codex so it picks up the new skill.

## Python dependencies

The `ozon-sku-presentation` generator script expects:

```powershell
pip install openpyxl python-pptx
```
