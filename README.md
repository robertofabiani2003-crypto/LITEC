# LITEC Codex Skills

This repository is a collection of shareable Codex skills.

## Structure

Each skill lives in its own folder under `skills/`.

```text
skills/
  codex-lb-quota/
    SKILL.md
    .gitignore
    agents/openai.yaml
    scripts/check_quota.py
  ozon-sku-presentation/
    SKILL.md
    agents/openai.yaml
    scripts/generate_ozon_sku_presentation.py
```

## Available skills

### `codex-lb-quota`

Checks a codex-lb API key quota and prints a deterministic report with:

- remaining key balance
- used percentage
- reset times
- organization-wide shared quota
- ASCII progress bars

Skill path in this repository:

```text
skills/codex-lb-quota
```

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
Install the Codex skill from https://github.com/robertofabiani2003-crypto/LITEC using the path skills/codex-lb-quota
```

or:

```text
Install the Codex skill from https://github.com/robertofabiani2003-crypto/LITEC using the path skills/ozon-sku-presentation
```

After installation, restart Codex so it picks up the new skill.

## Python dependencies

The `ozon-sku-presentation` generator script expects:

```powershell
pip install openpyxl python-pptx
```
