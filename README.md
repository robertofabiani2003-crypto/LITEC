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
  ozon-competitor-analysis-mcp/
    SKILL.md
    README.md
    agents/openai.yaml
  potrebnost-ramki-28d/
    SKILL.md
    agents/openai.yaml
    references/business-rules.md
    scripts/export_family_fbs.py
    scripts/build_frame_workbook.py
  ozon-seller-complaints/
    SKILL.md
    agents/openai.yaml
    references/config.example.json
    references/configuration.md
    scripts/open_cdp_browser.ps1
    scripts/run_ozon_complaints.mjs
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

### `potrebnost-ramki-28d`

Builds Ozon FBS family exports and an Excel workbook with:

- per-article family sheets
- an aggregated `рамки` sheet
- 28-day frame demand logic
- frame stock taken from the single article with the highest stock on each warehouse

Skill path in this repository:

```text
skills/potrebnost-ramki-28d
```

### `ozon-competitor-analysis-mcp`

Analyzes Ozon competitor-report rows through the LITEC OS MCP server and writes concise Russian comments back into the MCP report.

Includes:

- MCP token refresh workflow
- current-week vs previous-week analysis
- our SKU vs top competitor comparison when applicable
- direct comment write-back through MCP

Skill path in this repository:

```text
skills/ozon-competitor-analysis-mcp
```

### `ozon-seller-complaints`

Automates Ozon Seller support complaints against target SKUs using:

- a logged-in local browser session over CDP
- a seller-owned SKU
- evidence files
- a plain text SKU list or Excel source
- resumable progress and log files

Skill path in this repository:

```text
skills/ozon-seller-complaints
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

or:

```text
Install the Codex skill from https://github.com/robertofabiani2003-crypto/LITEC using the path skills/potrebnost-ramki-28d
```

or:

```text
Install the Codex skill from https://github.com/robertofabiani2003-crypto/LITEC using the path skills/ozon-competitor-analysis-mcp
```

or:

```text
Install the Codex skill from https://github.com/robertofabiani2003-crypto/LITEC using the path skills/ozon-seller-complaints
```

After installation, restart Codex so it picks up the new skill.

## Python dependencies

The `ozon-sku-presentation` generator script expects:

```powershell
pip install openpyxl python-pptx
```

The `potrebnost-ramki-28d` scripts expect:

```powershell
pip install openpyxl requests
```
