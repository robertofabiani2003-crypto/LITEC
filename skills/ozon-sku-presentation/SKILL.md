---
name: ozon-sku-presentation
description: Generate a fixed-format Ozon SKU PowerPoint deck from a merged Excel workbook and a PPTX template. Use when the task matches the local Ozon SKU reporting workflow and should be run explicitly via $ozon-sku-presentation.
---

# Ozon SKU Presentation

Use this skill for the local Ozon SKU deck generator in this skill folder. This is not a general presentation skill. If the user asks for a generic deck or a new visual concept, use `PowerPoint` instead.

## What This Skill Does

- Builds a `.pptx` deck from a merged Ozon Excel workbook.
- Uses a local PPTX template only to inherit slide size.
- Generates one slide per SKU from the built-in SKU list.
- Preserves the current reporting logic from `scripts/generate_ozon_sku_presentation.py`.

## Default Workflow

1. Resolve three paths:
   - source workbook `.xlsx`
   - template `.pptx`
   - output `.pptx`
2. Run:

```powershell
python <skill-folder>\scripts\generate_ozon_sku_presentation.py --input-xlsx "<workbook.xlsx>" --template-pptx "<template.pptx>" --output-pptx "<output.pptx>"
```

3. Verify the script prints `Saved:` and that the output file exists.
4. Return the saved `.pptx` path.

## When To Patch The Script

Patch `scripts/generate_ozon_sku_presentation.py` only when the user explicitly needs a different built-in setup:

- another `TARGET_SKUS` list;
- another `MANUAL_STOCK_LINES` block;
- changed metric mapping or table layout.

Do not rewrite the analytical logic unless the input workbook structure changed or the user asked for a different report shape.

## Notes

- The script accepts CLI args for workbook, template, and output paths.
- The current SKU set and manual stock overrides remain embedded in the script.
- Prefer reusing the existing generator instead of rebuilding the deck logic from scratch.
- Replace `<skill-folder>` with the local installed path to this skill.
