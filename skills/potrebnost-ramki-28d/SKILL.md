---
name: potrebnost-ramki-28d
description: Build Ozon FBS family exports and frame-level replenishment workbooks with 28-day frame demand logic for cabinets with multiple Client-Id values. Use when Codex needs to выгрузить продажи FBS и остатки по семействам артикулов like K_T_M1X or K_T_A13X, explain or change frame-based replenishment rules, merge related SKUs into frame codes like X133 on a рамки sheet, or produce Excel workbooks with family sheets plus an aggregated frame sheet.
---

# Ozon Frame Replenishment

Use this skill for Ozon Seller workbook workflows where several SKU families share one physical frame and replenishment must be reasoned about at both article and frame level.

## Workflow

1. Identify the cabinet and family prefixes.
   Typical examples: `CABINET_A`, `K_T_M1X`, `K_T_A13X`.

2. Read the workspace `.env`.
   The bundled scripts expect the cabinet format already used in this project:
   - one line with the cabinet name
   - one nearby `OZON_API_KEY=...`
   - one nearby `OZON_CLIENT_ID=...`
   - if the file is missing or incomplete, point the user to `.env.example` in this skill and ask them to fill only the placeholder values

3. Export raw family data.
   Run `scripts/export_family_fbs.py` once per family prefix, using the correct cabinet for that family, to fetch:
   - matching Ozon products
   - current FBS stocks by warehouse
   - FBS posting rows
   - CSV summaries in `outputs/`

4. Build the workbook.
   Run `scripts/build_frame_workbook.py` with the family list and explicit `--family-cabinet FAMILY=CABINET` mappings when families come from different client IDs.
   The workbook structure is:
   - one sheet per family prefix with per-article metrics
   - one aggregated frame sheet, default name `рамки`

5. If the user asks to change formulas or audit a specific frame, read `references/business-rules.md` and then patch the workbook builder rather than improvising the logic in prose.

## Commands

Export one family:

```powershell
python .\scripts\export_family_fbs.py CABINET_A K_T_M1X --env .\.env --output-root .\outputs
```

Build the combined workbook:

```powershell
python .\scripts\build_frame_workbook.py CABINET_B --family K_T_M1X --family K_T_A13X --family-cabinet 'K_T_M1X=CABINET_A' --family-cabinet 'K_T_A13X=CABINET_B' --output-root .\outputs --workbook-path .\outputs\frame_replenishment.xlsx
```

## What To Keep Straight

- Keep family sheets per-article. Do not replace article rows with frame aggregates there unless the user explicitly asks for that.
- Keep the `рамки` sheet aggregated by frame code.
- Treat stock on the frame sheet as stock from one article with the maximum stock for that warehouse, not the sum of all related articles.
- Treat replenishment on the frame sheet as a separate rule from per-article stock reporting.
- If Excel locks the workbook, save to a new filename and tell the user which file was created.

## Resources

### `scripts/export_family_fbs.py`

Pull Ozon FBS products, postings, and stocks for one family prefix and write reusable CSV exports.

### `scripts/build_frame_workbook.py`

Build the Excel workbook from the latest family exports and apply the current frame-level replenishment logic.

### `references/business-rules.md`

Load this file when you need the exact frame aggregation and recommendation rules or when the user asks why a frame was calculated in a certain way.
