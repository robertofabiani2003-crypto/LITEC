# potrebnost-ramki-28d

Codex skill for Ozon FBS family exports and frame-level replenishment workbooks.

The skill is designed for cases where several related SKUs share one physical frame and replenishment must be calculated at frame level while family sheets stay per article.

## What It Does

- Exports Ozon FBS postings, products, and current stocks for article families such as `K_T_M1X*` and `K_T_A13X*`
- Builds Excel workbooks with:
  - one sheet per family prefix
  - one aggregated frame sheet, usually `рамки`
- Explains and enforces frame-level replenishment logic
- Works with multiple Ozon cabinets stored in the same `.env`

## Repository Contents

- [SKILL.md](./SKILL.md): Codex skill entrypoint
- [agents/openai.yaml](./agents/openai.yaml): UI metadata for Codex
- [scripts/export_family_fbs.py](./scripts/export_family_fbs.py): raw Ozon export for one family prefix
- [scripts/build_frame_workbook.py](./scripts/build_frame_workbook.py): final workbook builder
- [scripts/ozon_frame_common.py](./scripts/ozon_frame_common.py): shared helpers
- [references/business-rules.md](./references/business-rules.md): exact replenishment and aggregation rules

## Skill Trigger

Use the skill in Codex as:

```text
$potrebnost-ramki-28d
```

Typical requests:

- "Выгрузи продажи и остатки по K_T_M1X и K_T_A13X"
- "Собери файл по рамкам для кабинета CABINET_A"
- "Разбери, как посчиталась потребность по X133"
- "Пересобери workbook с новой логикой рекомендаций по рамкам"

## Input Format

The scripts expect an `.env` in this format:

```env
CABINET_A
OZON_API_KEY=...
OZON_CLIENT_ID=1234567

CABINET_B
OZON_API_KEY=...
OZON_CLIENT_ID=2345678
```

For a safe handoff, give another person `.env.example` from this skill and ask them to copy it to `.env` and fill in only their own real values.

## Export One Family

```powershell
python .\scripts\export_family_fbs.py CABINET_A K_T_M1X --env .\.env --output-root .\outputs
python .\scripts\export_family_fbs.py CABINET_B K_T_A13X --env .\.env --output-root .\outputs
```

The export creates a timestamped folder with:

- `sales_detailed.csv`
- `sales_by_warehouse.csv`
- `sales_by_offer_and_warehouse.csv`
- `current_fbs_stocks.csv`
- `replenishment_recommendations.csv`
- `summary.json`

## Build Workbook

```powershell
python .\scripts\build_frame_workbook.py CABINET_B --family K_T_M1X --family K_T_A13X --family-cabinet 'K_T_M1X=CABINET_A' --family-cabinet 'K_T_A13X=CABINET_B' --output-root .\outputs --workbook-path .\outputs\frame_replenishment.xlsx
```

Result:

- family sheets keep original article rows
- frame sheet aggregates related articles into frame codes like `X133`
- each family can come from a different Ozon cabinet

## Current Business Logic

On family sheets:

- keep per-article sales
- keep per-article stock
- keep per-article article-level recommendation

On the `рамки` sheet:

- aggregate sales across related articles
- do not sum stock across all related articles
- use stock from the single article with the highest stock on that warehouse
- calculate frame recommendation from `28-day sales`

Frame recommendation formula by warehouse:

```text
max(sum_sales_28_for_selling_articles + optional_buffer_3 - max_single_article_stock, 0)
```

Where:

- `optional_buffer_3 = 3` if at least one related article had zero sales in the last 28 days on that warehouse
- otherwise `optional_buffer_3 = 0`

## Coloring Rules

- red: sales for `28 days > 0` and stock is `0`
- yellow: sales for `365 days > 0` and stock is `0`

## GitHub Upload

The repository can be pushed after GitHub authentication is configured.

### Option 1: GitHub CLI

```powershell
gh auth login
gh repo create potrebnost-ramki-28d --public --source . --remote origin --push
```

### Option 2: Existing GitHub Repository

```powershell
git remote add origin https://github.com/<username>/potrebnost-ramki-28d.git
git add .
git commit -m "Add potrebnost-ramki-28d Codex skill"
git branch -M main
git push -u origin main
```

## Notes

- If Excel keeps the workbook open, the builder saves to a new filename instead of failing.
- If you change replenishment logic, update [references/business-rules.md](./references/business-rules.md) together with the workbook script.
