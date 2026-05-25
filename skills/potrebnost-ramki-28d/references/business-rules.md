# Business Rules

## Scope

Use these rules when building or debugging the aggregated `рамки` sheet for Ozon FBS replenishment.

## Frame Identity

- `K_T_M1X283` and `K_T_A13X283` map to frame `X283`.
- `K_T_M1X133-1`, `K_T_M1X133-11`, `K_T_A13X133-3` map to frame `X133`.
- Build the frame code as `X` plus the suffix after the family prefix, then cut everything after the first `-`.

## Family Sheets

- Keep source rows by original article.
- Show real per-article sales and stock there.
- Do not overwrite family-sheet rows with frame totals unless the user explicitly asks for that variant.

## Frame Sheet

- Aggregate sales across all related articles that belong to the same frame.
- Keep warehouses separate.
- Default warehouse columns are `Уфа ФБС` and `МСК Наш Склад`.

## Stock Rule For A Frame

- Do not sum stocks from all related articles.
- For each warehouse, take the stock from one article with the maximum stock among all related articles for that frame.
- The chosen article may belong to either family, for example `M1X` or `A13X`.

## Replenishment Rule For A Frame

Calculate replenishment separately for each warehouse.

1. Sum sales for the last 28 days only for articles of this frame that had sales on that warehouse.
2. Detect whether the frame has at least one other related article on that warehouse with zero sales in the last 28 days.
3. If yes, add one shared buffer of `3` units for the non-selling part of the frame.
4. Subtract the chosen maximum single-article stock for that warehouse.
5. Clamp the result at zero.

Formula:

`recommendation = max(sum_sales_28_selling_articles + optional_buffer_3 - max_single_article_stock, 0)`

Where:

- `optional_buffer_3` is `3` if at least one related article had zero sales in the last 28 days on that warehouse.
- otherwise `optional_buffer_3` is `0`.

## Coloring

- Red: `sales_28 > 0` and `stock == 0`
- Yellow: `sales_365 > 0` and `stock == 0`

Apply the fill to the 4-column block for that warehouse:

- 28-day sales
- 365-day sales
- stock
- recommendation

## Output Expectations

- Save workbook as `.xlsx`, not CSV.
- Preserve numeric cell types as numbers.
- Freeze the header row.
- Add auto-filter.
- If the target workbook is locked by Excel, save to a new filename and report it explicitly.
