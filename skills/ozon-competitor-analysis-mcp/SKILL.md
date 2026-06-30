---
name: ozon-competitor-analysis-mcp
description: Analyze Ozon competitor-report rows through the LITEC OS MCP server, refresh the short-lived MCP token when needed, compare our SKU with the top competitor, include week-over-week metrics for both sides, and write concise Russian comments back to the report. Use when the user mentions $ozon-competitor-analysis-mcp, asks to analyze Ozon competitor data from MCP, or wants SKU comments written into the MCP report rather than an Excel workbook.
---

# Ozon Competitor Analysis MCP

Use this skill for MCP-backed Ozon competitor analysis in LITEC OS.

This skill is separate from the Excel workbook flow. Here the source of truth is the MCP report, not `.xlsx` files.

## Workflow

1. Ensure the MCP server is configured in `~/.codex/config.toml`:

```toml
[mcp_servers.ozon_analytics]
url = "https://admin.dixel.store/v1/integrations/ozon-analytics/mcp"
bearer_token_env_var = "OZON_ANALYTICS_MCP_TOKEN"
```

2. If `OZON_ANALYTICS_MCP_TOKEN` is missing or MCP returns `401`, refresh the token:
- open `https://admin.dixel.store/v1/auth/mcp-token` in the in-app browser
- if the page asks for login, ask the user to sign in to the LITEC OS admin, then reopen the URL
- read `accessToken` from `textarea#mcp-token` or `script#mcp-token-payload`
- set it into `OZON_ANALYTICS_MCP_TOKEN` for the current Codex process
- prefer also updating the user environment variable so the token survives the next shell command in the same desktop session
- never paste the token into chat

3. Verify that MCP is alive and only use tools returned by `tools/list`.

4. Read report data only through `ozon_analytics_report_read`.

5. Write comments only through `ozon_analytics_comment_write` with:
- `sku`
- `weekStart`
- `comment`

## Hard Rules

- Never read the report through browser scraping when MCP is available.
- Never write comments through any tool except `ozon_analytics_comment_write`.
- Never fabricate previous-week data. If history is absent for our SKU or the competitor, say that directly.
- Treat competitor analysis as part of our SKU analysis, not as a detached note.
- Do not reduce the conclusion to a generic "weak card" claim if the real issue is traffic, share, or visibility.

## Data To Pull

For any target SKU, collect:
- current-week row
- previous-week row when available

If the target SKU belongs to us, also collect:
- current-week row with the highest `orderedUnits`
- previous-week row for that competitor when available

Use these MCP fields when available:
- `totalImpressions`
- `searchCatalogViews`
- `picSharePercent`
- `cardViews`
- `ctrPercent`
- `carts`
- `searchCatalogToCartPercent`
- `cardToCartPercent`
- `orderedUnits`
- `cartToOrderConversionPercent`
- `orderedAmountRub`
- `adExpenseSharePercent`
- `outOfStockDays`

## Analysis Logic

First determine the mode:
- our SKU mode: the row seller is our seller, so compare our SKU against the top competitor in the same slice and also do week-over-week
- competitor self mode: the row seller is not our seller, so do not compare against our SKU or another competitor; analyze only the SKU against its own previous week

In our SKU mode, always evaluate both:
- relative metrics: share, CTR, search-to-cart, card-to-cart, cart-to-order, DRR
- absolute volume: impressions, search traffic, card views, carts, orders, revenue

Do not treat a strong lower funnel as enough if traffic collapsed.
Do not treat a healthy CTR as enough if absolute orders still lag.

The written analysis for our SKU mode must cover:
1. Our SKU versus the top competitor on the current week.
2. Our SKU week-over-week.
3. The top competitor week-over-week.
4. A short synthesis: risks, hypothesis, focus.

The written analysis for competitor self mode must cover:
1. Current-week result for that SKU with key funnel numbers.
2. SKU week-over-week versus its own previous week.
3. A short synthesis: risks, hypothesis, focus.
4. No comparison versus our SKU and no substitute competitor comparison.

## Comment Structure

Default to a compact Russian block with numbers first.

Use this shape for our SKU mode:

```text
Отстаем от <competitor_sku> по заказам: <our_orders> vs <comp_orders>.
При этом у нас больше/меньше охват:
показы <our> vs <comp>,
поиск <our> vs <comp>,
карточка <our> vs <comp>,
корзины <our> vs <comp>.
Главный разрыв в эффективности трафика: CTR <our> vs <comp>, в корзину из поиска <our> vs <comp>, в корзину из карточки <our> vs <comp>.

К прошлой неделе у нас:
показы ...,
поиск ...,
карточка ...,
корзины ...,
заказы ...,
выручка ....

У конкурента неделя к неделе:
показы ...,
поиск ...,
карточка ...,
корзины ...,
заказы ...,
выручка ....

Риски: ...
Гипотеза: ...
Фокус: ...
```

If the data show that our lower funnel is stronger than the competitor's, say that explicitly.

Use this shape for competitor self mode:

```text
За неделю <week_start> SKU <sku> сделал <orders> заказов на <revenue>.
Показы <impressions>, поиск <search>, карточка <card_views>, корзины <carts>.
CTR <ctr>, в корзину из поиска <search_to_cart>, в корзину из карточки <card_to_cart>, в заказ из корзины <cart_to_order>, ДРР <drr>.

К прошлой неделе:
показы ...,
поиск ...,
карточка ...,
корзины ...,
заказы ...,
выручка ....

Риски: ...
Гипотеза: ...
Фокус: ...
```

If previous-week data are absent for a competitor SKU, say that directly and base the synthesis only on the current week.

Example:
- "Главный разрыв в верхней части воронки."
- "После входа SKU держится не хуже конкурента."
- "Конкурент тоже снизился неделя к неделе, но заметно мягче."

## Metric Priorities

When deciding what matters most:
- first diagnose the top of funnel: `totalImpressions`, `searchCatalogViews`, `picSharePercent`, `ctrPercent`
- then diagnose card and cart efficiency: `cardViews`, `searchCatalogToCartPercent`, `cardToCartPercent`
- then diagnose order completion: `carts`, `orderedUnits`, `cartToOrderConversionPercent`
- then check monetization and efficiency: `orderedAmountRub`, `adExpenseSharePercent`
- always mention `outOfStockDays` when it is non-zero or materially worse than the comparison row

In competitor self mode, interpret "comparison row" as that same SKU on the previous week.

## Token Refresh Checklist

If MCP fails:
1. Assume the token may have expired before assuming the server is broken.
2. Decode or inspect the JWT only if useful, but trust `401` as the primary signal.
3. Refresh the token from `/v1/auth/mcp-token`.
4. Re-check `tools/list`.
5. Continue only after `ozon_analytics_report_read` and `ozon_analytics_comment_write` are visible.

## Deliverable

When the user asks for analysis:
- write the comment to MCP immediately by default
- return the written comment text in chat after saving as a confirmation
- only skip writing and show a draft first when the user explicitly asks to review before saving

When the user asks to update many SKUs:
- keep one consistent comment style across the batch
- prefer concise, management-readable blocks over long prose
- write each SKU comment through `ozon_analytics_comment_write` in the same run unless the user explicitly asked for a dry run or review first
- automatically choose the correct mode per row:
  - our seller rows -> compare with top competitor + week-over-week
  - non-our seller rows -> self week-over-week only
- after batch writing, re-read the report and verify that comments were actually persisted for the updated rows
