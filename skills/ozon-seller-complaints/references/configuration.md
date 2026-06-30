# Configuration Reference

Use `references/config.example.json` as the starting point.

## Required fields

- `ownSku`: Seller SKU that identifies the complainant product.
- `attachments`: Files sent as evidence when Ozon asks for proof.
- `skuListPath` or `excelPath`: Source of target SKUs to complain about.

## Input sources

- `skuListPath`: Plain text file with one SKU per line. Prefer this mode.
- `excelPath`: Excel workbook path.
- `pythonPath`: Python executable with `pandas` installed. Required for `excelPath`.
- `excelColumnIndex`: Zero-based column index used when reading Excel. Defaults to `0`.

## Runtime files

- `progressPath`: JSON file with `completed` and `failed` items.
- `logPath`: JSONL event log.
- `storageStatePath`: Saved browser storage state after the run.
- `runtimeDir`: Optional base directory for the runtime files above. Defaults to the config file directory.

## Browser settings

- `cdpUrl`: CDP endpoint. Defaults to `http://127.0.0.1:9223`.
- `startUrl`: Landing page opened before every run and retry.
- `closeBrowserOnFinish`: Close the whole browser at the end. Default `false`.
- `saveStorageStateOnFinish`: Save Playwright storage state. Default `true`.

## Retry settings

- `maxAttemptsPerSku`: How many times to retry the same SKU when the final hidden-product confirmation does not arrive.

## Flow overrides

Use the `flow` object only when Ozon changes visible Russian labels in the support UI.

- `helpEntry`
- `categoryProducts`
- `categoryQuality`
- `categoryViolation`
- `categoryContent`
- `actionComplain`
- `promptOwnSku`
- `promptTargetSku`
- `promptSingleValue`
- `successPrefix`
- `refresh`
- `loadError`
- `chatListError`
