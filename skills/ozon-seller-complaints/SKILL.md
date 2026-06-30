---
name: ozon-seller-complaints
description: Automate Ozon Seller support complaints about copied content or other seller violations using the logged-in seller account in a local browser. Use when Codex needs to configure, launch, resume, or adapt a complaint workflow for a user's own Ozon Seller cabinet, SKU list, evidence files, CDP browser session, or complaint flow labels.
---

# Ozon Seller Complaints

## Overview

Use this skill to run or adapt an Ozon Seller complaint flow from a locally logged-in browser. The bundled runner is config-driven, so another user can point it at their own seller SKU, target SKU list, evidence files, browser profile, and runtime output files without editing the script itself.

## Workflow

1. Prepare a working folder outside the skill with:
   - a config JSON based on `references/config.example.json`
   - a plain text SKU list with one target SKU per line
   - optional evidence files to upload

2. Prefer a text SKU list over Excel.
   Read `references/configuration.md` only when the user needs Excel input, custom runtime paths, retries, or label overrides.

3. Start a browser session with remote debugging and the user's own browser profile.
   Use `scripts/open_cdp_browser.ps1` when the user wants help opening Chrome, Yandex Browser, or Edge on a CDP port.

4. Ensure the browser is already logged into the correct Ozon Seller cabinet before running the complaint script.

5. Run the bundled script:

```powershell
node "C:\Users\rober\.codex\skills\ozon-seller-complaints\scripts\run_ozon_complaints.mjs" --config "C:\path\to\config.json"
```

Optional flags:

```powershell
node "C:\Users\rober\.codex\skills\ozon-seller-complaints\scripts\run_ozon_complaints.mjs" `
  --config "C:\path\to\config.json" `
  --start-index=10 `
  --limit=5 `
  --cdp-url=http://127.0.0.1:9223
```

6. Read progress from the configured `progressPath` and `logPath`.
   Resume by running the same command again. Restart from scratch by replacing or deleting the progress and log files in the user's working folder.

## What To Customize

- Change `ownSku` to the seller's own article.
- Change `attachments` to the user's own proof files.
- Change `skuListPath` or `excelPath` to the user's target SKU source.
- Change `cdpUrl`, browser profile, or `startUrl` only when the local browser setup differs.
- Override the `flow` labels only when Ozon changes visible Russian UI text.

## Bundled Resources

- `scripts/run_ozon_complaints.mjs`: Main complaint runner. Reads JSON config, resumes from progress, and retries a SKU when final confirmation times out.
- `scripts/open_cdp_browser.ps1`: Opens a local browser profile on a CDP port and verifies that the endpoint responds.
- `references/config.example.json`: Starting template for per-user configuration.
- `references/configuration.md`: Field-by-field configuration reference.
