---
name: codex-lb-quota
description: Check a codex-lb API key's remaining quota, usage, limits, reset times, and shared upstream organization quota. Use when the user asks how much quota is left, how much of a codex-lb key limit is used, whether a sk-clb key still has balance, or wants a deterministic quota report for a codex-lb server.
---

# Codex LB Quota

## Workflow

Use the bundled Python script for the whole task. Do not manually reimplement quota parsing unless the script itself needs to be fixed.

```bash
python3 <skill-folder>/scripts/check_quota.py
```

The script:
- reads and writes `quota_config.json` next to this skill;
- stores the codex-lb API key locally after the first successful token entry;
- calls `GET /v1/usage` on the configured codex-lb server;
- prints deterministic quota output with percentages and ASCII progress bars;
- exits with a clear status code and message for missing or invalid tokens.

## Token Handling

On the first run, if the script reports `MISSING_TOKEN`, ask the user for the `sk-clb-...` token, then store it with:

```bash
python3 <skill-folder>/scripts/check_quota.py --set-token '<TOKEN>'
```

If the script reports `INVALID_TOKEN`, tell the user the saved token was rejected and offer to overwrite it. When the user provides a new token, run the same `--set-token` command.

Never print the full saved token back to the user. The script masks tokens in normal output.

## Server

The default server is `https://gptproxy.dixel.store`.

To use another codex-lb instance for a run:

```bash
python3 <skill-folder>/scripts/check_quota.py --base-url 'https://example.com'
```

To save another server as the default:

```bash
python3 <skill-folder>/scripts/check_quota.py --base-url 'https://example.com' --save-base-url
```

## Reporting Rules

Summarize the script output in Russian:
- lead with the API key quota result;
- for `cost_usd`, treat values as microdollars and display dollars;
- use `remaining_value` as the actual key balance;
- use `upstream_limits` only as the shared organization quota context, not as the user's personal key balance;
- mention reset dates that are already in the past;
- if the user exposed a token in chat, remind them to rotate it after checking.
