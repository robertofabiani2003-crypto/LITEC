# ozon-competitor-analysis-mcp

Codex skill for MCP-based Ozon competitor analysis in LITEC OS.

The skill is designed for workflows where the source of truth is the LITEC MCP report rather than an Excel workbook. It reads report rows through MCP, compares our SKU with the leading competitor when applicable, and writes concise Russian comments back into the report.

## Includes

- [SKILL.md](./SKILL.md): Codex skill entrypoint
- [agents/openai.yaml](./agents/openai.yaml): UI metadata for Codex

## MCP setup

Add this server to `~/.codex/config.toml`:

```toml
[mcp_servers.ozon_analytics]
url = "https://admin.dixel.store/v1/integrations/ozon-analytics/mcp"
bearer_token_env_var = "OZON_ANALYTICS_MCP_TOKEN"
```

## Token handling

If the MCP token is missing or expired:

1. Open `https://admin.dixel.store/v1/auth/mcp-token` in the Codex in-app browser.
2. Sign in to the LITEC OS admin if prompted.
3. Read `accessToken` from `textarea#mcp-token` or `script#mcp-token-payload`.
4. Set `OZON_ANALYTICS_MCP_TOKEN` locally for the current Codex session.
5. Prefer also updating the user environment variable so the token survives the next shell command in the same desktop session.

Never paste the token into chat or commit it into this repository.

## Usage

Use the skill in Codex as:

```text
Use $ozon-competitor-analysis-mcp to analyze Ozon MCP report rows and write back concise Russian comments.
```

## Notes

- Read report data only through `ozon_analytics_report_read`.
- Write comments only through `ozon_analytics_comment_write`.
- If previous-week data are missing, the analysis should say that directly instead of inferring history.
