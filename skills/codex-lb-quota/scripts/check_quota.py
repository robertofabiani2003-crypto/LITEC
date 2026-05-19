#!/usr/bin/env python3
from __future__ import annotations

import argparse
import getpass
import json
import os
import stat
import sys
import urllib.error
import urllib.request
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path
from typing import Any


DEFAULT_BASE_URL = "https://gptproxy.dixel.store"
SKILL_DIR = Path(__file__).resolve().parents[1]
CONFIG_PATH = SKILL_DIR / "quota_config.json"


class QuotaError(Exception):
    pass


class InvalidTokenError(QuotaError):
    pass


@dataclass(frozen=True)
class Config:
    token: str | None
    base_url: str


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Check codex-lb API key quota via /v1/usage.",
    )
    parser.add_argument("--base-url", help=f"codex-lb base URL. Default: {DEFAULT_BASE_URL}")
    parser.add_argument("--save-base-url", action="store_true", help="Save --base-url into quota_config.json.")
    parser.add_argument("--set-token", help="Store a new sk-clb token before checking quota.")
    parser.add_argument("--reset-token", action="store_true", help="Remove the saved token and prompt again when possible.")
    parser.add_argument("--json", action="store_true", help="Print enriched JSON instead of the human report.")
    parser.add_argument("--timeout", type=float, default=15.0, help="HTTP timeout in seconds.")
    return parser.parse_args()


def load_config() -> Config:
    if not CONFIG_PATH.exists():
        return Config(token=None, base_url=DEFAULT_BASE_URL)
    try:
        data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return Config(token=None, base_url=DEFAULT_BASE_URL)
    token = data.get("token")
    base_url = data.get("base_url") or DEFAULT_BASE_URL
    return Config(
        token=token if isinstance(token, str) and token.strip() else None,
        base_url=str(base_url).strip() or DEFAULT_BASE_URL,
    )


def save_config(config: Config) -> None:
    payload = {
        "base_url": normalize_base_url(config.base_url),
        "token": config.token,
        "updated_at": datetime.now(UTC).isoformat().replace("+00:00", "Z"),
    }
    CONFIG_PATH.write_text(json.dumps(payload, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    try:
        CONFIG_PATH.chmod(stat.S_IRUSR | stat.S_IWUSR)
    except OSError:
        pass


def normalize_base_url(value: str) -> str:
    value = value.strip().rstrip("/")
    if not value:
        return DEFAULT_BASE_URL
    if not value.startswith(("http://", "https://")):
        value = "https://" + value
    return value.rstrip("/")


def mask_token(token: str) -> str:
    if len(token) <= 16:
        return token[:4] + "..."
    return token[:10] + "..." + token[-6:]


def prompt_token() -> str | None:
    if not sys.stdin.isatty():
        return None
    token = getpass.getpass("Введите codex-lb API key (sk-clb-...): ").strip()
    return token or None


def get_token_for_run(config: Config, explicit_token: str | None) -> tuple[str, bool]:
    if explicit_token:
        return explicit_token, True
    if config.token:
        return config.token, False
    token = prompt_token()
    if token:
        return token, True
    print(
        "MISSING_TOKEN: сохраненный токен не найден. "
        "Запустите с --set-token '<sk-clb-...>' или попросите пользователя дать токен для сохранения.",
        file=sys.stderr,
    )
    raise SystemExit(2)


def fetch_usage(base_url: str, token: str, timeout: float) -> dict[str, Any]:
    url = normalize_base_url(base_url) + "/v1/usage"
    request = urllib.request.Request(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
            "User-Agent": "codex-lb-quota-skill/1.0",
        },
        method="GET",
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            raw = response.read().decode("utf-8")
    except urllib.error.HTTPError as exc:
        body = exc.read().decode("utf-8", errors="replace")
        if exc.code in (401, 403):
            raise InvalidTokenError(body or f"HTTP {exc.code}") from exc
        raise QuotaError(f"HTTP {exc.code}: {body}") from exc
    except urllib.error.URLError as exc:
        raise QuotaError(f"NETWORK_ERROR: {exc.reason}") from exc

    try:
        payload = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise QuotaError(f"INVALID_JSON: {raw[:500]}") from exc

    if isinstance(payload, dict):
        error = payload.get("error")
        if isinstance(error, dict) and error.get("code") == "invalid_api_key":
            raise InvalidTokenError(str(error.get("message") or "Invalid API key"))
        return payload
    raise QuotaError("INVALID_RESPONSE: expected JSON object")


def maybe_replace_invalid_token(config: Config) -> str | None:
    if not sys.stdin.isatty():
        return None
    answer = input("Сохраненный токен не принят сервером. Перезаписать? [y/N]: ").strip().lower()
    if answer not in {"y", "yes", "д", "да"}:
        return None
    token = prompt_token()
    if token:
        save_config(Config(token=token, base_url=config.base_url))
    return token


def parse_dt(value: Any) -> datetime | None:
    if not isinstance(value, str) or not value:
        return None
    normalized = value.replace("Z", "+00:00")
    try:
        dt = datetime.fromisoformat(normalized)
    except ValueError:
        return None
    if dt.tzinfo is None:
        return dt.replace(tzinfo=UTC)
    return dt.astimezone(UTC)


def fmt_dt(value: Any) -> str:
    dt = parse_dt(value)
    if dt is None:
        return str(value or "unknown")
    return dt.strftime("%Y-%m-%d %H:%M:%S UTC")


def is_past(value: Any) -> bool:
    dt = parse_dt(value)
    return bool(dt and dt < datetime.now(UTC))


def as_number(value: Any) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def used_percent(limit: dict[str, Any]) -> float:
    max_value = as_number(limit.get("max_value"))
    current = as_number(limit.get("current_value"))
    if max_value <= 0:
        return 0.0
    return max(0.0, min(100.0, current / max_value * 100.0))


def remaining_percent(limit: dict[str, Any]) -> float:
    max_value = as_number(limit.get("max_value"))
    remaining = as_number(limit.get("remaining_value"))
    if max_value <= 0:
        return 0.0
    return max(0.0, min(100.0, remaining / max_value * 100.0))


def sort_limits(limits: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return sorted(
        limits,
        key=lambda item: (
            str(item.get("limit_type") or "unknown"),
            str(item.get("limit_window") or "unknown"),
            str(item.get("model_filter") or "all models"),
        ),
    )


def dollars(microdollars: Any) -> str:
    return f"${as_number(microdollars) / 1_000_000:.2f}"


def format_value(limit_type: str, value: Any) -> str:
    if limit_type == "cost_usd":
        return dollars(value)
    numeric = as_number(value)
    if numeric.is_integer():
        return f"{int(numeric):,}".replace(",", " ")
    return f"{numeric:,.2f}".replace(",", " ")


def tightest_limit(limits: list[dict[str, Any]]) -> dict[str, Any] | None:
    if not limits:
        return None
    return min(limits, key=remaining_percent)


def limit_label(limit: dict[str, Any]) -> str:
    limit_type = str(limit.get("limit_type") or "unknown")
    window = str(limit.get("limit_window") or "unknown")
    model = limit.get("model_filter") or "all models"
    return f"{limit_type} / {window} / {model}"


def ascii_bar(percent: float, width: int = 10) -> str:
    clamped = max(0.0, min(100.0, percent))
    filled = max(0, min(width, int(clamped / 100 * width + 0.5)))
    return "[" + "#" * filled + "-" * (width - filled) + "]"


def progress_line(label: str, value: str, maximum: str, percent: float, *, indent: str = "") -> str:
    return f"{indent}{label:<12}: {value} из {maximum} {ascii_bar(percent)} {percent:.2f}%"


def enriched_payload(payload: dict[str, Any], base_url: str, token: str) -> dict[str, Any]:
    limits = sort_limits([item for item in payload.get("limits", []) if isinstance(item, dict)])
    upstream = sort_limits([item for item in payload.get("upstream_limits", []) if isinstance(item, dict)])
    tightest = tightest_limit(limits)
    return {
        "base_url": normalize_base_url(base_url),
        "token_masked": mask_token(token),
        "request_count": payload.get("request_count", 0),
        "total_tokens": payload.get("total_tokens", 0),
        "cached_input_tokens": payload.get("cached_input_tokens", 0),
        "total_cost_usd": payload.get("total_cost_usd", 0.0),
        "limits": limits,
        "upstream_limits": upstream,
        "tightest_limit": tightest,
        "tightest_remaining_percent": remaining_percent(tightest) if tightest else None,
        "tightest_used_percent": used_percent(tightest) if tightest else None,
    }


def print_limit(limit: dict[str, Any], *, prefix: str = "-") -> None:
    limit_type = str(limit.get("limit_type") or "unknown")
    current = format_value(limit_type, limit.get("current_value"))
    maximum = format_value(limit_type, limit.get("max_value"))
    remaining = format_value(limit_type, limit.get("remaining_value"))
    reset_at = fmt_dt(limit.get("reset_at"))
    stale = "  ⚠ дата сброса уже в прошлом" if is_past(limit.get("reset_at")) else ""
    print(f"{prefix} {limit_label(limit)}")
    print(progress_line("Осталось", remaining, maximum, remaining_percent(limit), indent="  "))
    print(progress_line("Использовано", current, maximum, used_percent(limit), indent="  "))
    print(f"  Сброс       : {reset_at}{stale}")


def print_report(payload: dict[str, Any], base_url: str, token: str) -> None:
    data = enriched_payload(payload, base_url, token)
    limits = data["limits"]
    upstream = data["upstream_limits"]
    tightest = data["tightest_limit"]

    print(f"Codex LB quota: {data['base_url']}")
    print(f"Token          : {data['token_masked']}")
    print()
    print("[Квота ключа]")
    print(f"Запросы        : {data['request_count']}")
    print(f"Токены         : {data['total_tokens']} всего / {data['cached_input_tokens']} кэшированных")
    print(f"Учтённый расход: ${as_number(data['total_cost_usd']):.6f}")

    if tightest:
        limit_type = str(tightest.get("limit_type") or "unknown")
        remaining = format_value(limit_type, tightest.get("remaining_value"))
        current = format_value(limit_type, tightest.get("current_value"))
        maximum = format_value(limit_type, tightest.get("max_value"))
        reset_at = fmt_dt(tightest.get("reset_at"))
        stale = "  ⚠ дата сброса уже в прошлом" if is_past(tightest.get("reset_at")) else ""
        print(f"Главный лимит  : {limit_label(tightest)}")
        print(progress_line("Осталось", remaining, maximum, remaining_percent(tightest)))
        print(progress_line("Использовано", current, maximum, used_percent(tightest)))
        print(f"Сброс          : {reset_at}{stale}")
    else:
        print("Главный лимит  : лимиты на ключе не настроены")

    print()
    print("Все лимиты ключа:")
    if limits:
        for limit in limits:
            print_limit(limit)
    else:
        print("- лимиты на ключе не настроены")

    print()
    print("[Общая квота организации]")
    if upstream:
        for limit in upstream:
            print_limit(limit)
    else:
        print("- сервер не вернул общие лимиты организации")


def main() -> int:
    args = parse_args()
    config = load_config()

    if args.reset_token:
        config = Config(token=None, base_url=config.base_url)
        save_config(config)

    base_url = normalize_base_url(args.base_url or os.environ.get("CODEX_LB_QUOTA_BASE_URL") or config.base_url)
    explicit_token = args.set_token.strip() if isinstance(args.set_token, str) and args.set_token.strip() else None
    saved_base_url = base_url if args.save_base_url or explicit_token else config.base_url

    token, should_save_token = get_token_for_run(config, explicit_token)

    try:
        payload = fetch_usage(base_url, token, args.timeout)
    except InvalidTokenError as exc:
        replacement = maybe_replace_invalid_token(Config(token=None, base_url=base_url))
        if replacement:
            try:
                payload = fetch_usage(base_url, replacement, args.timeout)
                token = replacement
                should_save_token = True
            except InvalidTokenError:
                print("INVALID_TOKEN: новый токен тоже отклонен сервером.", file=sys.stderr)
                return 3
        else:
            print(f"INVALID_TOKEN: сохраненный токен отклонен сервером. {exc}", file=sys.stderr)
            print("Запустите с --set-token '<new-sk-clb-token>', чтобы перезаписать токен.", file=sys.stderr)
            return 3
    except QuotaError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    if should_save_token or args.save_base_url:
        save_config(Config(token=token, base_url=saved_base_url))

    if args.json:
        print(json.dumps(enriched_payload(payload, base_url, token), indent=2, ensure_ascii=False))
    else:
        print_report(payload, base_url, token)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
