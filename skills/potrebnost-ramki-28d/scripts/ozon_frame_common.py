from __future__ import annotations

import csv
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path
from typing import Iterable

import requests


BASE_URL = "https://api-seller.ozon.ru"


@dataclass
class CabinetCreds:
    client_id: str
    api_key: str


def load_cabinet_creds(env_path: Path, cabinet_name: str) -> CabinetCreds:
    lines = env_path.read_text(encoding="utf-8").splitlines()
    for index, raw in enumerate(lines):
        line = raw.strip()
        if line != cabinet_name:
            continue
        client_id = None
        api_key = None
        for sub_raw in lines[index + 1 : index + 4]:
            sub = sub_raw.strip()
            if sub.startswith("OZON_CLIENT_ID="):
                client_id = sub.split("=", 1)[1]
            elif sub.startswith("OZON_API_KEY="):
                api_key = sub.split("=", 1)[1]
        if client_id and api_key:
            return CabinetCreds(client_id=client_id, api_key=api_key)
    raise RuntimeError(f"Cabinet {cabinet_name!r} not found in {env_path}")


class OzonSellerClient:
    def __init__(self, creds: CabinetCreds) -> None:
        self.session = requests.Session()
        self.session.headers.update(
            {
                "Client-Id": creds.client_id,
                "Api-Key": creds.api_key,
                "Content-Type": "application/json",
            }
        )

    def post(self, path: str, payload: dict) -> dict:
        response = self.session.post(f"{BASE_URL}{path}", json=payload, timeout=120)
        response.raise_for_status()
        return response.json()


def chunked(values: list[str], size: int) -> Iterable[list[str]]:
    for start in range(0, len(values), size):
        yield values[start : start + size]


def iso_z(dt: datetime) -> str:
    return dt.astimezone(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def parse_dt(value: str) -> datetime | None:
    if not value:
        return None
    return datetime.fromisoformat(value.replace("Z", "+00:00")).astimezone(UTC)


def load_csv_rows(path: Path) -> list[dict]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def write_csv(path: Path, rows: list[dict], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def latest_output_dir(output_root: Path, cabinet: str, prefix: str) -> Path:
    pattern = f"ozon_fbs_{cabinet.lower()}_{prefix.lower()}_*"
    matches = sorted(output_root.glob(pattern))
    if not matches:
        raise RuntimeError(f"No output directories found for {pattern}")
    return matches[-1]


def family_suffix(offer_id: str, family_prefixes: list[str] | tuple[str, ...]) -> str:
    for prefix in family_prefixes:
        if offer_id.startswith(prefix):
            return "X" + offer_id[len(prefix) :]
    x_index = offer_id.find("X")
    if x_index >= 0:
        return offer_id[x_index:]
    return offer_id


def frame_code(offer_id: str, family_prefixes: list[str] | tuple[str, ...]) -> str:
    suffix = family_suffix(offer_id, family_prefixes)
    if "-" in suffix:
        return suffix.split("-", 1)[0]
    return suffix
