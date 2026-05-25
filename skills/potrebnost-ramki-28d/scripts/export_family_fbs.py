from __future__ import annotations

import argparse
import json
import math
from collections import defaultdict
from datetime import UTC, date, datetime, time, timedelta
from pathlib import Path

from ozon_frame_common import OzonSellerClient, chunked, iso_z, load_cabinet_creds, write_csv


POSTINGS_LIMIT = 1000
STOCKS_LIMIT = 1000
PRODUCTS_LIMIT = 1000
FETCH_WINDOW_DAYS = 365


def find_matching_products(client: OzonSellerClient, prefix: str) -> list[dict]:
    matches: list[dict] = []
    last_id = ""
    while True:
        payload = {
            "filter": {"visibility": "ALL"},
            "last_id": last_id,
            "limit": PRODUCTS_LIMIT,
        }
        result = client.post("/v3/product/list", payload).get("result", {})
        items = result.get("items", [])
        for item in items:
            offer_id = str(item.get("offer_id", ""))
            if offer_id.startswith(prefix):
                matches.append(item)
        last_id = result.get("last_id", "")
        if not items or not last_id:
            break
    return matches


def fetch_product_info(client: OzonSellerClient, offer_ids: list[str]) -> dict[str, dict]:
    info_by_offer: dict[str, dict] = {}
    for offer_chunk in chunked(offer_ids, 1000):
        payload = {"offer_id": offer_chunk}
        items = client.post("/v3/product/info/list", payload).get("items", [])
        for item in items:
            info_by_offer[str(item.get("offer_id", ""))] = item
    return info_by_offer


def fetch_warehouses(client: OzonSellerClient) -> dict[int, dict]:
    data = client.post("/v2/warehouse/list", {})
    return {int(item["warehouse_id"]): item for item in data.get("warehouses", [])}


def fetch_current_stocks(
    client: OzonSellerClient, warehouse_map: dict[int, dict], target_offers: set[str]
) -> list[dict]:
    rows: list[dict] = []
    for warehouse_id, warehouse in warehouse_map.items():
        cursor = ""
        while True:
            payload = {"warehouse_id": warehouse_id, "limit": STOCKS_LIMIT, "cursor": cursor}
            data = client.post("/v1/product/info/warehouse/stocks", payload)
            for stock in data.get("stocks", []):
                offer_id = str(stock.get("offer_id", ""))
                if offer_id in target_offers:
                    row = dict(stock)
                    row["warehouse_name"] = warehouse["name"]
                    rows.append(row)
            if not data.get("has_next"):
                break
            cursor = data.get("cursor", "")
            if not cursor:
                break
    return rows


def fetch_sales_rows(
    client: OzonSellerClient, prefix: str, start_date: date, end_date: date
) -> list[dict]:
    rows: list[dict] = []
    window_start = start_date
    while window_start <= end_date:
        window_end = min(window_start + timedelta(days=FETCH_WINDOW_DAYS - 1), end_date)
        offset = 0
        while True:
            payload = {
                "dir": "ASC",
                "filter": {
                    "since": iso_z(datetime.combine(window_start, time.min, tzinfo=UTC)),
                    "to": iso_z(datetime.combine(window_end, time.max, tzinfo=UTC)),
                },
                "limit": POSTINGS_LIMIT,
                "offset": offset,
                "with": {
                    "analytics_data": False,
                    "barcodes": False,
                    "financial_data": False,
                    "translit": False,
                },
            }
            result = client.post("/v3/posting/fbs/list", payload).get("result", {})
            postings = result.get("postings", [])
            for posting in postings:
                status = str(posting.get("status", ""))
                if "cancel" in status.lower():
                    continue
                delivery_method = posting.get("delivery_method", {}) or {}
                for product in posting.get("products", []):
                    offer_id = str(product.get("offer_id", ""))
                    if not offer_id.startswith(prefix):
                        continue
                    rows.append(
                        {
                            "posting_number": posting.get("posting_number", ""),
                            "order_id": posting.get("order_id", ""),
                            "order_number": posting.get("order_number", ""),
                            "status": status,
                            "substatus": posting.get("substatus", ""),
                            "in_process_at": posting.get("in_process_at", ""),
                            "shipment_date": posting.get("shipment_date", ""),
                            "warehouse_id": delivery_method.get("warehouse_id", ""),
                            "warehouse_name": delivery_method.get("warehouse", ""),
                            "delivery_method_id": delivery_method.get("id", ""),
                            "delivery_method_name": delivery_method.get("name", ""),
                            "offer_id": offer_id,
                            "sku": product.get("sku", ""),
                            "product_name": product.get("name", ""),
                            "quantity": int(product.get("quantity", 0) or 0),
                            "price": product.get("price", ""),
                            "currency_code": product.get("currency_code", ""),
                        }
                    )
            if not result.get("has_next"):
                break
            offset += POSTINGS_LIMIT
        window_start = window_end + timedelta(days=1)
    return rows


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Export one Ozon FBS family to CSV files.")
    parser.add_argument("cabinet", help="Cabinet name as it appears in .env")
    parser.add_argument("prefix", help="Family prefix such as K_T_M1X or K_T_A13X")
    parser.add_argument("--env", type=Path, default=Path(".env"), help="Path to .env")
    parser.add_argument("--output-root", type=Path, default=Path("outputs"), help="Directory for output folders")
    parser.add_argument("--start-date", default="2022-01-01", help="Inclusive start date in YYYY-MM-DD")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    started_at = datetime.now(UTC)
    start_date = date.fromisoformat(args.start_date)
    output_dir = args.output_root / f"ozon_fbs_{args.cabinet.lower()}_{args.prefix.lower()}_{started_at:%Y%m%d_%H%M%S}"

    creds = load_cabinet_creds(args.env, args.cabinet)
    client = OzonSellerClient(creds)

    product_matches = find_matching_products(client, args.prefix)
    matched_offer_ids = sorted(str(item["offer_id"]) for item in product_matches)
    product_info = fetch_product_info(client, matched_offer_ids) if matched_offer_ids else {}
    warehouse_map = fetch_warehouses(client)
    stock_rows = fetch_current_stocks(client, warehouse_map, set(matched_offer_ids)) if matched_offer_ids else []
    sales_rows = fetch_sales_rows(client, args.prefix, start_date, started_at.date())

    stock_by_offer_warehouse: dict[tuple[str, int], dict] = {}
    for row in stock_rows:
        stock_by_offer_warehouse[(str(row["offer_id"]), int(row["warehouse_id"]))] = row

    sales_summary: dict[tuple[str, int], dict] = defaultdict(
        lambda: {
            "offer_id": "",
            "warehouse_id": 0,
            "warehouse_name": "",
            "product_name": "",
            "sku": "",
            "sales_qty": 0,
            "sales_postings": 0,
            "first_sale_at": "",
            "last_sale_at": "",
        }
    )
    warehouse_totals: dict[tuple[int, str], dict] = defaultdict(
        lambda: {"warehouse_id": 0, "warehouse_name": "", "sales_qty": 0, "sales_postings": 0}
    )

    for row in sales_rows:
        warehouse_id = int(row["warehouse_id"] or 0)
        key = (str(row["offer_id"]), warehouse_id)
        summary = sales_summary[key]
        summary["offer_id"] = row["offer_id"]
        summary["warehouse_id"] = warehouse_id
        summary["warehouse_name"] = row["warehouse_name"]
        summary["product_name"] = row["product_name"]
        summary["sku"] = row["sku"]
        summary["sales_qty"] += int(row["quantity"])
        summary["sales_postings"] += 1
        in_process_at = row["in_process_at"]
        if in_process_at and (not summary["first_sale_at"] or in_process_at < summary["first_sale_at"]):
            summary["first_sale_at"] = in_process_at
        if in_process_at and (not summary["last_sale_at"] or in_process_at > summary["last_sale_at"]):
            summary["last_sale_at"] = in_process_at

        warehouse_total = warehouse_totals[(warehouse_id, row["warehouse_name"])]
        warehouse_total["warehouse_id"] = warehouse_id
        warehouse_total["warehouse_name"] = row["warehouse_name"]
        warehouse_total["sales_qty"] += int(row["quantity"])
        warehouse_total["sales_postings"] += 1

    stock_summary_rows: list[dict] = []
    seen_stock_keys: set[tuple[str, int]] = set()
    for (offer_id, warehouse_id), row in stock_by_offer_warehouse.items():
        info = product_info.get(offer_id, {})
        stock_summary_rows.append(
            {
                "offer_id": offer_id,
                "sku": row.get("sku") or info.get("sku", ""),
                "product_id": row.get("product_id") or info.get("id", ""),
                "product_name": info.get("name", ""),
                "warehouse_id": warehouse_id,
                "warehouse_name": row.get("warehouse_name", ""),
                "present": int(row.get("present", 0) or 0),
                "reserved": int(row.get("reserved", 0) or 0),
                "free_stock": int(row.get("free_stock", 0) or 0),
                "updated_at": row.get("updated_at", ""),
            }
        )
        seen_stock_keys.add((offer_id, warehouse_id))

    for (offer_id, warehouse_id), summary in sales_summary.items():
        if (offer_id, warehouse_id) in seen_stock_keys:
            continue
        info = product_info.get(offer_id, {})
        stock_summary_rows.append(
            {
                "offer_id": offer_id,
                "sku": summary.get("sku") or info.get("sku", ""),
                "product_id": info.get("id", ""),
                "product_name": summary.get("product_name") or info.get("name", ""),
                "warehouse_id": warehouse_id,
                "warehouse_name": summary["warehouse_name"],
                "present": 0,
                "reserved": 0,
                "free_stock": 0,
                "updated_at": "",
            }
        )

    sales_by_offer_warehouse_rows: list[dict] = []
    replenishment_rows: list[dict] = []
    for (offer_id, warehouse_id), summary in sorted(sales_summary.items()):
        stock = stock_by_offer_warehouse.get((offer_id, warehouse_id), {})
        free_stock = int(stock.get("free_stock", 0) or 0)
        first_sale_at = summary["first_sale_at"]
        span_days = 0
        if first_sale_at:
            first_dt = datetime.fromisoformat(first_sale_at.replace("Z", "+00:00"))
            span_days = max((started_at.date() - first_dt.date()).days + 1, 1)
        monthly_need = math.ceil(summary["sales_qty"] / span_days * 30) if span_days else 0

        base_row = {
            **summary,
            "current_free_stock": free_stock,
            "current_present": int(stock.get("present", 0) or 0),
            "current_reserved": int(stock.get("reserved", 0) or 0),
            "history_span_days": span_days,
            "monthly_need_qty": monthly_need,
        }
        sales_by_offer_warehouse_rows.append(base_row)

        if summary["sales_qty"] > 0 and free_stock == 0:
            replenishment_rows.append({**base_row, "recommended_purchase_qty": monthly_need})

    warehouse_sales_rows = sorted(
        warehouse_totals.values(),
        key=lambda row: (-int(row["sales_qty"]), str(row["warehouse_name"])),
    )
    stock_summary_rows.sort(key=lambda row: (row["offer_id"], row["warehouse_name"]))
    sales_rows.sort(key=lambda row: (row["in_process_at"], row["posting_number"], row["offer_id"]))
    sales_by_offer_warehouse_rows.sort(key=lambda row: (-int(row["sales_qty"]), row["offer_id"], row["warehouse_name"]))
    replenishment_rows.sort(key=lambda row: (-int(row["recommended_purchase_qty"]), row["offer_id"], row["warehouse_name"]))

    write_csv(
        output_dir / "sales_detailed.csv",
        sales_rows,
        [
            "posting_number",
            "order_id",
            "order_number",
            "status",
            "substatus",
            "in_process_at",
            "shipment_date",
            "warehouse_id",
            "warehouse_name",
            "delivery_method_id",
            "delivery_method_name",
            "offer_id",
            "sku",
            "product_name",
            "quantity",
            "price",
            "currency_code",
        ],
    )
    write_csv(
        output_dir / "sales_by_warehouse.csv",
        warehouse_sales_rows,
        ["warehouse_id", "warehouse_name", "sales_qty", "sales_postings"],
    )
    write_csv(
        output_dir / "sales_by_offer_and_warehouse.csv",
        sales_by_offer_warehouse_rows,
        [
            "offer_id",
            "sku",
            "product_name",
            "warehouse_id",
            "warehouse_name",
            "sales_qty",
            "sales_postings",
            "first_sale_at",
            "last_sale_at",
            "history_span_days",
            "current_free_stock",
            "current_present",
            "current_reserved",
            "monthly_need_qty",
        ],
    )
    write_csv(
        output_dir / "current_fbs_stocks.csv",
        stock_summary_rows,
        [
            "offer_id",
            "sku",
            "product_id",
            "product_name",
            "warehouse_id",
            "warehouse_name",
            "present",
            "reserved",
            "free_stock",
            "updated_at",
        ],
    )
    write_csv(
        output_dir / "replenishment_recommendations.csv",
        replenishment_rows,
        [
            "offer_id",
            "sku",
            "product_name",
            "warehouse_id",
            "warehouse_name",
            "sales_qty",
            "sales_postings",
            "first_sale_at",
            "last_sale_at",
            "history_span_days",
            "current_free_stock",
            "monthly_need_qty",
            "recommended_purchase_qty",
        ],
    )

    summary = {
        "cabinet": args.cabinet,
        "query": args.prefix,
        "exact_offer_found": args.prefix in set(matched_offer_ids),
        "matched_offer_count": len(matched_offer_ids),
        "matched_offer_sample": matched_offer_ids[:20],
        "sales_rows": len(sales_rows),
        "sales_offer_warehouse_rows": len(sales_by_offer_warehouse_rows),
        "stock_rows": len(stock_summary_rows),
        "replenishment_rows": len(replenishment_rows),
        "generated_at": started_at.isoformat(),
        "output_dir": str(output_dir.resolve()),
    }
    output_dir.mkdir(parents=True, exist_ok=True)
    (output_dir / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
