"""
export_static_data.py
======================
Run this LOCALLY, on your own machine, against your real Shopee database(s).
It reuses dashboard_api.py's normal data-loading logic (so the output always
matches your live dashboard), then anonymizes/pseudonymizes sensitive data
BEFORE writing anything to disk.

Output: data/bootstrap.json
  -> This is the ONLY file that should ever be committed to a public repo.
  -> Never commit your .db / .xlsx source files. Never commit dashboard_api.py's
     MANUAL_*_SOURCES if you leave real absolute paths in it.

What gets removed/replaced, and why:

  1) Customer personal data (PDPA):
     buyer_username, receiver_name, phone_number, shipping_address, postal_code
     are personal data of real customers under Thailand's PDPA. `customer` is
     replaced with a one-way pseudonym, `receiver` is dropped, and the raw
     columns are redacted everywhere they show up (including "Table Viewer"
     sample rows).

  2) Confidential business data (product names, order numbers):
     product_name, order_id, tracking_no/tracking_number, and sku_ref/
     parent_sku_ref reveal what you sell and your real order/shipment
     identifiers. These are replaced with sequential SAP-style codes
     (e.g. "Material 000001", "SO-000001", "DN-000001", "MATNR-000001").
     The same real value always maps to the same fake code within one
     export, so grouping/filtering/charts (top products, orders per
     customer, duplicate-order detection, packing lead time, cost per
     product, etc.) still work correctly -- only the labels change.

Usage:
    python export_static_data.py
    # then check data/bootstrap.json looks right before `git add` / `git commit`
"""

from __future__ import annotations

import hashlib
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

import dashboard_api as api  # noqa: E402  (reuse the real loading/normalizing logic)

# ------------------------------------------------------------------
# 1) Customer personal data -> pseudonym / redaction
# ------------------------------------------------------------------

PII_COLUMNS = {
    "buyer_username", "receiver_name", "phone_number", "shipping_address",
    "postal_code", "buyer_phone", "receiver_phone", "email", "buyer_email",
    "customer_email", "id_card", "national_id",
}

# Fixed salt so the same real customer always maps to the same pseudonym
# within one export, without the pseudonym being reversible back to the original.
SALT = "shopee-portfolio-demo-v1"


def pseudonymize_customer(value: str) -> str:
    if not value or not str(value).strip():
        return "Unknown customer"
    digest = hashlib.sha256(f"{SALT}:{value}".encode("utf-8")).hexdigest()[:8]
    return f"Customer-{digest}"


# ------------------------------------------------------------------
# 2) Confidential business identifiers -> sequential SAP-style codes
# ------------------------------------------------------------------

class SequentialCodeMap:
    """Deterministic, order-of-first-appearance mapping from a real value to
    a fake SAP-style code. The same real value always maps to the same fake
    code within a single export run, so charts/filters/grouping that key off
    this field stay meaningful -- only the display text changes."""

    def __init__(self, prefix: str, width: int = 6, skip_values: set[str] | None = None):
        self.prefix = prefix
        self.width = width
        self.skip_values = skip_values or set()
        self._map: dict[str, str] = {}
        self._next = 1

    def get(self, value) -> str:
        text = "" if value is None else str(value).strip()
        if not text or text in self.skip_values:
            return text
        if text not in self._map:
            self._map[text] = f"{self.prefix}{self._next:0{self.width}d}"
            self._next += 1
        return self._map[text]

    def lookup_or_redact(self, value) -> str:
        """For raw schema sample rows: only ever return a code already
        assigned via .get(), or a generic redaction. Never pass through
        the original text, even if we haven't seen this exact value yet."""
        text = "" if value is None else str(value).strip()
        if not text or text in self.skip_values:
            return text
        return self._map.get(text, "[redacted]")


def anonymize_orders(orders: list[dict], product_map, sku_map, order_map, tracking_map) -> list[dict]:
    for order in orders:
        order["customer"] = pseudonymize_customer(order.get("customer", ""))
        order["receiver"] = ""  # recipient name not needed for any chart
        order["product"] = product_map.get(order.get("product"))
        order["sku"] = sku_map.get(order.get("sku"))
        order["id"] = order_map.get(order.get("id"))
        order["trackingNo"] = tracking_map.get(order.get("trackingNo"))
        if order.get("sourceFile"):
            order["sourceFile"] = Path(order["sourceFile"]).name
    return orders


def anonymize_packing(rows: list[dict], product_map, tracking_map) -> list[dict]:
    for row in rows:
        if row.get("product"):
            row["product"] = product_map.get(row["product"])
        if row.get("trackingNumber"):
            row["trackingNumber"] = tracking_map.get(row["trackingNumber"])
    return rows


def anonymize_cost(rows: list[dict], product_map, order_map) -> list[dict]:
    for row in rows:
        if row.get("product"):
            row["product"] = product_map.get(row["product"])
        if row.get("orderId"):
            row["orderId"] = order_map.get(row["orderId"])
    return rows


def anonymize_duplicates(meta: dict, order_map) -> dict:
    for dup in meta.get("duplicates", []):
        if dup.get("order_id"):
            dup["order_id"] = order_map.get(dup["order_id"])
    return meta


def anonymize_schema(schema: dict, product_map, sku_map, order_map, tracking_map) -> dict:
    business_cols = {
        "product_name": product_map,
        "sku_ref": sku_map,
        "parent_sku_ref": sku_map,
        "order_id": order_map,
        "tracking_no": tracking_map,
        "tracking_number": tracking_map,
    }
    for table in schema.get("tables", []):
        for row in table.get("sampleRows", []):
            for col in list(row.keys()):
                lower = col.lower()
                if lower in PII_COLUMNS:
                    row[col] = "[redacted]"
                elif lower in business_cols:
                    row[col] = business_cols[lower].lookup_or_redact(row[col])
    for db in schema.get("databases", []):
        db["path"] = db.get("label", "database")  # don't leak local filesystem layout
    return schema


def anonymize_health(health: dict) -> dict:
    for entry in health.get("databases", []):
        entry["dbPath"] = entry.get("label", entry.get("dbFile", ""))
    health["dbPath"] = ""
    return health


def anonymize_import_log(rows: list[dict]) -> list[dict]:
    for row in rows:
        if row.get("source_file"):
            row["source_file"] = Path(row["source_file"]).name
    return rows


def main():
    print("Loading data via dashboard_api.bootstrap() ...")
    data = api.bootstrap()

    if not data.get("orders"):
        print("No orders found. Check dashboard_api.py's MANUAL_DB_SOURCES / "
              "SHOPEE_DB_PATHS point at a real database before exporting.")

    # Build the code maps ONCE from the full order list (already sorted by
    # date), then reuse them everywhere else so the same real product/order/
    # tracking/sku always gets the same fake code across the whole file.
    product_map = SequentialCodeMap("Material ", width=6, skip_values={"Unknown product"})
    sku_map = SequentialCodeMap("MATNR-", width=6)
    order_map = SequentialCodeMap("SO-", width=6)
    tracking_map = SequentialCodeMap("DN-", width=6)

    data["orders"] = anonymize_orders(data.get("orders", []), product_map, sku_map, order_map, tracking_map)
    data["packing"] = anonymize_packing(data.get("packing", []), product_map, tracking_map)
    data["cost"] = anonymize_cost(data.get("cost", []), product_map, order_map)
    data["meta"] = anonymize_duplicates(data.get("meta", {}), order_map)
    data["schema"] = anonymize_schema(data.get("schema", {}), product_map, sku_map, order_map, tracking_map)
    data["health"] = anonymize_health(data.get("health", {}))
    data["importLog"] = anonymize_import_log(data.get("importLog", []))

    out_path = ROOT / "data" / "bootstrap.json"
    out_path.parent.mkdir(exist_ok=True)
    out_path.write_text(json.dumps(data, ensure_ascii=False, default=str), encoding="utf-8")

    print(f"Wrote anonymized dataset -> {out_path}")
    print(f"  orders: {len(data['orders'])}")
    print(f"  unique products -> {product_map._next - 1} Material codes")
    print(f"  unique orders   -> {order_map._next - 1} SO codes")
    print(f"  unique tracking -> {tracking_map._next - 1} DN codes")
    print(f"  unique SKUs     -> {sku_map._next - 1} MATNR codes")
    print("Open dashboard.html locally to preview it, then commit data/bootstrap.json.")
    print("Do NOT commit your .db/.xlsx files or dashboard_api.py with real absolute paths in it.")


if __name__ == "__main__":
    main()
