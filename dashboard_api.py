from __future__ import annotations

import argparse
import json
import mimetypes
import os
import sqlite3
import sys
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse


DEFAULT_DB_PATH = r"D:\Shopee\Shopee Sale Data\Month-sale\sales_data.db"
ROOT = Path(__file__).resolve().parent

# ============================================================
# 📌 เพิ่มฐานข้อมูลตรงนี้ได้เลย (แก้ไฟล์นี้แล้ว save ก็พอ)
# ============================================================
# ใส่ path ของไฟล์ .db แต่ละไฟล์เป็นสมาชิกใน list ด้านล่าง
# รองรับทั้ง .db และ .xlsx
# ถ้าเป็น .xlsx ระบบจะแปลงเป็น SQLite อัตโนมัติ
# รูปแบบ: "ชื่อที่อยากให้โชว์ในแดชบอร์ด=path ของไฟล์ .db"
#   (หรือใส่แค่ path เฉยๆ ก็ได้ ระบบจะตั้งชื่อให้อัตโนมัติจากชื่อไฟล์)
#
# ตัวอย่างการเพิ่มฐานข้อมูลอีก 2 ไฟล์:
#
# MANUAL_DB_SOURCES = [
#     r"2024=D:\Shopee\Shopee Sale Data\Month-sale\sales_data.db",
#     r"2025=D:\Shopee\Shopee Sale Data\Month-sale\sales_data_2025.db",
#     r"Store2=D:\Shopee\Store2\sales_data.db",
# ]
#
# เพิ่มได้ไม่จำกัดจำนวนไฟล์ แค่เพิ่มบรรทัดใหม่ในลิสต์ด้านล่างนี้:
# (เฉพาะไฟล์ที่มีตาราง "sales_data" เท่านั้น เช่นไฟล์ยอดขายรายเดือน/รายปีอื่น ๆ)
MANUAL_DB_SOURCES: list[str] = [
    r"Sales=D:\Shopee\Shopee Sale Data\Month-sale\sales_data.db",
]

# Traffic (ยอดเข้าชม/คลิก/conversion), Packing (เวลาแพ็คของ) และ Cost (ต้นทุนสินค้า/ค่าส่งขาเข้า)
# มีโครงสร้างตารางคนละแบบกับ sales_data จึงตั้งค่าแยกกันคนละลิสต์ด้านล่างนี้:
MANUAL_TRAFFIC_SOURCES: list[str] = [
    r"Traffic=D:\Shopee\Shopee Sale Data\DAYS_Traffic\traffic.db",
]
MANUAL_PACKING_SOURCES: list[str] = [
    r"Packing=D:\Shopee\Shopee Sale Data\Parcel\packing.db",
]
MANUAL_COST_SOURCES: list[str] = [
    r"Cost=D:\Shopee\Shopee Sale Data\Cost\cost.xlsx",
]
# ============================================================

# ลำดับความสำคัญ (สูง -> ต่ำ) ถ้ามีการตั้งค่าซ้อนกันหลายที่:
#   1) --db / --traffic-db / --packing-db / --cost-db ตอนรันคำสั่ง
#   2) environment variable SHOPEE_DB_PATHS / SHOPEE_TRAFFIC_PATHS / SHOPEE_PACKING_PATHS / SHOPEE_COST_PATHS
#   3) MANUAL_*_SOURCES ที่แก้ไว้ในไฟล์นี้ (ด้านบน)


def parse_db_sources(raw: str) -> list[dict]:
    """Parse a comma-separated list of DB paths into labeled sources.

    Each item can be a plain path ("C:\\a.db") or "label=path"
    ("2024=C:\\a.db") to control the label shown in the UI. Labels default
    to the file stem and are de-duplicated if they collide.
    """
    sources = []
    seen_labels = set()
    for part in raw.split(","):
        part = part.strip()
        if not part:
            continue
        if "=" in part:
            label, path_str = part.split("=", 1)
            label, path_str = label.strip(), path_str.strip()
        else:
            path_str = part
            label = Path(path_str).stem or path_str
        base_label, n = label, 2
        while label in seen_labels:
            label = f"{base_label}_{n}"
            n += 1
        seen_labels.add(label)
        real_path = ensure_sqlite(Path(path_str))

        sources.append({
            "label": label,
            "path": real_path,
            "originalPath": Path(path_str)
            })
    return sources

import pandas as pd
import tempfile

_excel_cache = {}


def ensure_sqlite(path: Path) -> Path:
    """
    ถ้าเป็น .db คืน path เดิม
    ถ้าเป็น .xlsx แปลงเป็น sqlite ชั่วคราว
    """

    if path.suffix.lower() in [".db", ".sqlite"]:
        return path

    if path.suffix.lower() not in [".xlsx", ".xls"]:
        return path

    if not path.exists():
        # Leave the (missing) .xlsx path as-is; health() will report it as not connected
        # instead of crashing the whole server at startup.
        return path

    db_path = Path(tempfile.gettempdir()) / (path.stem + ".db")

    if (
        not db_path.exists()
        or path.stat().st_mtime > db_path.stat().st_mtime
    ):
        excel = pd.ExcelFile(path)

        import sqlite3

        with sqlite3.connect(db_path) as conn:
            for sheet in excel.sheet_names:
                df = pd.read_excel(excel, sheet_name=sheet)
                df.to_sql(
                    sheet,
                    conn,
                    if_exists="replace",
                    index=False
                )

    return db_path

# Multiple SQLite databases can be combined into one dashboard.
# Configure via SHOPEE_DB_PATHS="C:\a.db,C:\b.db" (or "label=path,label2=path2"),
# falling back to MANUAL_DB_SOURCES above, then the legacy SHOPEE_DB_PATH / DEFAULT_DB_PATH.
_raw_db_paths = (
    os.environ.get("SHOPEE_DB_PATHS")
    or os.environ.get("SHOPEE_DB_PATH")
    or ",".join(MANUAL_DB_SOURCES)
    or DEFAULT_DB_PATH
)
DB_SOURCES: list[dict] = parse_db_sources(_raw_db_paths)

_raw_traffic_paths = os.environ.get("SHOPEE_TRAFFIC_PATHS") or ",".join(MANUAL_TRAFFIC_SOURCES)
TRAFFIC_SOURCES: list[dict] = parse_db_sources(_raw_traffic_paths) if _raw_traffic_paths else []

_raw_packing_paths = os.environ.get("SHOPEE_PACKING_PATHS") or ",".join(MANUAL_PACKING_SOURCES)
PACKING_SOURCES: list[dict] = parse_db_sources(_raw_packing_paths) if _raw_packing_paths else []

_raw_cost_paths = os.environ.get("SHOPEE_COST_PATHS") or ",".join(MANUAL_COST_SOURCES)
COST_SOURCES: list[dict] = parse_db_sources(_raw_cost_paths) if _raw_cost_paths else []


def as_float(value) -> float:
    if value is None or value == "":
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def as_int(value) -> int:
    return int(round(as_float(value)))


def clean_text(value, fallback: str = "") -> str:
    if value is None:
        return fallback
    text = str(value).strip()
    return text if text else fallback


def parse_dt(value: str | None):
    if not value:
        return None
    text = str(value).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d", "%Y%m%d"):
        try:
            return datetime.strptime(text[:19] if "%H" in fmt else text[:10], fmt)
        except ValueError:
            pass
    return None


def iso_date(value: str | None) -> str:
    dt = parse_dt(value)
    return dt.strftime("%Y-%m-%d") if dt else ""


def hour_of_day(value: str | None):
    dt = parse_dt(value)
    return dt.hour if dt else None


def connect(path: Path):
    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    return conn


def existing_sources() -> list[dict]:
    return [src for src in DB_SOURCES if src["path"].exists()]


def existing(sources: list[dict]) -> list[dict]:
    return [src for src in sources if src["path"].exists()]


import re

WEEKDAY_NAMES = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]


def pct_to_float(value) -> float:
    if value is None:
        return 0.0
    text = str(value).strip().replace("%", "").replace(",", "")
    try:
        return float(text)
    except ValueError:
        return 0.0


def leading_number(value):
    """Best-effort numeric extraction from messy qty strings like '50pcs', '1kg', '1'."""
    if value is None:
        return None
    match = re.match(r"[-+]?\d*\.?\d+", str(value).strip())
    return float(match.group()) if match else None


def parse_flexible_dt(value: str | None):
    """Parse timestamps in either 'YYYY-MM-DD HH:MM:SS' or EXIF-style 'YYYY:MM:DD HH:MM:SS'."""
    if not value:
        return None
    text = str(value).strip()
    if len(text) >= 10 and text[4] == ":" and text[7] == ":":
        text = text[:4] + "-" + text[5:7] + "-" + text[8:]
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def parse_thai_range_date(value: str | None):
    """traffic_summary.'วันที่' looks like '15-04-2024-15-04-2024' (DD-MM-YYYY start-end). Return the start date."""
    if not value:
        return None
    text = str(value).strip()
    parts = text.split("-")
    if len(parts) >= 3:
        day, month, year = parts[0], parts[1], parts[2]
        try:
            return datetime(int(year), int(month), int(day))
        except ValueError:
            return None
    return None


def normalize_status(row: sqlite3.Row) -> str:
    raw = clean_text(row["order_status"])
    refund = clean_text(row["refund_return_status"])
    combined = f"{raw} {refund}"
    if "ยกเลิก" in combined or "cancel" in combined.lower():
        return "Cancelled"
    if "คืน" in combined or "return" in combined.lower() or as_float(row["return_qty"]) > 0:
        return "Returned"
    if "สำเร็จ" in combined or "ได้รับสินค้า" in combined or "completed" in combined.lower():
        return "Completed"
    return "Pending"


def normalize_payment(value: str | None) -> str:
    text = clean_text(value, "Unknown")
    lower = text.lower()
    if "ปลายทาง" in text or "cod" in lower:
        return "COD"
    if "บัตร" in text or "card" in lower or "google pay" in lower:
        return "Credit Card"
    if "shopeepay" in lower or "wallet" in lower:
        return "Wallet / ShopeePay"
    if "bank" in lower or "พร้อมเพย์" in text or "qr" in lower:
        return "Bank Transfer"
    if "spaylater" in lower:
        return "PayLater"
    return text


def normalize_province(value: str | None) -> str:
    text = clean_text(value, "Unknown")
    return text.replace("จังหวัด", "").strip() or text


def campaign_name(row: sqlite3.Row) -> str:
    ignored = {"-", "0", "0.0", "none", "nan", "null", "n/a"}
    for key in ("discount_code", "discount_code_seller", "discount_code_shopee", "coins_cashback_code_seller"):
        value = clean_text(row[key])
        if value and value.lower() not in ignored:
            return value
    if clean_text(row["is_bundle_campaign"]).lower() in {"yes", "y", "true", "1"}:
        return "Bundle Campaign"
    if as_float(row["coin_discount"]) > 0:
        return "Coins Discount"
    if as_float(row["payment_channel_promo"]) > 0:
        return "Payment Promo"
    if as_float(row["hot_listing"]) > 0:
        return "Hot Listing"
    return "No Campaign"


def normalize_order(row: sqlite3.Row) -> dict:
    qty = as_float(row["qty"])
    list_price = as_float(row["list_price"])
    buyer_paid = as_float(row["buyer_paid_price_thb"])
    gross_sales = max(as_float(row["gross_sales"]), list_price * qty, buyer_paid)
    net_sales = as_float(row["net_sales"]) or as_float(row["total_amount"]) or max(gross_sales, buyer_paid)
    transaction_fee = as_float(row["transaction_fee"])
    commission_fee = as_float(row["commission_fee"])
    service_fee = as_float(row["service_fee"])
    seller_discount = (
        as_float(row["bundle_discount_seller"])
        + as_float(row["trade_in_bonus_seller"])
    )
    platform_discount = (
        as_float(row["bundle_discount_shopee"])
        + as_float(row["shopee_discount"])
        + as_float(row["coin_discount"])
        + as_float(row["trade_in_discount"])
        + as_float(row["trade_in_bonus"])
        + as_float(row["payment_channel_promo"])
    )
    shipping_cost = max(
        0.0,
        as_float(row["estimated_shipping_fee"])
        + as_float(row["return_shipping_fee"])
        - as_float(row["shopee_shipping_subsidy"]),
    )
    total_fees = transaction_fee + commission_fee + service_fee
    total_discount = seller_discount + platform_discount
    net_profit = net_sales - total_fees - shipping_cost
    date = iso_date(row["order_date"])
    product = clean_text(row["product_name"], "Unknown product")
    variant = clean_text(row["variant_name"])

    return {
        "id": clean_text(row["order_id"], str(row["id"])),
        "rowId": row["id"],
        "trackingNo": clean_text(row["tracking_no"]),
        "date": date,
        "month": clean_text(row["sale_month"], date[:7]),
        "hour": hour_of_day(row["order_date"]),
        "completedAt": iso_date(row["order_completed_at"]),
        "product": product,
        "variant": variant,
        "sku": clean_text(row["sku_ref"]) or clean_text(row["parent_sku_ref"]),
        "category": "Shopee Product",
        "qty": qty,
        "returnQty": as_float(row["return_qty"]),
        "unitPrice": list_price or (buyer_paid / qty if qty else buyer_paid),
        "buyerPaidPrice": buyer_paid,
        "grossSales": gross_sales,
        "netSales": net_sales,
        "totalAmount": as_float(row["total_amount"]),
        "customer": clean_text(row["buyer_username"], clean_text(row["receiver_name"], "Unknown customer")),
        "receiver": clean_text(row["receiver_name"]),
        "province": normalize_province(row["province"]),
        "district": clean_text(row["district"]),
        "payment": normalize_payment(row["payment_method"]),
        "paymentRaw": clean_text(row["payment_method"], "Unknown"),
        "status": normalize_status(row),
        "statusRaw": clean_text(row["order_status"], "Unknown"),
        "campaign": campaign_name(row),
        "shippingMethod": clean_text(row["shipping_method"], clean_text(row["shipping_option"], "Unknown")),
        "shippingOption": clean_text(row["shipping_option"], "Unknown"),
        "buyerShippingFee": as_float(row["buyer_shipping_fee"]),
        "shippingCost": shipping_cost,
        "commissionFee": commission_fee,
        "transactionFee": transaction_fee,
        "serviceFee": service_fee,
        "totalFees": total_fees,
        "sellerDiscount": seller_discount,
        "platformDiscount": platform_discount,
        "discount": total_discount,
        "netProfit": net_profit,
        "sourceFile": clean_text(row["source_file"]),
        "importedAt": clean_text(row["imported_at"]),
    }


def rows_to_dicts(rows):
    return [{key: row[key] for key in row.keys()} for row in rows]


def load_orders():
    all_orders = []
    for src in existing_sources():
        with connect(src["path"]) as conn:
            rows = conn.execute("SELECT * FROM sales_data ORDER BY order_date, id").fetchall()
        for row in rows:
            order = normalize_order(row)
            order["sourceDb"] = src["label"]
            order["rowId"] = f'{src["label"]}:{order["rowId"]}'
            all_orders.append(order)
    all_orders.sort(key=lambda o: (o["date"] or "", o["sourceDb"], o["rowId"]))
    return all_orders


def load_import_log():
    all_rows = []
    for src in existing_sources():
        with connect(src["path"]) as conn:
            try:
                rows = conn.execute(
                    """
                    SELECT id, source_file, sale_month, period_start, period_end, imported_at
                    FROM sales_import_log
                    ORDER BY imported_at DESC, id DESC
                    """
                ).fetchall()
            except sqlite3.OperationalError:
                rows = []
        for row in rows:
            entry = dict(row)
            entry["sourceDb"] = src["label"]
            all_rows.append(entry)
    all_rows.sort(key=lambda r: (str(r.get("imported_at") or ""), r.get("id") or 0), reverse=True)
    return all_rows


def load_schema():
    out = {"tables": [], "starSchema": star_schema(), "databases": []}
    for src in existing_sources():
        out["databases"].append({"label": src["label"], "path": str(src["path"])})
        with connect(src["path"]) as conn:
            table_rows = conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"
            ).fetchall()
            for table_row in table_rows:
                table = table_row["name"]
                if table == "sqlite_sequence":
                    continue
                row_count = conn.execute(f'SELECT COUNT(*) AS n FROM "{table}"').fetchone()["n"]
                columns = conn.execute(f'PRAGMA table_info("{table}")').fetchall()
                column_info = []
                for col in columns:
                    name = col["name"]
                    missing = conn.execute(
                        f'''
                        SELECT SUM(
                            CASE
                                WHEN "{name}" IS NULL OR TRIM(CAST("{name}" AS TEXT)) = '' THEN 1
                                ELSE 0
                            END
                        ) AS missing
                        FROM "{table}"
                        '''
                    ).fetchone()["missing"]
                    column_info.append(
                        {
                            "name": name,
                            "type": col["type"],
                            "notnull": col["notnull"],
                            "pk": col["pk"],
                            "missing": int(missing or 0),
                        }
                    )
                samples = conn.execute(f'SELECT * FROM "{table}" LIMIT 50').fetchall()
                out["tables"].append(
                    {
                        "name": table,
                        "db": src["label"],
                        "key": f'{src["label"]}::{table}',
                        "rowCount": row_count,
                        "columnCount": len(column_info),
                        "columns": column_info,
                        "sampleRows": rows_to_dicts(samples),
                    }
                )
    return out


TRAFFIC_HOURLY_COLS = {
    "date": "วันที่",
    "sales_thb": "ยอดขายทั้งหมด (THB)",
    "orders": "คำสั่งซื้อทั้งหมด",
    "aov": "ยอดขายเฉลี่ยต่อคำสั่งซื้อ",
    "clicks": "จำนวนคลิก",
    "visitors": "จำนวนผู้เยี่ยมชม",
    "conversion_rate": "อัตราการซื้อสินค้า",
    "cancelled_orders": "คำสั่งซื้อที่ยกเลิก",
    "cancelled_sales": "ยอดขายที่ยกเลิก",
    "returned_orders": "คำสั่งซื้อที่คืนเงิน/คืนสินค้า",
    "returned_sales": "ยอดขายที่คืนเงิน/คืนสินค้า",
}
TRAFFIC_DAILY_EXTRA_COLS = {
    "buyers": "# ของผู้ซื้อ",
    "new_buyers": "# ของผู้ซื้อใหม่",
    "returning_buyers": "# ของผู้ซื้อเดิม",
    "potential_buyers": "# ผู้ที่อาจจะซื้อ",
    "repeat_rate": "อัตราการกลับมาซื้อซ้ำ",
}


def normalize_traffic_row(row: sqlite3.Row, grain: str) -> dict:
    get = lambda col: row[col] if col in row.keys() else None
    if grain == "hourly":
        dt = parse_dt(get(TRAFFIC_HOURLY_COLS["date"]))
    else:
        dt = parse_thai_range_date(get(TRAFFIC_HOURLY_COLS["date"]))
    date = dt.strftime("%Y-%m-%d") if dt else ""
    out = {
        "grain": grain,
        "date": date,
        "hour": dt.hour if (dt and grain == "hourly") else None,
        "weekday": WEEKDAY_NAMES[dt.weekday()] if dt else "Unknown",
        "salesThb": as_float(get(TRAFFIC_HOURLY_COLS["sales_thb"])),
        "orders": as_int(get(TRAFFIC_HOURLY_COLS["orders"])),
        "aov": as_float(get(TRAFFIC_HOURLY_COLS["aov"])),
        "clicks": as_int(get(TRAFFIC_HOURLY_COLS["clicks"])),
        "visitors": as_int(get(TRAFFIC_HOURLY_COLS["visitors"])),
        "conversionRate": pct_to_float(get(TRAFFIC_HOURLY_COLS["conversion_rate"])),
        "cancelledOrders": as_int(get(TRAFFIC_HOURLY_COLS["cancelled_orders"])),
        "cancelledSales": as_float(get(TRAFFIC_HOURLY_COLS["cancelled_sales"])),
        "returnedOrders": as_int(get(TRAFFIC_HOURLY_COLS["returned_orders"])),
        "returnedSales": as_float(get(TRAFFIC_HOURLY_COLS["returned_sales"])),
    }
    if grain == "daily":
        out["buyers"] = as_int(get(TRAFFIC_DAILY_EXTRA_COLS["buyers"]))
        out["newBuyers"] = as_int(get(TRAFFIC_DAILY_EXTRA_COLS["new_buyers"]))
        out["returningBuyers"] = as_int(get(TRAFFIC_DAILY_EXTRA_COLS["returning_buyers"]))
        out["potentialBuyers"] = as_int(get(TRAFFIC_DAILY_EXTRA_COLS["potential_buyers"]))
        out["repeatRate"] = pct_to_float(get(TRAFFIC_DAILY_EXTRA_COLS["repeat_rate"]))
    return out


def load_traffic():
    hourly, daily = [], []
    for src in existing(TRAFFIC_SOURCES):
        with connect(src["path"]) as conn:
            tables = {r["name"] for r in conn.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            ).fetchall()}
            if "traffic_data" in tables:
                for row in conn.execute("SELECT * FROM traffic_data").fetchall():
                    entry = normalize_traffic_row(row, "hourly")
                    entry["sourceDb"] = src["label"]
                    if entry["date"]:
                        hourly.append(entry)
            if "traffic_summary" in tables:
                for row in conn.execute("SELECT * FROM traffic_summary").fetchall():
                    entry = normalize_traffic_row(row, "daily")
                    entry["sourceDb"] = src["label"]
                    if entry["date"]:
                        daily.append(entry)
    hourly.sort(key=lambda r: (r["date"], r["hour"] or 0))
    daily.sort(key=lambda r: r["date"])
    return {"hourly": hourly, "daily": daily}


def _sales_lookup_by_tracking() -> dict:
    """tracking_no -> {order_date, product_name, province, net_sales} from the sales sources, for joining packing time to order context."""
    lookup: dict[str, dict] = {}
    for src in existing_sources():
        with connect(src["path"]) as conn:
            try:
                rows = conn.execute(
                    """
                    SELECT tracking_no, order_date, product_name, province, net_sales
                    FROM sales_data
                    WHERE tracking_no IS NOT NULL AND TRIM(tracking_no) != ''
                    """
                ).fetchall()
            except sqlite3.OperationalError:
                rows = []
        for row in rows:
            key = clean_text(row["tracking_no"])
            if key and key not in lookup:
                lookup[key] = {
                    "orderDate": iso_date(row["order_date"]),
                    "product": clean_text(row["product_name"]),
                    "province": normalize_province(row["province"]),
                    "netSales": as_float(row["net_sales"]),
                }
    return lookup


def load_packing():
    sales_lookup = _sales_lookup_by_tracking()
    out = []
    for src in existing(PACKING_SOURCES):
        with connect(src["path"]) as conn:
            rows = conn.execute("SELECT * FROM packed_orders").fetchall()
        for row in rows:
            photo_dt = parse_flexible_dt(row["photo_taken_at"])
            processed_dt = parse_flexible_dt(row["processed_at"])
            import_lag_min = None
            if photo_dt and processed_dt:
                import_lag_min = round((processed_dt - photo_dt).total_seconds() / 60, 1)
                if import_lag_min < 0:
                    import_lag_min = None
            tracking = clean_text(row["tracking_number"])
            order = sales_lookup.get(tracking)
            lead_days = None
            if order and order["orderDate"] and photo_dt:
                order_dt = parse_dt(order["orderDate"])
                if order_dt:
                    lead_days = round((photo_dt - order_dt).total_seconds() / 86400, 2)
            out.append({
                "id": row["id"],
                "sourceDb": src["label"],
                "trackingNumber": tracking,
                "weight": as_float(row["weight"]),
                "photoAt": photo_dt.strftime("%Y-%m-%d %H:%M:%S") if photo_dt else "",
                "processedAt": processed_dt.strftime("%Y-%m-%d %H:%M:%S") if processed_dt else "",
                "date": photo_dt.strftime("%Y-%m-%d") if photo_dt else "",
                "hour": photo_dt.hour if photo_dt else None,
                "weekday": WEEKDAY_NAMES[photo_dt.weekday()] if photo_dt else "Unknown",
                # Time from when the packed-parcel photo was scanned/logged into the tracker.
                # This reflects batch import cadence, not how long the parcel took to pack -
                # kept for data-hygiene visibility, not a packing-speed metric.
                "importLagMinutes": import_lag_min,
                "orderDate": order["orderDate"] if order else "",
                "product": order["product"] if order else "",
                "province": order["province"] if order else "",
                "netSales": order["netSales"] if order else None,
                # Real fulfillment-speed signal: days between the customer's order and when it was packed.
                "leadDaysFromOrder": lead_days,
            })
    out.sort(key=lambda r: r["photoAt"])

    # Pace between consecutively packed parcels on the same day - the closest proxy we have
    # to actual packing speed (gap between finishing one parcel's photo and the next).
    prev_by_date: dict[str, datetime] = {}
    for row in out:
        if not row["date"] or not row["photoAt"]:
            row["paceMinutes"] = None
            continue
        photo_dt = datetime.strptime(row["photoAt"], "%Y-%m-%d %H:%M:%S")
        prev = prev_by_date.get(row["date"])
        row["paceMinutes"] = round((photo_dt - prev).total_seconds() / 60, 1) if prev else None
        prev_by_date[row["date"]] = photo_dt
    return out


def load_cost():
    out = []
    for src in existing(COST_SOURCES):
        try:
            df = pd.read_excel(src["originalPath"])
        except Exception:
            continue
        df.columns = [str(c).strip() for c in df.columns]
        for _, row in df.iterrows():
            date = row.get("date")
            date_str = ""
            if pd.notna(date):
                try:
                    date_str = pd.to_datetime(date).strftime("%Y-%m-%d")
                except Exception:
                    date_str = ""
            qty_raw = row.get("qty")
            qty_num = leading_number(qty_raw)
            cost_product = as_float(row.get("cost_product"))
            cost_ship_in = as_float(row.get("cost_shipping_inbound"))
            cost_other = as_float(row.get("cost_other"))
            total_cost = as_float(row.get("total_cost")) or (cost_product + cost_ship_in + cost_other)
            out.append({
                "sourceDb": src["label"],
                "product": clean_text(row.get("product_name"), "Unknown product"),
                "date": date_str,
                "month": date_str[:7] if date_str else "Unknown",
                "qtyRaw": clean_text(row.get("qty")),
                "qty": qty_num,
                "costProduct": cost_product,
                "costShippingInbound": cost_ship_in,
                "costOther": cost_other,
                "totalCost": total_cost,
                "costPerUnit": (total_cost / qty_num) if qty_num else None,
                "supplier": clean_text(row.get("supplier_name"), "Unknown supplier"),
                "notes": clean_text(row.get("notes")),
                "orderId": clean_text(row.get("Order_id")),
            })
    out.sort(key=lambda r: r["date"])
    return out


def duplicate_orders():
    agg: dict[str, dict] = {}
    for src in existing_sources():
        with connect(src["path"]) as conn:
            rows = conn.execute(
                """
                SELECT order_id, COUNT(*) AS row_count, SUM(COALESCE(net_sales, 0)) AS net_sales
                FROM sales_data
                WHERE order_id IS NOT NULL AND TRIM(order_id) != ''
                GROUP BY order_id
                """
            ).fetchall()
        for row in rows:
            entry = agg.setdefault(row["order_id"], {"row_count": 0, "net_sales": 0.0, "sourceDbs": []})
            entry["row_count"] += row["row_count"]
            entry["net_sales"] += row["net_sales"] or 0.0
            entry["sourceDbs"].append(src["label"])

    duplicates = [
        {"order_id": order_id, **data}
        for order_id, data in agg.items()
        if data["row_count"] > 1
    ]
    duplicates.sort(key=lambda d: (-d["row_count"], -d["net_sales"]))
    return duplicates[:30]


def star_schema():
    return {
        "fact_sales": [
            "order_id",
            "order_date",
            "product_name",
            "buyer_username",
            "province",
            "payment_method",
            "order_status",
            "qty",
            "gross_sales",
            "net_sales",
            "fees",
            "discounts",
            "shipping_cost",
        ],
        "dimensions": {
            "dim_date": ["order_date", "paid_at", "shipped_at", "order_completed_at", "sale_month"],
            "dim_product": ["product_name", "variant_name", "parent_sku_ref", "sku_ref"],
            "dim_customer": ["buyer_username", "receiver_name", "phone_number"],
            "dim_location": ["shipping_address", "district", "province", "postal_code", "country"],
            "dim_payment": ["payment_method", "payment_method_detail", "installment_plan"],
            "dim_status": ["order_status", "refund_return_status", "cancel_reason"],
            "dim_discount": [
                "bundle_discount_shopee",
                "bundle_discount_seller",
                "shopee_discount",
                "coin_discount",
                "payment_channel_promo",
                "discount_code",
            ],
            "dim_shipping": [
                "shipping_option",
                "shipping_method",
                "buyer_shipping_fee",
                "estimated_shipping_fee",
                "return_shipping_fee",
            ],
        },
    }


def _health_entries(sources: list[dict]) -> list[dict]:
    entries = []
    for src in sources:
        exists = src["path"].exists()
        entry = {
            "label": src["label"],
            "dbPath": str(src["path"]),
            "dbFile": src["path"].name,
            "ok": exists,
        }
        if exists:
            stat = src["path"].stat()
            entry["sizeBytes"] = stat.st_size
            entry["modifiedAt"] = datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M:%S")
        entries.append(entry)
    return entries


def health():
    entries = _health_entries(DB_SOURCES)
    traffic_entries = _health_entries(TRAFFIC_SOURCES)
    packing_entries = _health_entries(PACKING_SOURCES)
    cost_entries = _health_entries(COST_SOURCES)

    connected = [e for e in entries if e["ok"]]
    ok = len(connected) > 0
    message = (
        f"{len(connected)}/{len(entries)} database(s) connected" if entries else "No database configured"
    )

    return {
        "ok": ok,
        "databases": entries,
        "dbCount": len(entries),
        "connectedCount": len(connected),
        "traffic": traffic_entries,
        "packing": packing_entries,
        "cost": cost_entries,
        # Back-compat single-db fields, used by older UI code paths
        "dbPath": entries[0]["dbPath"] if entries else "",
        "dbFile": ", ".join(e["dbFile"] for e in connected) if connected else (entries[0]["dbFile"] if entries else ""),
        "message": message,
        "modifiedAt": connected[0].get("modifiedAt", "") if connected else "",
    }


def bootstrap():
    if not existing_sources():
        orders, duplicates, meta = [], [], {}
        schema = {"tables": [], "starSchema": star_schema(), "databases": []}
        import_log = []
    else:
        orders = load_orders()
        duplicates = duplicate_orders()
        completed_orders = {row["id"] for row in orders if row["status"] == "Completed"}
        rows_by_source: dict[str, int] = {}
        for row in orders:
            rows_by_source[row["sourceDb"]] = rows_by_source.get(row["sourceDb"], 0) + 1
        meta = {
            "orderRows": len(orders),
            "uniqueOrders": len({row["id"] for row in orders}),
            "completedOrders": len(completed_orders),
            "duplicateOrderIds": len(duplicates),
            "duplicates": duplicates,
            "rowsBySource": rows_by_source,
            "dateRange": {
                "from": min((row["date"] for row in orders if row["date"]), default=""),
                "to": max((row["date"] for row in orders if row["date"]), default=""),
            },
        }
        schema = load_schema()
        import_log = load_import_log()

    traffic = load_traffic() if existing(TRAFFIC_SOURCES) else {"hourly": [], "daily": []}
    packing = load_packing() if existing(PACKING_SOURCES) else []
    cost = load_cost() if existing(COST_SOURCES) else []

    traffic_meta = {
        "hourlyRows": len(traffic["hourly"]),
        "dailyRows": len(traffic["daily"]),
        "dateRange": {
            "from": min((r["date"] for r in traffic["daily"] if r["date"]), default=""),
            "to": max((r["date"] for r in traffic["daily"] if r["date"]), default=""),
        },
    }
    packing_meta = {
        "rows": len(packing),
        "matchedToOrders": sum(1 for r in packing if r["orderDate"]),
        "dateRange": {
            "from": min((r["date"] for r in packing if r["date"]), default=""),
            "to": max((r["date"] for r in packing if r["date"]), default=""),
        },
    }
    cost_meta = {
        "rows": len(cost),
        "totalCost": sum(r["totalCost"] for r in cost),
        "suppliers": len({r["supplier"] for r in cost}),
        "dateRange": {
            "from": min((r["date"] for r in cost if r["date"]), default=""),
            "to": max((r["date"] for r in cost if r["date"]), default=""),
        },
    }

    return {
        "health": health(),
        "orders": orders,
        "schema": schema,
        "importLog": import_log,
        "traffic": traffic,
        "packing": packing,
        "cost": cost,
        "meta": meta,
        "trafficMeta": traffic_meta,
        "packingMeta": packing_meta,
        "costMeta": cost_meta,
    }


def json_response(handler: BaseHTTPRequestHandler, payload, status: int = 200):
    body = json.dumps(payload, ensure_ascii=False, default=str).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.send_header("Cache-Control", "no-store")
    handler.send_header("Access-Control-Allow-Origin", "*")
    handler.end_headers()
    handler.wfile.write(body)


class DashboardHandler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        sys.stderr.write("%s - %s\n" % (self.address_string(), fmt % args))

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        try:
            if path == "/api/health":
                json_response(self, health())
            elif path == "/api/orders":
                json_response(self, {"orders": load_orders()})
            elif path == "/api/schema":
                json_response(self, load_schema())
            elif path == "/api/import-log":
                json_response(self, {"importLog": load_import_log(), "duplicates": duplicate_orders()})
            elif path == "/api/traffic":
                json_response(self, load_traffic())
            elif path == "/api/packing":
                json_response(self, {"packing": load_packing()})
            elif path == "/api/cost":
                json_response(self, {"cost": load_cost()})
            elif path == "/api/bootstrap":
                json_response(self, bootstrap())
            else:
                self.serve_static(path)
        except Exception as exc:
            json_response(self, {"ok": False, "error": str(exc)}, status=500)

    def serve_static(self, path: str):
        target = ROOT / ("dashboard.html" if path in {"/", ""} else path.lstrip("/"))
        try:
            target = target.resolve()
            target.relative_to(ROOT)
        except ValueError:
            self.send_error(403)
            return
        if not target.exists() or not target.is_file():
            self.send_error(404)
            return
        content_type = mimetypes.guess_type(str(target))[0] or "application/octet-stream"
        body = target.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type + ("; charset=utf-8" if content_type.startswith("text/") else ""))
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def main():
    global DB_SOURCES, TRAFFIC_SOURCES, PACKING_SOURCES, COST_SOURCES

    parser = argparse.ArgumentParser(description="Local API for Shopee SQLite dashboard")
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", default=8787, type=int)
    parser.add_argument("--open", action="store_true")
    parser.add_argument(
        "--db",
        action="append",
        default=None,
        metavar="PATH",
        help=(
            "Path to a SQLite database file (must contain a sales_data table) to include. "
            "Repeat for multiple databases, e.g. --db shop2024.db --db shop2025.db. "
            "Use label=path to control the display label, e.g. --db 2025=D:\\data\\2025.db. "
            "Overrides SHOPEE_DB_PATHS / SHOPEE_DB_PATH env vars if given."
        ),
    )
    parser.add_argument(
        "--traffic-db", action="append", default=None, metavar="PATH",
        help="Path to a traffic .db (traffic_data / traffic_summary tables). Repeatable, supports label=path.",
    )
    parser.add_argument(
        "--packing-db", action="append", default=None, metavar="PATH",
        help="Path to a packing .db (packed_orders table). Repeatable, supports label=path.",
    )
    parser.add_argument(
        "--cost-db", action="append", default=None, metavar="PATH",
        help="Path to a cost .xlsx/.db file. Repeatable, supports label=path.",
    )
    args = parser.parse_args()

    if args.db:
        DB_SOURCES = parse_db_sources(",".join(args.db))
    if args.traffic_db:
        TRAFFIC_SOURCES = parse_db_sources(",".join(args.traffic_db))
    if args.packing_db:
        PACKING_SOURCES = parse_db_sources(",".join(args.packing_db))
    if args.cost_db:
        COST_SOURCES = parse_db_sources(",".join(args.cost_db))

    server = ThreadingHTTPServer((args.host, args.port), DashboardHandler)
    url = f"http://{args.host}:{args.port}/"
    print(f"Shopee dashboard API running at {url}")
    for label, sources in (
        ("Sales", DB_SOURCES), ("Traffic", TRAFFIC_SOURCES),
        ("Packing", PACKING_SOURCES), ("Cost", COST_SOURCES),
    ):
        for src in sources:
            status = "OK" if src["path"].exists() else "MISSING"
            print(f"  [{label:7s}] [{status}] {src['label']}: {src['path']}")
    if args.open:
        webbrowser.open(url)
    server.serve_forever()


if __name__ == "__main__":
    main()