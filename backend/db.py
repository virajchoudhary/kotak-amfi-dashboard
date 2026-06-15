import sqlite3
from pathlib import Path

import openpyxl

from excel_engine import (
    METRIC_ORDER,
    classify_header,
    enrich_scheme_record,
    get_financial_year,
    infer_month,
    match_catalog_entry,
    metric_columns,
    month_blocks,
    norm,
    number_or_none,
    stable_scheme_key,
    template_catalog,
)

BACKEND_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = BACKEND_DIR.parent
DB_PATH = BACKEND_DIR / "amfi.db"
TEMPLATE_PATH = BACKEND_DIR / "data" / "template file.xlsx"

SCHEMA = """
    CREATE TABLE IF NOT EXISTS {table} (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        scheme_key TEXT,
        scheme_name TEXT,
        asset_type TEXT,
        scheme_structure TEXT,
        debt_equity TEXT,
        sales_prod_mis TEXT,
        fund_type TEXT,
        no_schemes REAL,
        folios REAL,
        funds_mobilized REAL,
        redemption REAL,
        net_inflow REAL,
        aum REAL,
        avg_aum REAL,
        seg_portfolios REAL,
        seg_aum REAL,
        month INTEGER,
        year INTEGER,
        financial_year TEXT,
        last_modified TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(scheme_key, month, year, financial_year)
    )
"""

INSERT_SQL = """
    INSERT OR IGNORE INTO amfi_metrics (
        scheme_key, scheme_name, asset_type, scheme_structure, debt_equity,
        sales_prod_mis, fund_type,
        no_schemes, folios, funds_mobilized, redemption, net_inflow,
        aum, avg_aum, seg_portfolios, seg_aum,
        month, year, financial_year
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
"""

def get_db_connection():
    conn = sqlite3.connect(DB_PATH, timeout=10.0)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    return conn

def init_db():
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute(SCHEMA.format(table="amfi_metrics"))
        conn.commit()
        migrate_schema_if_needed(conn)
        cleanup_template_baseline(conn)
        canonicalize_existing_records(conn)

        cursor.execute("SELECT COUNT(*) FROM amfi_metrics")
        if cursor.fetchone()[0] == 0:
            seed_db(conn)
        else:
            seed_missing_reference_rows(conn)

        cursor.execute("SELECT COUNT(*) FROM amfi_metrics WHERE financial_year = '2026-2027'")
        if cursor.fetchone()[0] == 0:
            ensure_april_rollover(cursor, 4, 2026, "2026-2027")
            conn.commit()
    finally:
        conn.close()

def migrate_schema_if_needed(conn):
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM sqlite_master WHERE type = 'table' AND name = 'amfi_metrics_old'")
    interrupted_old_table = cursor.fetchone() is not None
    cursor.execute("PRAGMA table_info(amfi_metrics)")
    columns = {row["name"] for row in cursor.fetchall()}
    cursor.execute("SELECT sql FROM sqlite_master WHERE type = 'table' AND name = 'amfi_metrics'")
    table_sql = (cursor.fetchone()["sql"] or "").lower()
    if (
        not interrupted_old_table
        and {"scheme_key", "fund_type"}.issubset(columns)
        and "unique(scheme_key" in table_sql.replace(" ", "")
    ):
        return

    source_table = "amfi_metrics_old" if interrupted_old_table else "amfi_metrics"
    rows = [dict(row) for row in cursor.execute(f"SELECT * FROM {source_table}").fetchall()]
    if interrupted_old_table:
        cursor.execute("DROP TABLE IF EXISTS amfi_metrics")
    else:
        cursor.execute("ALTER TABLE amfi_metrics RENAME TO amfi_metrics_old")
    cursor.execute(SCHEMA.format(table="amfi_metrics"))
    catalog_by_key = {entry["scheme_key"]: entry for entry in template_catalog()}

    migrated = []
    for row in rows:
        scheme_key = row.get("scheme_key") or stable_scheme_key(
            row.get("scheme_name"), row.get("asset_type"), row.get("scheme_structure")
        )
        entry = catalog_by_key.get(scheme_key)
        if not entry:
            match = enrich_scheme_record({
                "scheme": row.get("scheme_name"),
                "asset_type": row.get("asset_type"),
                "scheme_structure": row.get("scheme_structure"),
                "debt_equity": row.get("debt_equity"),
                "sales_prod_mis": row.get("sales_prod_mis"),
            })
            scheme_key = match["scheme_key"]
            entry = catalog_by_key.get(scheme_key)
        migrated.append((
            scheme_key,
            row.get("scheme_name"),
            row.get("asset_type") or (entry or {}).get("asset_type"),
            row.get("scheme_structure") or (entry or {}).get("scheme_structure"),
            row.get("debt_equity") or (entry or {}).get("debt_equity"),
            row.get("sales_prod_mis") or (entry or {}).get("sales_prod_mis"),
            row.get("fund_type") or (entry or {}).get("fund_type"),
            row.get("no_schemes"),
            row.get("folios"),
            row.get("funds_mobilized"),
            row.get("redemption"),
            row.get("net_inflow"),
            row.get("aum"),
            row.get("avg_aum"),
            row.get("seg_portfolios"),
            row.get("seg_aum"),
            row.get("month"),
            row.get("year"),
            row.get("financial_year"),
        ))
    cursor.executemany(INSERT_SQL, migrated)
    cursor.execute("DROP TABLE IF EXISTS amfi_metrics_old")
    conn.commit()

def ensure_april_rollover(cursor, month, year, fy):
    if month != 4:
        return
    cursor.execute("SELECT COUNT(*) FROM amfi_metrics WHERE financial_year = ? AND month = 3", (fy,))
    if cursor.fetchone()[0] > 0:
        return

    prev_fy = f"{year - 1}-{year}"
    cursor.execute("""
        SELECT scheme_key, scheme_name, asset_type, scheme_structure, debt_equity,
               sales_prod_mis, fund_type,
               no_schemes, folios, funds_mobilized, redemption, net_inflow,
               aum, avg_aum, seg_portfolios, seg_aum
        FROM amfi_metrics
        WHERE month = 3 AND year = ? AND financial_year = ?
    """, (year, prev_fy))
    march_records = cursor.fetchall()

    if not march_records:
        cursor.execute("""
            SELECT scheme_key, scheme_name, asset_type, scheme_structure, debt_equity,
                   sales_prod_mis, fund_type,
                   no_schemes, folios, funds_mobilized, redemption, net_inflow,
                   aum, avg_aum, seg_portfolios, seg_aum
            FROM amfi_metrics
            WHERE month = 3 AND year = ?
        """, (year,))
        march_records = cursor.fetchall()

    cursor.executemany(
        INSERT_SQL,
        [
            (
                r["scheme_key"], r["scheme_name"], r["asset_type"], r["scheme_structure"],
                r["debt_equity"], r["sales_prod_mis"], r["fund_type"],
                r["no_schemes"], r["folios"], r["funds_mobilized"], r["redemption"],
                r["net_inflow"], r["aum"], r["avg_aum"], r["seg_portfolios"], r["seg_aum"],
                3, year, fy
            )
            for r in march_records
        ],
    )

def cleanup_template_baseline(conn):
    conn.cursor().execute("""
        DELETE FROM amfi_metrics
        WHERE financial_year = '2024-2025' AND (
            (month = 3 AND year = 2025) OR (month = 1 AND year = 2025)
        )
    """)
    conn.commit()

def canonicalize_existing_records(conn):
    cursor = conn.cursor()
    rows = [dict(row) for row in cursor.execute("SELECT * FROM amfi_metrics").fetchall()]
    buckets = {}
    for row in rows:
        enriched = enrich_scheme_record({
            "scheme": row.get("scheme_name"),
            "asset_type": row.get("asset_type"),
            "scheme_structure": row.get("scheme_structure"),
            "debt_equity": row.get("debt_equity"),
            "sales_prod_mis": row.get("sales_prod_mis"),
            "fund_type": row.get("fund_type"),
        })
        matched = None
        if not row.get("scheme_key"):
            matched = match_catalog_entry(
                row.get("scheme_name"),
                row.get("asset_type"),
                row.get("scheme_structure"),
                row.get("debt_equity"),
                row.get("sales_prod_mis"),
            )
        canonical_key = row.get("scheme_key") or (matched["scheme_key"] if matched else enriched["scheme_key"])
        bucket_key = (canonical_key, row.get("month"), row.get("year"), row.get("financial_year"))
        current = buckets.get(bucket_key)
        if not current or str(row.get("last_modified") or "") >= str(current.get("last_modified") or ""):
            buckets[bucket_key] = {**row, **{
                "scheme_key": canonical_key,
                "asset_type": (matched or {}).get("asset_type") or row.get("asset_type") or enriched.get("asset_type"),
                "scheme_structure": (matched or {}).get("scheme_structure") or row.get("scheme_structure") or enriched.get("scheme_structure"),
                "debt_equity": (matched or {}).get("debt_equity") or row.get("debt_equity") or enriched.get("debt_equity"),
                "sales_prod_mis": (matched or {}).get("sales_prod_mis") or row.get("sales_prod_mis") or enriched.get("sales_prod_mis"),
                "fund_type": (matched or {}).get("fund_type") or row.get("fund_type") or enriched.get("fund_type"),
            }}

    cursor.execute("DELETE FROM amfi_metrics")
    cursor.executemany(
        """
        INSERT OR REPLACE INTO amfi_metrics (
            id, scheme_key, scheme_name, asset_type, scheme_structure, debt_equity,
            sales_prod_mis, fund_type, no_schemes, folios, funds_mobilized,
            redemption, net_inflow, aum, avg_aum, seg_portfolios, seg_aum,
            month, year, financial_year, last_modified
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        [
            (
                row.get("id"), row.get("scheme_key"), row.get("scheme_name"),
                row.get("asset_type"), row.get("scheme_structure"), row.get("debt_equity"),
                row.get("sales_prod_mis"), row.get("fund_type"), row.get("no_schemes"),
                row.get("folios"), row.get("funds_mobilized"), row.get("redemption"),
                row.get("net_inflow"), row.get("aum"), row.get("avg_aum"),
                row.get("seg_portfolios"), row.get("seg_aum"), row.get("month"),
                row.get("year"), row.get("financial_year"), row.get("last_modified"),
            )
            for row in buckets.values()
        ],
    )
    conn.commit()

def seed_db(conn):
    seed_missing_reference_rows(conn)

def seed_missing_reference_rows(conn):
    if not TEMPLATE_PATH.exists():
        return

    wb = openpyxl.load_workbook(TEMPLATE_PATH, data_only=True, keep_links=False)
    ws_flat = wb["AMFI-Mar'25 to Mar'26"]
    catalog = template_catalog()
    blocks = month_blocks(ws_flat)
    records = []

    for entry in catalog:
        row = entry["template_row"]
        for block_index, block in enumerate(blocks):
            month_key = block["month"]
            info = infer_month("", month_key)
            month = info["month"]
            year = info["year"]
            if month_key == "Jan'25":
                year = 2026
            fy = f"{year}-{year + 1}" if block_index == 0 and month == 3 else get_financial_year(month, year)
            cols = metric_columns(ws_flat, block["start"], block["end"])

            def val(metric):
                col = cols.get(metric)
                return number_or_none(ws_flat.cell(row, col).value) if col else None

            metrics = {metric: val(metric) for metric in METRIC_ORDER}
            if all(value is None for value in metrics.values()):
                continue

            base_record = (
                entry["scheme_key"], entry["scheme_name"], entry["asset_type"],
                entry["scheme_structure"], entry["debt_equity"], entry["sales_prod_mis"],
                entry["fund_type"], metrics.get("no_schemes"), metrics.get("folios"),
                metrics.get("funds_mobilized"), metrics.get("redemption"),
                metrics.get("net_inflow"), metrics.get("net_aum"), metrics.get("avg_aum"),
                metrics.get("seg_portfolios"), metrics.get("seg_aum"), month, year, fy
            )
            records.append(base_record)
            if month == 3 and year == 2026:
                records.append((*base_record[:-1], "2026-2027"))

    conn.cursor().executemany(INSERT_SQL, records)
    conn.commit()

if __name__ == "__main__":
    init_db()
    print("Database initialized successfully.")
