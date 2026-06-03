import sqlite3
from pathlib import Path
import openpyxl
from excel_engine import month_blocks, METRIC_ORDER, infer_month, classify_header, norm

DB_PATH = Path(__file__).resolve().parent / "amfi.db"
TEMPLATE_PATH = Path(__file__).resolve().parent / "data" / "AMFI_MOM DATA - Apr'25 to Mar26.xlsx"

def get_db_connection():
    conn = sqlite3.connect(DB_PATH, timeout=10.0)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    return conn

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS amfi_metrics (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            scheme_name TEXT,
            asset_type TEXT,
            scheme_structure TEXT,
            debt_equity TEXT,
            sales_prod_mis TEXT,
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
            UNIQUE(scheme_name, month, year, financial_year)
        )
    """)
    conn.commit()
    
    cursor.execute("SELECT COUNT(*) FROM amfi_metrics")
    if cursor.fetchone()[0] == 0:
        seed_db(conn)
    else:
        cursor.execute("SELECT COUNT(*) FROM amfi_metrics WHERE financial_year = '2026-2027'")
        if cursor.fetchone()[0] == 0:
            ensure_april_rollover(cursor, 4, 2026, "2026-2027")
            conn.commit()
    conn.close()

def ensure_april_rollover(cursor, month, year, fy):
    if month != 4:
        return
    cursor.execute("SELECT COUNT(*) FROM amfi_metrics WHERE financial_year = ? AND month = 3", (fy,))
    if cursor.fetchone()[0] > 0:
        return

    prev_fy = f"{year - 1}-{year}"
    cursor.execute("""
        SELECT scheme_name, asset_type, scheme_structure, debt_equity, sales_prod_mis,
               no_schemes, folios, funds_mobilized, redemption, net_inflow,
               aum, avg_aum, seg_portfolios, seg_aum
        FROM amfi_metrics
        WHERE month = 3 AND year = ? AND financial_year = ?
    """, (year, prev_fy))
    march_records = cursor.fetchall()
    
    if not march_records:
        cursor.execute("""
            SELECT scheme_name, asset_type, scheme_structure, debt_equity, sales_prod_mis,
                   no_schemes, folios, funds_mobilized, redemption, net_inflow,
                   aum, avg_aum, seg_portfolios, seg_aum
            FROM amfi_metrics
            WHERE month = 3 AND year = ?
        """, (year,))
        march_records = cursor.fetchall()
        
    if march_records:
        records_to_insert = [
            (
                r["scheme_name"], r["asset_type"], r["scheme_structure"], r["debt_equity"], r["sales_prod_mis"],
                r["no_schemes"], r["folios"], r["funds_mobilized"], r["redemption"], r["net_inflow"],
                r["aum"], r["avg_aum"], r["seg_portfolios"], r["seg_aum"],
                3, year, fy
            )
            for r in march_records
        ]
        cursor.executemany("""
            INSERT OR IGNORE INTO amfi_metrics (
                scheme_name, asset_type, scheme_structure, debt_equity, sales_prod_mis,
                no_schemes, folios, funds_mobilized, redemption, net_inflow,
                aum, avg_aum, seg_portfolios, seg_aum,
                month, year, financial_year
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, records_to_insert)

def seed_db(conn):
    if not TEMPLATE_PATH.exists():
        return
    
    wb = openpyxl.load_workbook(TEMPLATE_PATH, data_only=True)
    ws_flat = wb["AMFI-Mar'25 to Mar'26"]
    blocks = month_blocks(ws_flat)
    records_to_insert = []
    
    for row in range(4, ws_flat.max_row + 1):
        scheme_name = ws_flat.cell(row, 1).value
        if not scheme_name or str(scheme_name).strip().lower() in ("grand total", "total", "fund of funds", "growth") or "sub total" in str(scheme_name).strip().lower():
            continue
            
        asset_type = ws_flat.cell(row, 2).value
        scheme_structure = ws_flat.cell(row, 3).value
        debt_equity = ws_flat.cell(row, 4).value
        sales_prod_mis = ws_flat.cell(row, 5).value
        
        for block in blocks:
            month_key = block["month"]
            if month_key == "Jan'25":
                m_val = 1
                y_val = 2026
            else:
                month_info = infer_month("", month_key)
                m_val = month_info["month"]
                y_val = month_info["year"]
            
            fy = "2025-2026"
            metric_cols = {}
            for col in range(block["start"], block["end"] + 1):
                lbl = norm(ws_flat.cell(3, col).value)
                key = classify_header(lbl)
                if key:
                    metric_cols[key] = col
            
            def get_val(key):
                col = metric_cols.get(key)
                if not col:
                    return None
                val = ws_flat.cell(row, col).value
                if val is None or val == "-":
                    return None
                try:
                    return float(str(val).replace(",", "").strip())
                except ValueError:
                    return None
            
            no_schemes = get_val("no_schemes")
            folios = get_val("folios")
            funds_mobilized = get_val("funds_mobilized")
            redemption = get_val("redemption")
            net_inflow = get_val("net_inflow")
            aum = get_val("net_aum")
            avg_aum = get_val("avg_aum")
            seg_portfolios = get_val("seg_portfolios")
            seg_aum = get_val("seg_aum")
            
            if all(v is None for v in [no_schemes, folios, funds_mobilized, redemption, net_inflow, aum, avg_aum, seg_portfolios, seg_aum]):
                continue
                
            records_to_insert.append((
                str(scheme_name).strip(),
                str(asset_type).strip() if asset_type else None,
                str(scheme_structure).strip() if scheme_structure else None,
                str(debt_equity).strip() if debt_equity else None,
                str(sales_prod_mis).strip() if sales_prod_mis else None,
                no_schemes,
                folios,
                funds_mobilized,
                redemption,
                net_inflow,
                aum,
                avg_aum,
                seg_portfolios,
                seg_aum,
                m_val,
                y_val,
                fy
            ))
            
            if m_val == 3 and y_val == 2026:
                records_to_insert.append((
                    str(scheme_name).strip(),
                    str(asset_type).strip() if asset_type else None,
                    str(scheme_structure).strip() if scheme_structure else None,
                    str(debt_equity).strip() if debt_equity else None,
                    str(sales_prod_mis).strip() if sales_prod_mis else None,
                    no_schemes,
                    folios,
                    funds_mobilized,
                    redemption,
                    net_inflow,
                    aum,
                    avg_aum,
                    seg_portfolios,
                    seg_aum,
                    m_val,
                    y_val,
                    "2026-2027"
                ))
            
    cursor = conn.cursor()
    cursor.executemany("""
        INSERT OR IGNORE INTO amfi_metrics (
            scheme_name, asset_type, scheme_structure, debt_equity, sales_prod_mis,
            no_schemes, folios, funds_mobilized, redemption, net_inflow,
            aum, avg_aum, seg_portfolios, seg_aum, month, year, financial_year
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, records_to_insert)
    conn.commit()

if __name__ == "__main__":
    init_db()
    print("Database initialized successfully.")
