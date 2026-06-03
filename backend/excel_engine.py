import calendar
import io
import re
import sqlite3
from copy import copy
from datetime import datetime
from pathlib import Path
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

DB_PATH = Path(__file__).resolve().parent / "amfi.db"
TEMPLATE_PATH = Path(__file__).resolve().parent / "data" / "AMFI_MOM DATA - Apr'25 to Mar26.xlsx"
SHEET_SIP = "AMFI-SIP"
SHEET_FLAT = "AMFI-Mar'25 to Mar'26"
SHEET_FORM = "AMFI-Mar'25 to Jan'26-AMFI form"
SHEET_FLAT_PREFIX = "AMFI-Mar'25 to "
SHEET_FORM_PREFIX = "AMFI-Mar'25 to "
SHEET_FORM_SUFFIX = "-AMFI form"
HEADER_ROW = 3
MONTH_ROW = 2
DATA_START_ROW = 4

MONTHS = {
    "jan": 1, "january": 1, "feb": 2, "february": 2, "mar": 3, "march": 3,
    "apr": 4, "april": 4, "may": 5, "jun": 6, "june": 6, "jul": 7, "july": 7,
    "aug": 8, "august": 8, "sep": 9, "sept": 9, "september": 9, "oct": 10,
    "october": 10, "nov": 11, "november": 11, "dec": 12, "december": 12,
}
MONTH_ABBR = {v: k.title() for k, v in {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6, "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12}.items()}

METRIC_ORDER = [
    "no_schemes", "folios", "funds_mobilized", "redemption", "net_inflow",
    "net_aum", "avg_aum", "seg_portfolios", "seg_aum",
]

METRIC_LABELS = {
    "no_schemes": "No. of Schemes as on {date}",
    "folios": "No. of Folios as on {date}",
    "funds_mobilized": "Funds Mobilized for the month of {month} {year} (INR in crore)",
    "redemption": "Repurchase/Redemption for the month of {month} {year} (INR in crore)",
    "net_inflow": "Net Inflow (+ve)/Outflow (-ve) for the month of {month} {year} (INR in crore)",
    "net_aum": "Net Assets Under Management as on {date} (INR in crore)",
    "avg_aum": "Average Net Assets Under Management for the month {month} {year} (INR in crore)",
    "seg_portfolios": "No. of segregated portfolios created as on {date}",
    "seg_aum": "Net Assets Under Management in segregated portfolio as on {date} (INR in crore)",
}

_FORM_SUBTOTAL_ROWS = {
    22: [(6, 21)],
    36: [(25, 35)],
    45: [(39, 44)],
    50: [(48, 49)],
    57: [(53, 56)],
    59: [(6, 21), (25, 35), (39, 44), (48, 49), (53, 56)],
    67: [(63, 66)],
    72: [(70, 71)],
    76: [(63, 66), (70, 71), (74, 74)],
    85: [(79, 83)],
    87: [(59, 59), (76, 76), (85, 85)],
}

def get_db_connection():
    conn = sqlite3.connect(DB_PATH, timeout=10.0)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    return conn

def template_bytes() -> bytes:
    return TEMPLATE_PATH.read_bytes()

def get_financial_year(month: int, year: int) -> str:
    if month >= 4:
        return f"{year}-{year + 1}"
    return f"{year - 1}-{year}"

def process_upload_db(upload_bytes: bytes, filename: str) -> tuple[str, list[str]]:
    try:
        rows, month_info, warnings = parse_upload(upload_bytes, filename)
    except Exception as e:
        raise ValueError(f"Failed to parse uploaded Excel file: {str(e)}")
        
    if not rows:
        raise ValueError("Invalid sheet structure: No scheme rows or metric columns detected.")
        
    month = month_info["month"]
    year = month_info["year"]
    fy = get_financial_year(month, year)
    
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        
        # April rollover baseline creation
        if month == 4:
            cursor.execute("SELECT COUNT(*) FROM amfi_metrics WHERE financial_year = ?", (fy,))
            fy_exists = cursor.fetchone()[0] > 0
            
            if not fy_exists:
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
                
                baseline_records = [
                    (
                        r["scheme_name"], r["asset_type"], r["scheme_structure"], r["debt_equity"], r["sales_prod_mis"],
                        r["no_schemes"], r["folios"], r["funds_mobilized"], r["redemption"], r["net_inflow"],
                        r["aum"], r["avg_aum"], r["seg_portfolios"], r["seg_aum"],
                        3, year, fy
                    )
                    for r in march_records
                ]
                
                if baseline_records:
                    cursor.executemany("""
                        INSERT OR IGNORE INTO amfi_metrics (
                            scheme_name, asset_type, scheme_structure, debt_equity, sales_prod_mis,
                            no_schemes, folios, funds_mobilized, redemption, net_inflow,
                            aum, avg_aum, seg_portfolios, seg_aum,
                            month, year, financial_year
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, baseline_records)
        
        insert_records = [
            (
                record["scheme"],
                record.get("asset_type"),
                record.get("scheme_structure"),
                record.get("debt_equity"),
                record.get("sales_prod_mis"),
                record["metrics"].get("no_schemes"),
                record["metrics"].get("folios"),
                record["metrics"].get("funds_mobilized"),
                record["metrics"].get("redemption"),
                record["metrics"].get("net_inflow"),
                record["metrics"].get("net_aum"),
                record["metrics"].get("avg_aum"),
                record["metrics"].get("seg_portfolios"),
                record["metrics"].get("seg_aum"),
                month,
                year,
                fy
            )
            for record in rows
        ]
            
        cursor.executemany("""
            INSERT OR REPLACE INTO amfi_metrics (
                scheme_name, asset_type, scheme_structure, debt_equity, sales_prod_mis,
                no_schemes, folios, funds_mobilized, redemption, net_inflow,
                aum, avg_aum, seg_portfolios, seg_aum,
                month, year, financial_year, last_modified
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
        """, insert_records)
        
        conn.commit()
        return month_info["key"], warnings
    except Exception as e:
        conn.rollback()
        raise ValueError(f"Database error during ingestion: {str(e)}")
    finally:
        conn.close()

def extract_column_styles(ws, start_col: int, count: int) -> tuple[list[float], list[list[dict]]]:
    widths = []
    styles = []
    for col_offset in range(count):
        col = start_col + col_offset
        widths.append(ws.column_dimensions[get_column_letter(col)].width or 12.0)
        col_styles = []
        for row in range(1, ws.max_row + 1):
            cell = ws.cell(row, col)
            col_styles.append({
                "font": copy(cell.font) if cell.font else None,
                "fill": copy(cell.fill) if cell.fill else None,
                "border": copy(cell.border) if cell.border else None,
                "alignment": copy(cell.alignment) if cell.alignment else None,
                "number_format": cell.number_format,
                "value": cell.value if row <= 3 else None
            })
        styles.append(col_styles)
    return widths, styles

def apply_column_style(ws, col: int, width: float, styles: list[dict]):
    ws.column_dimensions[get_column_letter(col)].width = width
    for row_idx, style_info in enumerate(styles, 1):
        cell = ws.cell(row_idx, col)
        if style_info["font"]: cell.font = copy(style_info["font"])
        if style_info["fill"]: cell.fill = copy(style_info["fill"])
        if style_info["border"]: cell.border = copy(style_info["border"])
        if style_info["alignment"]: cell.alignment = copy(style_info["alignment"])
        cell.number_format = style_info["number_format"]
        if style_info["value"] is not None:
            cell.value = style_info["value"]

def write_form_column_subtotals(ws, col: int) -> None:
    col_letter = get_column_letter(col)
    for summary_row, row_ranges in _FORM_SUBTOTAL_ROWS.items():
        parts = [f"SUM({col_letter}{r1}:{col_letter}{r2})" for r1, r2 in row_ranges]
        ws.cell(summary_row, col).value = "=" + " + ".join(parts)

def get_month_seq(month: int, year: int, start_year: int) -> int:
    if month == 3 and year == start_year:
        return 0
    elif month >= 4 and year == start_year:
        return month - 3
    elif month <= 3 and year == start_year + 1:
        return month + 9
    return (year - start_year) * 12 + month

def clean_merged_ranges(ws, start_col: int):
    for r in list(ws.merged_cells.ranges):
        if r.min_col >= start_col or r.max_col >= start_col:
            ws.merged_cells.remove(r)

def compile_excel_for_fy(fy: str) -> bytes:
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM amfi_metrics WHERE financial_year = ?", (fy,))
    records = [dict(r) for r in cursor.fetchall()]
    
    if not records:
        try:
            start_year = int(fy.split("-")[0])
        except Exception:
            start_year = 2026
        prev_fy = f"{start_year - 1}-{start_year}"
        cursor.execute("""
            SELECT scheme_name, asset_type, scheme_structure, debt_equity, sales_prod_mis,
                   no_schemes, folios, funds_mobilized, redemption, net_inflow,
                   aum, avg_aum, seg_portfolios, seg_aum
            FROM amfi_metrics
            WHERE month = 3 AND year = ? AND financial_year = ?
        """, (start_year, prev_fy))
        march_records = cursor.fetchall()
        if march_records:
            baseline_records = [
                (
                    r["scheme_name"], r["asset_type"], r["scheme_structure"], r["debt_equity"], r["sales_prod_mis"],
                    r["no_schemes"], r["folios"], r["funds_mobilized"], r["redemption"], r["net_inflow"],
                    r["aum"], r["avg_aum"], r["seg_portfolios"], r["seg_aum"],
                    3, start_year, fy
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
            """, baseline_records)
            conn.commit()
            
            cursor.execute("SELECT * FROM amfi_metrics WHERE financial_year = ?", (fy,))
            records = [dict(r) for r in cursor.fetchall()]
            
    conn.close()
    
    if not records:
        raise ValueError(f"No records found in database for fiscal year: {fy}")
        
    try:
        start_year = int(fy.split("-")[0])
    except Exception:
        start_year = 2025
        
    records_by_seq = {}
    for r in records:
        seq = get_month_seq(r["month"], r["year"], start_year)
        records_by_seq.setdefault(seq, []).append(r)
        
    sorted_seqs = sorted(records_by_seq.keys())
    
    wb = load_workbook(TEMPLATE_PATH)
    ws_flat = wb[SHEET_FLAT] if SHEET_FLAT in wb.sheetnames else wb.worksheets[1]
    ws_form = wb[SHEET_FORM] if SHEET_FORM in wb.sheetnames else wb.worksheets[2]
    
    flat_march_widths, flat_march_styles = extract_column_styles(ws_flat, 6, 2)
    flat_sep_widths, flat_sep_styles = extract_column_styles(ws_flat, 8, 1)
    flat_metric_widths, flat_metric_styles = extract_column_styles(ws_flat, 9, 9)
    
    cum_idx = 142 if ws_flat.max_column >= 143 else 19
    flat_cum_widths, flat_cum_styles = extract_column_styles(ws_flat, cum_idx, 2)
    
    grw_idx = 145 if ws_flat.max_column >= 147 else 22
    flat_growth_widths, flat_growth_styles = extract_column_styles(ws_flat, grw_idx, 3)
    
    form_march_widths, form_march_styles = extract_column_styles(ws_form, 3, 2)
    form_sep_widths, form_sep_styles = extract_column_styles(ws_form, 5, 1)
    form_metric_widths, form_metric_styles = extract_column_styles(ws_form, 6, 11)
    
    ws_flat.delete_cols(6, ws_flat.max_column - 5)
    clean_merged_ranges(ws_flat, 6)
    
    ws_form.delete_cols(3, ws_form.max_column - 2)
    clean_merged_ranges(ws_form, 3)
    
    flat_months_cols = []
    latest_month_key = ""
    prev_month_key = ""
    
    for seq in sorted_seqs:
        group = records_by_seq[seq]
        m_val = group[0]["month"]
        y_val = group[0]["year"]
        month_abbr = MONTH_ABBR[m_val]
        month_key = f"{month_abbr}'{str(y_val)[-2:]}"
        
        if seq == 0:
            latest_month_key = month_key
            col_start = ws_flat.max_column + 1
            for offset in range(2):
                apply_column_style(ws_flat, col_start + offset, flat_march_widths[offset], flat_march_styles[offset])
            ws_flat.cell(MONTH_ROW, col_start).value = month_key
            ws_flat.cell(HEADER_ROW, col_start).value = "Net Assets Under Management as on " + datetime(y_val, m_val, 31).strftime("%B %d, %Y") + " (INR in crore)"
            ws_flat.cell(HEADER_ROW, col_start + 1).value = "Average Net Assets Under Management for the month " + datetime(y_val, m_val, 1).strftime("%B %Y") + " (INR in crore)"
            ws_flat.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_start + 1)
            ws_flat.merge_cells(start_row=2, start_column=col_start, end_row=2, end_column=col_start + 1)
            
            db_map = {norm_key(r["scheme_name"]): r for r in group}
            for row in range(DATA_START_ROW, ws_flat.max_row + 1):
                s_name = ws_flat.cell(row, 1).value
                if s_name:
                    r_db = db_map.get(norm_key(s_name))
                    if r_db:
                        ws_flat.cell(row, col_start).value = r_db["aum"]
                        ws_flat.cell(row, col_start + 1).value = r_db["avg_aum"]
                        
            col_start_form = ws_form.max_column + 1
            for offset in range(2):
                apply_column_style(ws_form, col_start_form + offset, form_march_widths[offset], form_march_styles[offset])
            ws_form.cell(MONTH_ROW, col_start_form).value = month_key
            ws_form.cell(HEADER_ROW, col_start_form).value = "Net Assets Under Management as on " + datetime(y_val, m_val, 31).strftime("%B %d, %Y") + " (INR in crore)"
            ws_form.cell(HEADER_ROW, col_start_form + 1).value = "Average Net Assets Under Management for the month " + datetime(y_val, m_val, 1).strftime("%B %Y") + " (INR in crore)"
            ws_form.merge_cells(start_row=2, start_column=col_start_form, end_row=2, end_column=col_start_form + 1)
            
            for row in range(DATA_START_ROW, ws_form.max_row + 1):
                if row in _FORM_SUBTOTAL_ROWS:
                    write_form_column_subtotals(ws_form, col_start_form)
                    write_form_column_subtotals(ws_form, col_start_form + 1)
                else:
                    s_name = ws_form.cell(row, 2).value
                    if s_name and not is_aggregate_scheme(s_name):
                        r_db = db_map.get(norm_key(s_name))
                        if r_db:
                            ws_form.cell(row, col_start_form).value = r_db["aum"]
                            ws_form.cell(row, col_start_form + 1).value = r_db["avg_aum"]
                            
        else:
            prev_month_key = latest_month_key
            latest_month_key = month_key
            
            flat_sep = ws_flat.max_column + 1
            apply_column_style(ws_flat, flat_sep, flat_sep_widths[0], flat_sep_styles[0])
            ws_flat.cell(MONTH_ROW, flat_sep).value = "-"
            ws_flat.cell(HEADER_ROW, flat_sep).value = "-"
            
            form_sep = ws_form.max_column + 1
            apply_column_style(ws_form, form_sep, form_sep_widths[0], form_sep_styles[0])
            ws_form.cell(MONTH_ROW, form_sep).value = "-"
            ws_form.cell(HEADER_ROW, form_sep).value = "-"
            
            flat_start = ws_flat.max_column + 1
            flat_months_cols.append(flat_start)
            for offset in range(9):
                apply_column_style(ws_flat, flat_start + offset, flat_metric_widths[offset], flat_metric_styles[offset])
            
            ws_flat.cell(MONTH_ROW, flat_start).value = month_key
            m_info = month_info(m_val, y_val)
            for offset, metric in enumerate(METRIC_ORDER):
                ws_flat.cell(HEADER_ROW, flat_start + offset).value = metric_label(metric, m_info)
                
            ws_flat.merge_cells(start_row=1, start_column=flat_start, end_row=1, end_column=flat_start + 8)
            ws_flat.merge_cells(start_row=2, start_column=flat_start, end_row=2, end_column=flat_start + 8)
            
            db_map = {norm_key(r["scheme_name"]): r for r in group}
            for row in range(DATA_START_ROW, ws_flat.max_row + 1):
                s_name = ws_flat.cell(row, 1).value
                if s_name:
                    r_db = db_map.get(norm_key(s_name))
                    if r_db:
                        metrics_map = {
                            "no_schemes": r_db["no_schemes"],
                            "folios": r_db["folios"],
                            "funds_mobilized": r_db["funds_mobilized"],
                            "redemption": r_db["redemption"],
                            "net_inflow": r_db["net_inflow"],
                            "net_aum": r_db["aum"],
                            "avg_aum": r_db["avg_aum"],
                            "seg_portfolios": r_db["seg_portfolios"],
                            "seg_aum": r_db["seg_aum"]
                        }
                        for offset, metric in enumerate(METRIC_ORDER):
                            ws_flat.cell(row, flat_start + offset).value = metrics_map.get(metric)
                            
            form_start = ws_form.max_column + 1
            for offset in range(11):
                apply_column_style(ws_form, form_start + offset, form_metric_widths[offset], form_metric_styles[offset])
                
            ws_form.cell(MONTH_ROW, form_start).value = f"Monthly Report for the month of {m_info['full_month']} {y_val} "
            ws_form.cell(HEADER_ROW, form_start).value = "Sr "
            ws_form.cell(HEADER_ROW, form_start + 1).value = "Scheme Name "
            for offset, metric in enumerate(METRIC_ORDER):
                ws_form.cell(HEADER_ROW, form_start + 2 + offset).value = metric_label(metric, m_info)
                
            ws_form.merge_cells(start_row=1, start_column=form_start, end_row=1, end_column=form_start + 10)
            ws_form.merge_cells(start_row=2, start_column=form_start, end_row=2, end_column=form_start + 10)
            
            for row in range(DATA_START_ROW, ws_form.max_row + 1):
                if row in _FORM_SUBTOTAL_ROWS:
                    for offset in range(9):
                        write_form_column_subtotals(ws_form, form_start + 2 + offset)
                else:
                    s_name = ws_form.cell(row, 2).value
                    if s_name and not is_aggregate_scheme(s_name):
                        r_db = db_map.get(norm_key(s_name))
                        ws_form.cell(row, form_start).value = ws_form.cell(row, 1).value
                        ws_form.cell(row, form_start + 1).value = ws_form.cell(row, 2).value
                        if r_db:
                            metrics_map = {
                                "no_schemes": r_db["no_schemes"],
                                "folios": r_db["folios"],
                                "funds_mobilized": r_db["funds_mobilized"],
                                "redemption": r_db["redemption"],
                                "net_inflow": r_db["net_inflow"],
                                "net_aum": r_db["aum"],
                                "avg_aum": r_db["avg_aum"],
                                "seg_portfolios": r_db["seg_portfolios"],
                                "seg_aum": r_db["seg_aum"]
                            }
                            for offset, metric in enumerate(METRIC_ORDER):
                                ws_form.cell(row, form_start + 2 + offset).value = metrics_map.get(metric)

    if flat_months_cols:
        flat_sep = ws_flat.max_column + 1
        apply_column_style(ws_flat, flat_sep, flat_sep_widths[0], flat_sep_styles[0])
        ws_flat.cell(MONTH_ROW, flat_sep).value = "-"
        ws_flat.cell(HEADER_ROW, flat_sep).value = "-"
        
        cum_start = ws_flat.max_column + 1
        for offset in range(2):
            apply_column_style(ws_flat, cum_start + offset, flat_cum_widths[offset], flat_cum_styles[offset])
            
        ws_flat.cell(MONTH_ROW, cum_start).value = f"Apr'{str(start_year)[-2:]} to {latest_month_key}"
        ws_flat.cell(HEADER_ROW, cum_start).value = "Funds Mobilized  (INR in crore)"
        ws_flat.cell(HEADER_ROW, cum_start + 1).value = "Net Inflow (+ve)/Outflow (-ve)  (INR in crore)"
        ws_flat.merge_cells(start_row=1, start_column=cum_start, end_row=1, end_column=cum_start + 1)
        ws_flat.merge_cells(start_row=2, start_column=cum_start, end_row=2, end_column=cum_start + 1)
        
        funds_refs = ",".join(f"{get_column_letter(c + 2)}{{row}}" for c in flat_months_cols)
        net_refs = ",".join(f"{get_column_letter(c + 4)}{{row}}" for c in flat_months_cols)
        for row in range(DATA_START_ROW, ws_flat.max_row + 1):
            ws_flat.cell(row, cum_start).value = f"=SUM({funds_refs.format(row=row)})"
            ws_flat.cell(row, cum_start + 1).value = f"=SUM({net_refs.format(row=row)})"
            
        flat_sep2 = ws_flat.max_column + 1
        apply_column_style(ws_flat, flat_sep2, flat_sep_widths[0], flat_sep_styles[0])
        ws_flat.cell(MONTH_ROW, flat_sep2).value = "-"
        ws_flat.cell(HEADER_ROW, flat_sep2).value = "-"
        
        grw_start = ws_flat.max_column + 1
        for offset in range(3):
            apply_column_style(ws_flat, grw_start + offset, flat_growth_widths[offset], flat_growth_styles[offset])
            
        ws_flat.cell(MONTH_ROW, grw_start).value = "-"
        ws_flat.cell(HEADER_ROW, grw_start).value = f"GS - Growth-({prev_month_key} Vs {latest_month_key})"
        ws_flat.cell(HEADER_ROW, grw_start + 1).value = f"NS - Growth-({prev_month_key} Vs {latest_month_key})"
        ws_flat.cell(HEADER_ROW, grw_start + 2).value = f"AUM - Growth-({prev_month_key} Vs {latest_month_key})"
        
        latest_idx = flat_months_cols[-1]
        prev_idx = flat_months_cols[-2] if len(flat_months_cols) >= 2 else 6
        
        for row in range(DATA_START_ROW, ws_flat.max_row + 1):
            if prev_idx == 6:
                ws_flat.cell(row, grw_start).value = "=0"
            else:
                ws_flat.cell(row, grw_start).value = growth_formula(prev_idx + 2, latest_idx + 2, row)
                
            if prev_idx == 6:
                ws_flat.cell(row, grw_start + 1).value = "=0"
            else:
                ws_flat.cell(row, grw_start + 1).value = growth_formula(prev_idx + 4, latest_idx + 4, row)
                
            prev_aum_col = 6 if prev_idx == 6 else prev_idx + 5
            ws_flat.cell(row, grw_start + 2).value = growth_formula(prev_aum_col, latest_idx + 5, row)

    baseline_month_key = f"Mar'{str(start_year)[-2:]}"
    ws_flat.title = f"AMFI-{baseline_month_key} to {latest_month_key}"
    ws_form.title = f"AMFI-{baseline_month_key} to {latest_month_key}-AMFI form"
    wb.active = wb.worksheets.index(ws_flat)
    
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

def rename_month_sheets(flat, form, latest_month: str) -> None:
    flat.title = f"{SHEET_FLAT_PREFIX}{latest_month}"
    form.title = f"{SHEET_FORM_PREFIX}{latest_month}{SHEET_FORM_SUFFIX}"

def growth_formula(previous_col: int | None, current_col: int | None, row: int) -> str:
    if not previous_col or not current_col or previous_col == current_col:
        return "=0"
    prev_col = get_column_letter(previous_col)
    curr_col = get_column_letter(current_col)
    return f"=IF({prev_col}{row}=0, 0, ({curr_col}{row}-{prev_col}{row})/{prev_col}{row})"

def build_summary_db(fy: str, wb) -> dict:
    conn = get_db_connection()
    cursor = conn.cursor()
    
    try:
        start_year = int(fy.split("-")[0])
    except Exception:
        start_year = 2025
        
    cursor.execute("""
        SELECT month, year,
               SUM(funds_mobilized) as funds_mobilized,
               SUM(redemption) as redemption,
               SUM(net_inflow) as net_inflow,
               SUM(aum) as net_aum,
               SUM(avg_aum) as avg_aum
        FROM amfi_metrics
        WHERE financial_year = ?
        GROUP BY month, year
    """, (fy,))
    
    rows = [dict(r) for r in cursor.fetchall()]
    conn.close()
    
    rows.sort(key=lambda r: get_month_seq(r["month"], r["year"], start_year))
    
    series = []
    for r in rows:
        abbr = MONTH_ABBR[r["month"]]
        month_key = f"{abbr}'{str(r['year'])[-2:]}"
        series.append({
            "month": month_key,
            "funds_mobilized": round(r["funds_mobilized"] or 0.0, 2),
            "redemption": round(r["redemption"] or 0.0, 2),
            "net_inflow": round(r["net_inflow"] or 0.0, 2),
            "net_aum": round(r["net_aum"] or 0.0, 2),
            "avg_aum": round(r["avg_aum"] or 0.0, 2),
        })
        
    latest = series[-1] if series else {}
    return {
        "sheetCount": len(wb.sheetnames),
        "timeSeries": series,
        "latestMonth": latest.get("month"),
        "latestFundsMobilized": latest.get("funds_mobilized", 0),
        "latestNetInflow": latest.get("net_inflow", 0),
        "latestNetAum": latest.get("net_aum", 0),
    }

def dashboard_payload(workbook_bytes: bytes, warnings: list[str] | None = None, upload_month: str | None = None, fy: str | None = None) -> dict:
    wb = load_workbook(io.BytesIO(workbook_bytes), data_only=False)
    sheets = sheet_payloads(workbook_bytes, wb.sheetnames)
    return {
        "sheets": sheets,
        "summary": build_summary(wb, fy=fy),
        "warnings": warnings or [],
        "uploadMonth": upload_month,
        "generatedAt": datetime.utcnow().isoformat(timespec="seconds") + "Z",
    }

def sheet_payloads(workbook_bytes: bytes, sheet_names: list[str]) -> dict:
    sheets = {}
    for sheet_name in sheet_names:
        df = pd.read_excel(io.BytesIO(workbook_bytes), sheet_name=sheet_name, header=None).fillna("")
        rows = df.values.tolist()
        max_col = len(rows[0]) if rows else 0
        sheets[sheet_name] = {
            "name": sheet_name,
            "maxRow": len(rows),
            "maxColumn": max_col,
            "columns": rows[HEADER_ROW - 1] if len(rows) >= HEADER_ROW else [get_column_letter(c) for c in range(1, max_col + 1)],
            "rows": rows,
        }
    return sheets

_SKIP_LABELS = ("grand total", "total", "fund of funds", "growth")

def _is_summary_row(ws, row: int) -> bool:
    value = norm_key(ws.cell(row, 1).value)
    if not value:
        return True
    return any(label in value for label in _SKIP_LABELS)

def build_summary(wb, fy: str | None = None) -> dict:
    if fy:
        try:
            return build_summary_db(fy, wb)
        except Exception:
            pass

    flat = flat_sheet(wb)
    blocks = month_blocks(flat)
    metric_keys = ("funds_mobilized", "redemption", "net_inflow", "net_aum", "avg_aum")
    series = []
    for block in blocks:
        cols = metric_columns(flat, block["start"], block["end"])
        sums = {k: 0.0 for k in metric_keys}
        for row in range(DATA_START_ROW, flat.max_row + 1):
            if _is_summary_row(flat, row):
                continue
            for key in metric_keys:
                col = cols.get(key)
                if col:
                    val = flat.cell(row, col).value
                    if isinstance(val, (int, float)):
                        sums[key] += float(val)
        sums = {k: round(v, 2) for k, v in sums.items()}
        series.append({"month": block["month"], **sums})

    latest = series[-1] if series else {}
    return {
        "sheetCount": len(wb.sheetnames),
        "timeSeries": series,
        "latestMonth": latest.get("month"),
        "latestFundsMobilized": latest.get("funds_mobilized", 0),
        "latestNetInflow": latest.get("net_inflow", 0),
        "latestNetAum": latest.get("net_aum", 0),
    }

def flat_sheet(wb):
    return sheet_by_name(wb, SHEET_FLAT, "AMFI-Mar'", exclude_suffix=SHEET_FORM_SUFFIX)

def form_sheet(wb):
    return sheet_by_name(wb, SHEET_FORM, "AMFI-Mar'", suffix=SHEET_FORM_SUFFIX)

def sheet_by_name(wb, exact: str, prefix: str, suffix: str | None = None, exclude_suffix: str | None = None):
    if exact in wb.sheetnames:
        return wb[exact]
    for name in wb.sheetnames:
        if not name.startswith(prefix):
            continue
        if suffix and not name.endswith(suffix):
            continue
        if exclude_suffix and name.endswith(exclude_suffix):
            continue
        return wb[name]
    raise KeyError(exact)

def parse_upload(upload_bytes: bytes, filename: str) -> tuple[list[dict], dict, list[str]]:
    wb = read_upload_workbook(upload_bytes)
    candidates = []
    for ws in wb.worksheets:
        header_row, columns = detect_header(ws)
        if header_row and "scheme" in columns and len([k for k in columns if k in METRIC_ORDER]) >= 3:
            candidates.append((ws, header_row, columns))
    if not candidates:
        return [], infer_month(filename, ""), ["No upload sheet contained a recognizable AMFI metric header row."]

    ws, header_row, columns = max(candidates, key=lambda item: item[0].max_row * len(item[2]))
    header_text = " ".join(str(ws.cell(header_row, c).value or "") for c in range(1, ws.max_column + 1))
    month_info = infer_month(filename, header_text)
    rows = extract_scheme_rows(ws, header_row, columns)
    return rows, month_info, []

def read_upload_workbook(upload_bytes: bytes):
    file_bytes = io.BytesIO(upload_bytes)
    try:
        frames = pd.read_excel(file_bytes, sheet_name=None, header=None, engine="openpyxl")
    except Exception:
        file_bytes.seek(0)
        try:
            frames = pd.read_excel(file_bytes, sheet_name=None, header=None, engine="xlrd")
        except Exception:
            file_bytes.seek(0)
            frames = {"Upload": pd.read_html(file_bytes)[0]}
    return frames_to_workbook(frames)

def frames_to_workbook(frames: dict[str, pd.DataFrame]):
    wb = Workbook()
    wb.remove(wb.active)
    for name, df in frames.items():
        ws = wb.create_sheet(str(name)[:31])
        for row in df.fillna("").values.tolist():
            ws.append(row)
    return wb

def detect_header(ws) -> tuple[int | None, dict]:
    for row_idx in range(1, min(ws.max_row, 50) + 1):
        columns = {}
        for col_idx in range(1, ws.max_column + 1):
            label = norm(ws.cell(row_idx, col_idx).value)
            key = classify_header(label)
            if key and key not in columns:
                columns[key] = col_idx
        if "scheme" in columns:
            return row_idx, columns
    return None, {}

def classify_header(label: str) -> str | None:
    if not label:
        return None
    if "scheme name" in label:
        return "scheme"
    if label == "asset type":
        return "asset_type"
    if "open ended" in label or "closed ended" in label:
        return "scheme_structure"
    if "debt / equity" in label or "debt equity" in label:
        return "debt_equity"
    if "sales_prod_mis" in label or "sales prod mis" in label:
        return "sales_prod_mis"
    if "no. of schemes" in label or "number of schemes" in label:
        return "no_schemes"
    if "no. of folios" in label or "number of folios" in label:
        return "folios"
    if "funds mobilized" in label:
        return "funds_mobilized"
    if "repurchase" in label or "redemption" in label:
        return "redemption"
    if "net inflow" in label or "outflow" in label:
        return "net_inflow"
    if "average net assets" in label or "average aum" in label:
        return "avg_aum"
    if "segregated portfolios created" in label:
        return "seg_portfolios"
    if "segregated portfolio" in label and "net assets" in label:
        return "seg_aum"
    if "net assets under management" in label or "net aum" in label:
        return "net_aum"
    return None

def extract_scheme_rows(ws, header_row: int, columns: dict) -> list[dict]:
    rows = []
    parent = None
    scheme_col = columns["scheme"]
    for row_idx in range(header_row + 1, ws.max_row + 1):
        scheme = clean_scheme(ws.cell(row_idx, scheme_col).value)
        if not scheme:
            continue
        metrics = {key: number_or_none(ws.cell(row_idx, columns[key]).value) for key in METRIC_ORDER if key in columns}
        has_metrics = any(value is not None for value in metrics.values())
        if not has_metrics:
            parent = scheme
            continue
        if is_aggregate_scheme(scheme):
            continue
        record = {
            "row": row_idx,
            "scheme": scheme,
            "parent": parent,
            "asset_type": cell_value(ws, row_idx, columns.get("asset_type")) or parent,
            "scheme_structure": cell_value(ws, row_idx, columns.get("scheme_structure")),
            "debt_equity": cell_value(ws, row_idx, columns.get("debt_equity")),
            "sales_prod_mis": cell_value(ws, row_idx, columns.get("sales_prod_mis")),
            "metrics": metrics,
        }
        rows.append(record)
    return rows

def month_blocks(ws) -> list[dict]:
    blocks = []
    for col in range(1, ws.max_column + 1):
        value = ws.cell(MONTH_ROW, col).value
        if is_month_key(value) and has_metric_header(ws, col):
            end = col
            while end + 1 <= ws.max_column and ws.cell(HEADER_ROW, end + 1).value != "-":
                end += 1
            sep = col - 1 if col > 1 and ws.cell(HEADER_ROW, col - 1).value == "-" else None
            limit = 2 if value.startswith("Mar") and col == 6 else len(METRIC_ORDER)
            blocks.append({"month": value, "start": col, "end": min(end, col + limit - 1), "separator": sep})
    return blocks

def form_blocks(ws) -> list[dict]:
    blocks = []
    for col in range(1, ws.max_column + 1):
        value = str(ws.cell(MONTH_ROW, col).value or "")
        if value.lower().startswith("monthly report"):
            sep = col - 1 if col > 1 and ws.cell(HEADER_ROW, col - 1).value == "-" else None
            blocks.append({"start": col, "end": col + len(METRIC_ORDER) + 1, "separator": sep})
    return blocks

def has_metric_header(ws, col: int) -> bool:
    window = " ".join(norm(ws.cell(HEADER_ROW, c).value) for c in range(col, min(ws.max_column, col + 9) + 1))
    return "funds mobilized" in window or "no. of schemes" in window or "net assets under" in window

def metric_columns(ws, start_col: int, end_col: int) -> dict:
    columns = {}
    for col in range(start_col, min(ws.max_column, end_col) + 1):
        key = classify_header(norm(ws.cell(HEADER_ROW, col).value))
        if key in METRIC_ORDER and key not in columns:
            columns[key] = col
    return columns

def infer_month(filename: str, text: str) -> dict:
    patterns = [
        r"month\s+of\s+([A-Za-z]+)\s+(20\d{2})",
        r"as\s+on\s+([A-Za-z]+)\s+\d{1,2},?\s+(20\d{2})",
        r"([A-Za-z]{3,9})[' -]*(\d{2,4})",
    ]
    for source in (filename, f"{filename} {text}"):
        for pattern in patterns:
            match = re.search(pattern, source, flags=re.IGNORECASE)
            if not match:
                continue
            month_raw, year_raw = match.group(1).lower(), match.group(2)
            month_num = MONTHS.get(month_raw[:3]) or MONTHS.get(month_raw)
            if month_num:
                year = int(year_raw) if len(year_raw) == 4 else 2000 + int(year_raw)
                return month_info(month_num, year)
    now = datetime.utcnow()
    return month_info(now.month, now.year)

def month_info(month_num: int, year: int) -> dict:
    abbr = MONTH_ABBR[month_num]
    last_day = calendar.monthrange(year, month_num)[1]
    full_month = datetime(year, month_num, 1).strftime("%B")
    return {
        "month": month_num,
        "year": year,
        "key": f"{abbr}'{str(year)[-2:]}",
        "full_month": full_month,
        "date": f"{full_month} {last_day}, {year}",
    }

def metric_label(metric: str, month_info: dict) -> str:
    return METRIC_LABELS[metric].format(date=month_info["date"], month=month_info["full_month"], year=month_info["year"])

def is_month_key(value) -> bool:
    return isinstance(value, str) and bool(re.fullmatch(r"[A-Z][a-z]{2}'\d{2}", value.strip()))

def norm(value) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip().lower().replace("/", " / "))

def norm_key(value) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(value or "").strip().lower()).strip()

def clean_scheme(value) -> str:
    text = str(value or "").strip()
    return re.sub(r"\s+", " ", text)

def is_aggregate_scheme(value) -> bool:
    key = norm_key(value)
    return key.startswith("sub total") or key.startswith("subtotal") or key.startswith("total ") or key == "grand total"

def number_or_none(value):
    if value is None or value == "-":
        return None
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return value
    cleaned = str(value).replace(",", "").strip()
    try:
        return float(cleaned)
    except ValueError:
        return None

def cell_value(ws, row: int, col: int | None):
    if not col:
        return None
    return ws.cell(row, col).value
