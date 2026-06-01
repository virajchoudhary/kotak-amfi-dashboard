import calendar
import io
import re
from copy import copy
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter


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
    "jan": 1,
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}
MONTH_ABBR = {v: k.title() for k, v in {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6, "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12}.items()}

METRIC_ORDER = [
    "no_schemes",
    "folios",
    "funds_mobilized",
    "redemption",
    "net_inflow",
    "net_aum",
    "avg_aum",
    "seg_portfolios",
    "seg_aum",
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

# Sub-total / total rows in the AMFI form sheet and the row-ranges each one
# aggregates.  The pattern is identical across every monthly block — only the
# column letter changes.  Extracted from the company template.
_FORM_SUBTOTAL_ROWS = {
    22: [(6, 21)],                                           # Sub Total - I
    36: [(25, 35)],                                          # Sub Total - II
    45: [(39, 44)],                                          # Sub Total - III
    50: [(48, 49)],                                          # Sub Total - IV
    57: [(53, 56)],                                          # Sub Total - V
    59: [(6, 21), (25, 35), (39, 44), (48, 49), (53, 56)],   # Total A — Open ended
    67: [(63, 66)],                                          # Sub Total (close-ended)
    72: [(70, 71)],                                          # Sub Total (close-ended)
    76: [(63, 66), (70, 71), (74, 74)],                      # Total B — Close ended
    85: [(79, 83)],                                          # Total C — Interval
    87: [(59, 59), (76, 76), (85, 85)],                      # Grand Total
}



def template_bytes() -> bytes:
    return TEMPLATE_PATH.read_bytes()


def process_upload(master_bytes: bytes, upload_bytes: bytes, filename: str) -> tuple[bytes, dict]:
    master = load_workbook(io.BytesIO(master_bytes))
    rows, month_info, warnings = parse_upload(upload_bytes, filename)
    if not rows:
        raise ValueError("No Tier-3 scheme rows with AMFI metrics were found in the uploaded workbook.")

    flat = flat_sheet(master)
    form = form_sheet(master)
    flat_block = ensure_flat_month(flat, month_info)
    form_block = ensure_form_month(form, month_info)

    flat_warnings = update_flat_sheet(flat, rows, flat_block)
    form_warnings = update_form_sheet(form, rows, form_block)
    warnings.extend(flat_warnings)
    warnings.extend(form_warnings)
    rename_month_sheets(flat, form, month_info["key"])
    set_active_sheet(master, flat, flat_block["start"], MONTH_ROW)

    out = io.BytesIO()
    master.save(out)
    updated_bytes = out.getvalue()
    return updated_bytes, dashboard_payload(updated_bytes, warnings=warnings, upload_month=month_info["key"])


def dashboard_payload(workbook_bytes: bytes, warnings: list[str] | None = None, upload_month: str | None = None) -> dict:
    wb = load_workbook(io.BytesIO(workbook_bytes), data_only=False)
    sheets = sheet_payloads(workbook_bytes, wb.sheetnames)
    return {
        "sheets": sheets,
        "summary": build_summary(wb),
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


def build_summary(wb) -> dict:
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
    return sheet_by_name(wb, SHEET_FLAT, SHEET_FLAT_PREFIX, exclude_suffix=SHEET_FORM_SUFFIX)


def form_sheet(wb):
    return sheet_by_name(wb, SHEET_FORM, SHEET_FORM_PREFIX, suffix=SHEET_FORM_SUFFIX)


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


def rename_month_sheets(flat, form, latest_month: str) -> None:
    flat.title = f"{SHEET_FLAT_PREFIX}{latest_month}"
    form.title = f"{SHEET_FORM_PREFIX}{latest_month}{SHEET_FORM_SUFFIX}"


def set_active_sheet(wb, ws, col: int, row: int) -> None:
    wb.active = wb.worksheets.index(ws)
    set_active_cell(ws, col, row)


def set_active_cell(ws, col: int, row: int) -> None:
    coordinate = f"{get_column_letter(col)}{row}"
    for selection in ws.sheet_view.selection:
        selection.activeCell = coordinate
        selection.sqref = coordinate
    if ws.sheet_view.pane:
        ws.sheet_view.pane.activePane = "topRight" if row <= int(ws.sheet_view.pane.ySplit or 0) else "bottomRight"
        ws.sheet_view.pane.topLeftCell = f"{get_column_letter(col)}{int(ws.sheet_view.pane.ySplit or 0) + 1}"


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


def ensure_flat_month(ws, month_info: dict) -> dict:
    existing = find_month_block(ws, month_info["key"])
    if existing:
        fill_month_headers(ws, existing["start"], month_info)
        set_active_cell(ws, existing["start"], MONTH_ROW)
        return existing

    latest = month_blocks(ws)[-1]

    # Source positions in the latest block's trailing section
    latest_sep_after = latest["end"] + 1
    latest_cum = latest_sep_after + 1
    latest_sep_mid = latest_cum + 2
    latest_grw = latest_sep_mid + 1

    # Layout for the new block
    start = ws.max_column + 1
    sep_col = start
    block_start = start + 1
    block_end = block_start + len(METRIC_ORDER) - 1
    cumulative_start = block_end + 2
    growth_start = cumulative_start + 3

    # Copy entire columns (values + styles + fills + number formats) from correct sources
    copy_columns(ws, latest["separator"] or latest["start"] - 1, sep_col, 1)
    copy_columns(ws, latest["start"], block_start, len(METRIC_ORDER))
    copy_columns(ws, latest_sep_after, block_end + 1, 1)
    copy_columns(ws, latest_cum, cumulative_start, 2)
    copy_columns(ws, latest_sep_mid, cumulative_start + 2, 1)
    copy_columns(ws, latest_grw, growth_start, 3)

    # Set header text (values only — fills/formats already copied)
    fill_month_headers(ws, block_start, month_info)
    ws.cell(MONTH_ROW, sep_col).value = "-"
    ws.cell(HEADER_ROW, sep_col).value = "-"
    ws.cell(MONTH_ROW, block_end + 1).value = "-"
    ws.cell(HEADER_ROW, block_end + 1).value = "-"
    ws.cell(MONTH_ROW, cumulative_start).value = f"Apr'25 to {month_info['key']}"
    ws.cell(HEADER_ROW, cumulative_start).value = "Funds Mobilized  (INR in crore)"
    ws.cell(HEADER_ROW, cumulative_start + 1).value = "Net Inflow (+ve)/Outflow (-ve)  (INR in crore)"
    ws.cell(MONTH_ROW, cumulative_start + 2).value = "-"
    ws.cell(HEADER_ROW, cumulative_start + 2).value = "-"
    ws.cell(MONTH_ROW, growth_start).value = "-"
    ws.cell(HEADER_ROW, growth_start).value = f"GS - Growth-({latest['month']} Vs {month_info['key']})"
    ws.cell(MONTH_ROW, growth_start + 1).value = "-"
    ws.cell(HEADER_ROW, growth_start + 1).value = f"NS - Growth-({latest['month']} Vs {month_info['key']})"
    ws.cell(MONTH_ROW, growth_start + 2).value = "-"
    ws.cell(HEADER_ROW, growth_start + 2).value = f"AUM - Growth-({latest['month']} Vs {month_info['key']})"

    recreate_flat_header_merges(ws, latest, block_start, cumulative_start, latest_cum)

    new_block = {"month": month_info["key"], "start": block_start, "end": block_end, "separator": sep_col}
    write_flat_formulas(ws, new_block, latest, cumulative_start, growth_start)
    set_active_cell(ws, block_start, MONTH_ROW)
    return new_block


def ensure_form_month(ws, month_info: dict) -> dict:
    existing = find_form_block(ws, month_info["full_month"], month_info["year"])
    if existing:
        return existing

    latest = form_blocks(ws)[-1]
    sep_col = ws.max_column + 1
    block_start = sep_col + 1
    block_end = block_start + len(METRIC_ORDER) + 1
    copy_columns(ws, latest["separator"] or latest["start"] - 1, sep_col, 1)
    copy_columns(ws, latest["start"], block_start, len(METRIC_ORDER) + 2)
    ws.cell(MONTH_ROW, sep_col).value = "-"
    ws.cell(HEADER_ROW, sep_col).value = "-"
    ws.cell(MONTH_ROW, block_start).value = f"Monthly Report for the month of {month_info['full_month']} {month_info['year']} "
    ws.cell(HEADER_ROW, block_start).value = "Sr "
    ws.cell(HEADER_ROW, block_start + 1).value = "Scheme Name "
    for offset, metric in enumerate(METRIC_ORDER, start=2):
        ws.cell(HEADER_ROW, block_start + offset).value = metric_label(metric, month_info)
    return {"month": month_info["key"], "start": block_start, "end": block_end, "separator": sep_col}


def update_flat_sheet(ws, records: list[dict], block: dict) -> list[str]:
    warnings = []
    scheme_rows = scheme_row_map(ws, 1)
    metric_cols = metric_columns(ws, block["start"], block["end"])
    for record in records:
        row_idx = scheme_rows.get(norm_key(record["scheme"]))
        if not row_idx:
            row_idx = insert_flat_record(ws, record)
            scheme_rows[norm_key(record["scheme"])] = row_idx
            warnings.append(f"Added new fund category to flat sheet: {record['scheme']}")
        for key, value in record["metrics"].items():
            col = metric_cols.get(key)
            if col and value is not None:
                ws.cell(row_idx, col).value = value
    return warnings


def update_form_sheet(ws, records: list[dict], block: dict) -> list[str]:
    warnings = []
    scheme_col = block["start"] + 1
    metric_start = block["start"] + 2
    rows = scheme_rows_map(ws, scheme_col)
    sr_col = block["start"]
    records_by_scheme = {norm_key(record["scheme"]): record for record in records}

    # Write individual scheme data — skip aggregate / sub-total rows so we
    # don't overwrite their cells with zeros before formulas are applied.
    for scheme_key, row_indexes in rows.items():
        record = records_by_scheme.get(scheme_key)
        for row_idx in row_indexes:
            if row_idx in _FORM_SUBTOTAL_ROWS or is_aggregate_scheme(ws.cell(row_idx, scheme_col).value or ""):
                continue
            for offset, metric in enumerate(METRIC_ORDER):
                ws.cell(row_idx, metric_start + offset).value = metric_cell_value(record, metric)

    for record in records:
        row_indexes = rows.get(norm_key(record["scheme"]))
        if not row_indexes:
            row_idx = ws.max_row + 1
            copy_row(ws, row_idx - 1, row_idx)
            ws.cell(row_idx, sr_col).value = ""
            ws.cell(row_idx, scheme_col).value = record["scheme"]
            row_indexes = [row_idx]
            rows[norm_key(record["scheme"])] = row_indexes
            warnings.append(f"Added new fund category to AMFI form sheet: {record['scheme']}")
        for row_idx in row_indexes:
            if row_idx in _FORM_SUBTOTAL_ROWS:
                continue
            for offset, metric in enumerate(METRIC_ORDER):
                ws.cell(row_idx, metric_start + offset).value = metric_cell_value(record, metric)

    # Generate SUM formulas for sub-total / total / grand-total rows
    write_form_subtotal_formulas(ws, metric_start)
    return warnings


def write_form_subtotal_formulas(ws, metric_start: int) -> None:
    """Write SUM formulas for every sub-total row in the form sheet.

    Each metric column in the block gets a formula like
    ``=SUM(XX6:XX21)`` or ``=SUM(XX6:XX21) + SUM(XX25:XX35) + …``
    where XX is the column letter for that metric.
    """
    for summary_row, row_ranges in _FORM_SUBTOTAL_ROWS.items():
        for offset in range(len(METRIC_ORDER)):
            col = metric_start + offset
            col_letter = get_column_letter(col)
            parts = [f"SUM({col_letter}{r1}:{col_letter}{r2})" for r1, r2 in row_ranges]
            ws.cell(summary_row, col).value = "=" + " + ".join(parts)


def insert_flat_record(ws, record: dict) -> int:
    asset = record.get("asset_type") or record.get("parent")
    row_idx = ws.max_row + 1
    if asset:
        for r in range(ws.max_row, DATA_START_ROW - 1, -1):
            if norm_key(ws.cell(r, 2).value) == norm_key(asset):
                row_idx = r + 1
                break
    ws.insert_rows(row_idx)
    copy_row(ws, row_idx - 1 if row_idx > DATA_START_ROW else row_idx + 1, row_idx)
    ws.cell(row_idx, 1).value = record["scheme"]
    ws.cell(row_idx, 2).value = asset
    ws.cell(row_idx, 3).value = record.get("scheme_structure")
    ws.cell(row_idx, 4).value = record.get("debt_equity")
    ws.cell(row_idx, 5).value = record.get("sales_prod_mis")
    return row_idx


def fill_month_headers(ws, block_start: int, month_info: dict) -> None:
    ws.cell(MONTH_ROW, block_start).value = month_info["key"]
    for offset, metric in enumerate(METRIC_ORDER):
        ws.cell(HEADER_ROW, block_start + offset).value = metric_label(metric, month_info)


def recreate_flat_header_merges(ws, previous_block: dict, month_start_col: int, summary_start_col: int, source_cum_col: int | None = None) -> None:
    source_cum = source_cum_col or (previous_block["start"] + 10)
    apply_merged_header(ws, 1, month_start_col, month_start_col + 8, previous_block["start"])
    apply_merged_header(ws, 2, month_start_col, month_start_col + 8, previous_block["start"])
    apply_merged_header(ws, 1, summary_start_col, summary_start_col + 1, source_cum)
    apply_merged_header(ws, 2, summary_start_col, summary_start_col + 1, source_cum)


def apply_merged_header(ws, row: int, start_col: int, end_col: int, source_col: int) -> None:
    source = merged_top_left(ws, row, source_col)
    style = {
        "fill": copy(source.fill),
        "font": copy(source.font),
        "alignment": copy(source.alignment),
    }
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row, col)
        cell.fill = copy(style["fill"])
        cell.font = copy(style["font"])
        cell.alignment = copy(style["alignment"])
    merge_range(ws, row, start_col, end_col)
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row, col)
        cell.fill = copy(style["fill"])
        cell.font = copy(style["font"])
        cell.alignment = copy(style["alignment"])


def merged_top_left(ws, row: int, col: int):
    for merged_range in ws.merged_cells.ranges:
        if merged_range.min_row <= row <= merged_range.max_row and merged_range.min_col <= col <= merged_range.max_col:
            return ws.cell(merged_range.min_row, merged_range.min_col)
    return ws.cell(row, col)


def merge_range(ws, row: int, start_col: int, end_col: int) -> None:
    for merged_range in list(ws.merged_cells.ranges):
        if merged_range.min_row <= row <= merged_range.max_row and merged_range.min_col <= start_col and end_col <= merged_range.max_col:
            ws.unmerge_cells(str(merged_range))
    ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)


def write_flat_formulas(ws, block: dict, previous: dict, cumulative_start: int, growth_start: int) -> None:
    all_blocks = month_blocks(ws)
    current_cols = metric_columns(ws, block["start"], block["end"])
    previous_cols = metric_columns(ws, previous["start"], previous["end"])
    funds_cols = [metric_columns(ws, b["start"], b["end"]).get("funds_mobilized") for b in all_blocks]
    net_cols = [metric_columns(ws, b["start"], b["end"]).get("net_inflow") for b in all_blocks]
    funds_cols = [c for c in funds_cols if c]
    net_cols = [c for c in net_cols if c]

    for row in range(DATA_START_ROW, ws.max_row + 1):
        fund_refs = ",".join(f"{get_column_letter(c)}{row}" for c in funds_cols)
        net_refs = ",".join(f"{get_column_letter(c)}{row}" for c in net_cols)
        ws.cell(row, cumulative_start).value = f"=SUM({fund_refs})"
        ws.cell(row, cumulative_start + 1).value = f"=SUM({net_refs})"
        ws.cell(row, growth_start).value = growth_formula(previous_cols.get("funds_mobilized"), current_cols.get("funds_mobilized"), row)
        ws.cell(row, growth_start + 1).value = growth_formula(previous_cols.get("net_inflow"), current_cols.get("net_inflow"), row)
        ws.cell(row, growth_start + 2).value = growth_formula(previous_cols.get("net_aum"), current_cols.get("net_aum"), row)


def growth_formula(previous_col: int | None, current_col: int | None, row: int) -> str:
    if not previous_col or not current_col or previous_col == current_col:
        return "=0"
    prev_col = get_column_letter(previous_col)
    curr_col = get_column_letter(current_col)
    return f"=IF({prev_col}{row}=0, 0, ({curr_col}{row}-{prev_col}{row})/{prev_col}{row})"


def month_blocks(ws) -> list[dict]:
    blocks = []
    for col in range(1, ws.max_column + 1):
        value = ws.cell(MONTH_ROW, col).value
        if is_month_key(value) and has_metric_header(ws, col):
            end = col
            while end + 1 <= ws.max_column and ws.cell(HEADER_ROW, end + 1).value != "-":
                end += 1
            sep = col - 1 if col > 1 and ws.cell(HEADER_ROW, col - 1).value == "-" else None
            blocks.append({"month": value, "start": col, "end": min(end, col + len(METRIC_ORDER) - 1), "separator": sep})
    return blocks


def form_blocks(ws) -> list[dict]:
    blocks = []
    for col in range(1, ws.max_column + 1):
        value = str(ws.cell(MONTH_ROW, col).value or "")
        if value.lower().startswith("monthly report"):
            sep = col - 1 if col > 1 and ws.cell(HEADER_ROW, col - 1).value == "-" else None
            blocks.append({"start": col, "end": col + len(METRIC_ORDER) + 1, "separator": sep})
    return blocks


def find_month_block(ws, month_key: str) -> dict | None:
    for block in month_blocks(ws):
        if block["month"] == month_key:
            return block
    return None


def find_form_block(ws, full_month: str, year: int) -> dict | None:
    needle = f"month of {full_month.lower()} {year}"
    for block in form_blocks(ws):
        if needle in str(ws.cell(MONTH_ROW, block["start"]).value or "").lower():
            return block
    return None


def has_metric_header(ws, col: int) -> bool:
    window = " ".join(norm(ws.cell(HEADER_ROW, c).value) for c in range(col, min(ws.max_column, col + 9) + 1))
    return "funds mobilized" in window or "no. of schemes" in window


def metric_columns(ws, start_col: int, end_col: int) -> dict:
    columns = {}
    for col in range(start_col, min(ws.max_column, end_col) + 1):
        key = classify_header(norm(ws.cell(HEADER_ROW, col).value))
        if key in METRIC_ORDER and key not in columns:
            columns[key] = col
    return columns


def copy_columns(ws, src_col: int, dst_col: int, width: int) -> None:
    for offset in range(width):
        source = src_col + offset
        target = dst_col + offset
        ws.column_dimensions[get_column_letter(target)].width = ws.column_dimensions[get_column_letter(source)].width
        for row in range(1, ws.max_row + 1):
            copy_cell(ws.cell(row, source), ws.cell(row, target))


def copy_row(ws, src_row: int, dst_row: int) -> None:
    ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height
    for col in range(1, ws.max_column + 1):
        copy_cell(ws.cell(src_row, col), ws.cell(dst_row, col))


def copy_cell(source, target) -> None:
    target._style = copy(source._style)
    if source.has_style:
        target.font = copy(source.font)
        target.fill = copy(source.fill)
        target.border = copy(source.border)
        target.alignment = copy(source.alignment)
        target.number_format = source.number_format
        target.protection = copy(source.protection)
    target.value = source.value


def scheme_row_map(ws, scheme_col: int) -> dict:
    mapping = {}
    for row in range(DATA_START_ROW, ws.max_row + 1):
        scheme = norm_key(ws.cell(row, scheme_col).value)
        if scheme:
            mapping[scheme] = row
    return mapping


def scheme_rows_map(ws, scheme_col: int) -> dict:
    mapping = {}
    for row in range(DATA_START_ROW, ws.max_row + 1):
        scheme = norm_key(ws.cell(row, scheme_col).value)
        if scheme:
            mapping.setdefault(scheme, []).append(row)
    return mapping


def sum_number_column(ws, col: int | None) -> float:
    if not col:
        return 0
    total = 0.0
    for row in range(DATA_START_ROW, ws.max_row + 1):
        value = ws.cell(row, col).value
        if isinstance(value, (int, float)):
            total += float(value)
    return total


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
    return key.startswith("sub total") or key.startswith("subtotal") or key.startswith("total ")


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


def metric_cell_value(record: dict | None, metric: str):
    if not record:
        return 0
    value = record["metrics"].get(metric)
    return 0 if value is None or pd.isna(value) else value
