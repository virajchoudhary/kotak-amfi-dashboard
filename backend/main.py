import io
import logging
import os
import re
import sqlite3
import sys
import threading
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from fastapi import Depends, FastAPI, HTTPException, Query, Request, UploadFile, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.middleware.trustedhost import TrustedHostMiddleware
from fastapi.responses import JSONResponse, RedirectResponse, StreamingResponse
from db import database_path, get_db_connection, init_db
import excel_engine

LOGGER = logging.getLogger("amfi_dashboard")
BACKEND_DIR = Path(__file__).resolve().parent


def _env_bool(name: str, default: bool = False) -> bool:
    value = os.getenv(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}


def _env_int(name: str, default: int) -> int:
    value = os.getenv(name)
    if not value:
        return default
    try:
        parsed = int(value)
    except ValueError:
        LOGGER.warning("Invalid integer for %s=%r; using default %s", name, value, default)
        return default
    return parsed if parsed > 0 else default


def _env_list(name: str) -> list[str]:
    value = os.getenv(name, "")
    return [item.strip() for item in value.split(",") if item.strip()]


APP_ENV = os.getenv("APP_ENV", "development").strip().lower()
IS_PRODUCTION = APP_ENV in {"prod", "production"}
ENABLE_DOCS = _env_bool("AMFI_ENABLE_DOCS", default=not IS_PRODUCTION)
IDENTITY_HEADER = os.getenv("AMFI_IDENTITY_HEADER", "X-Forwarded-User").strip() or "X-Forwarded-User"
REQUIRE_PROXY_IDENTITY = _env_bool("AMFI_REQUIRE_PROXY_IDENTITY", default=IS_PRODUCTION)
MAX_UPLOAD_BYTES = _env_int("AMFI_MAX_UPLOAD_MB", 25) * 1024 * 1024
MAX_XLSX_UNZIPPED_BYTES = _env_int("AMFI_MAX_WORKBOOK_UNZIPPED_MB", 100) * 1024 * 1024


def _default_dev_origins() -> list[str]:
    return [
        "http://localhost:3000",
        "http://127.0.0.1:3000",
        "http://localhost:4173",
        "http://127.0.0.1:4173",
        "http://localhost:5173",
        "http://127.0.0.1:5173",
        "http://localhost:8000",
        "http://127.0.0.1:8000",
    ]


ALLOWED_ORIGINS = _env_list("AMFI_ALLOWED_ORIGINS") or ([] if IS_PRODUCTION else _default_dev_origins())
ALLOWED_HOSTS = _env_list("AMFI_ALLOWED_HOSTS") or ([] if IS_PRODUCTION else ["localhost", "127.0.0.1", "testserver"])

app = FastAPI(
    title="AMFI Dashboard API",
    docs_url="/docs" if ENABLE_DOCS else None,
    redoc_url="/redoc" if ENABLE_DOCS else None,
    openapi_url="/openapi.json" if ENABLE_DOCS else None,
)

if ALLOWED_HOSTS:
    app.add_middleware(TrustedHostMiddleware, allowed_hosts=ALLOWED_HOSTS)

if ALLOWED_ORIGINS:
    app.add_middleware(
        CORSMiddleware,
        allow_origins=ALLOWED_ORIGINS,
        allow_credentials=True,
        allow_methods=["GET", "POST", "OPTIONS"],
        allow_headers=["Authorization", "Content-Type", "X-Requested-With", IDENTITY_HEADER],
    )


_FY_RE = re.compile(r"^(20\d{2})-(20\d{2})$")
_XLS_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
_XLSX_MAGIC = b"PK\x03\x04"


def require_identity(request: Request) -> str | None:
    identity = (request.headers.get(IDENTITY_HEADER) or "").strip()
    if REQUIRE_PROXY_IDENTITY and not identity:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Authentication required.",
        )
    return identity or None


def audit_log(request: Request, action: str, identity: str | None, **extra) -> None:
    client = request.client.host if request.client else "-"
    LOGGER.info(
        "audit action=%s identity=%s client=%s path=%s extra=%s",
        action,
        identity or "anonymous",
        client,
        request.url.path,
        extra,
    )


def _internal_error(action: str, exc: Exception) -> HTTPException:
    LOGGER.exception("%s failed", action)
    return HTTPException(status_code=500, detail="Internal server error.")


def _validate_financial_year(value: str | None) -> str | None:
    if not value:
        return None
    candidate = value.strip()
    match = _FY_RE.fullmatch(candidate)
    if not match:
        raise HTTPException(status_code=400, detail="Invalid financial year format. Expected YYYY-YYYY.")
    start, end = int(match.group(1)), int(match.group(2))
    if end != start + 1:
        raise HTTPException(status_code=400, detail="Invalid financial year range.")
    return candidate


def _safe_download_filename(financial_year: str) -> str:
    start, end = financial_year.split("-")
    return f"AMFI_MOM_DATA_Apr{start[-2:]}_to_Mar{end[-2:]}.xlsx"


def _validate_xlsx_container(data: bytes) -> None:
    if not data.startswith(_XLSX_MAGIC):
        raise HTTPException(status_code=400, detail="Invalid XLSX file signature.")
    try:
        with zipfile.ZipFile(io.BytesIO(data), "r") as workbook_zip:
            names = set(workbook_zip.namelist())
            if "[Content_Types].xml" not in names or "xl/workbook.xml" not in names:
                raise HTTPException(status_code=400, detail="Invalid XLSX workbook structure.")

            total_uncompressed = 0
            for item in workbook_zip.infolist():
                parts = Path(item.filename).parts
                if item.filename.startswith("/") or ".." in parts:
                    raise HTTPException(status_code=400, detail="Invalid XLSX workbook path.")
                total_uncompressed += item.file_size
                if total_uncompressed > MAX_XLSX_UNZIPPED_BYTES:
                    raise HTTPException(status_code=400, detail="Uploaded workbook expands beyond the allowed size.")
    except zipfile.BadZipFile:
        raise HTTPException(status_code=400, detail="Invalid XLSX workbook.")


def _validate_upload(filename: str, data: bytes) -> None:
    suffix = Path(filename or "").suffix.lower()
    if suffix not in {".xlsx", ".xls"}:
        raise HTTPException(status_code=400, detail="Invalid file type. Please upload a valid Excel spreadsheet.")
    if not data:
        raise HTTPException(status_code=400, detail="Uploaded file is empty.")
    if len(data) > MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=413, detail=f"Uploaded file exceeds {MAX_UPLOAD_BYTES // (1024 * 1024)} MB.")
    if suffix == ".xlsx":
        _validate_xlsx_container(data)
        return
    if not data.startswith(_XLS_MAGIC):
        raise HTTPException(status_code=400, detail="Invalid XLS file signature.")


def _is_relative_to(path: Path, parent: Path) -> bool:
    try:
        path.resolve().relative_to(parent.resolve())
        return True
    except ValueError:
        return False

def _warm_dashboard_cache() -> None:
    """Pre-compile the latest FY dashboard once at boot so the first page load is
    instant instead of paying the workbook-compile cost. Read-only and best-effort."""
    try:
        conn = get_db_connection()
        try:
            row = conn.execute("SELECT MAX(financial_year) FROM amfi_metrics").fetchone()
        finally:
            conn.close()
        target_fy = row[0] if row else None
        if target_fy:
            excel_engine.compile_dashboard_payload(target_fy)
            LOGGER.info("Dashboard cache warmed for %s", target_fy)
    except Exception:
        LOGGER.warning("Dashboard cache warm-up skipped", exc_info=True)


@app.on_event("startup")
def startup_event():
    init_db()
    threading.Thread(target=_warm_dashboard_cache, name="dashboard-cache-warm", daemon=True).start()

@app.get("/")
def read_root():
    if not ENABLE_DOCS:
        return {"status": "ok"}
    return RedirectResponse(url="/docs")

@app.get("/api/health")
def health():
    return {"status": "ok", "environment": APP_ENV}

@app.get("/api/readiness")
def readiness():
    checks = {
        "environment": APP_ENV,
        "docsEnabled": ENABLE_DOCS,
        "proxyIdentityRequired": REQUIRE_PROXY_IDENTITY,
        "allowedOriginsConfigured": bool(ALLOWED_ORIGINS),
        "allowedHostsConfigured": bool(ALLOWED_HOSTS),
        "databasePath": str(database_path()),
        "templatePath": str(excel_engine.TEMPLATE_PATH),
    }
    errors = []

    db_path = database_path()
    if IS_PRODUCTION and not os.getenv("AMFI_DB_PATH"):
        errors.append("AMFI_DB_PATH is required in production.")
    if IS_PRODUCTION and _is_relative_to(db_path, BACKEND_DIR):
        errors.append("Production database must live outside the application directory.")
    if not db_path.parent.exists():
        errors.append("Database directory does not exist.")
    elif not os.access(db_path.parent, os.W_OK):
        errors.append("Database directory is not writable.")
    if db_path.exists():
        if not os.access(db_path, os.R_OK | os.W_OK):
            errors.append("Database file is not readable and writable.")
        else:
            try:
                uri = f"file:{db_path.as_posix()}?mode=ro"
                with sqlite3.connect(uri, uri=True, timeout=5.0) as conn:
                    conn.execute("SELECT 1")
            except sqlite3.Error:
                errors.append("Database file cannot be opened read-only.")

    if not excel_engine.TEMPLATE_PATH.exists():
        errors.append("Template workbook is missing.")
    if IS_PRODUCTION and not ALLOWED_HOSTS:
        errors.append("AMFI_ALLOWED_HOSTS is required in production.")
    if IS_PRODUCTION and ENABLE_DOCS:
        errors.append("API docs must be disabled in production unless explicitly approved.")
    if IS_PRODUCTION and not REQUIRE_PROXY_IDENTITY:
        errors.append("Proxy identity enforcement must be enabled in production.")

    ready = not errors
    return JSONResponse(
        status_code=200 if ready else 503,
        content={"status": "ready" if ready else "not_ready", "checks": checks, "errors": errors},
    )

@app.get("/dashboard-data")
@app.get("/api/metrics")
def get_metrics(
    request: Request,
    financial_year: str = Query(None),
    fy: str = Query(None),
    identity: str | None = Depends(require_identity),
):
    target_fy = _validate_financial_year(financial_year or fy)
    if not target_fy:
        conn = get_db_connection()
        try:
            cursor = conn.cursor()
            cursor.execute("SELECT MAX(financial_year) FROM amfi_metrics")
            row = cursor.fetchone()
            target_fy = row[0] if row else None
        finally:
            conn.close()
    
    if not target_fy:
        return {"sheets": {}, "summary": {}, "warnings": [], "uploadMonth": None}
        
    try:
        audit_log(request, "metrics.read", identity, financial_year=target_fy)
        return excel_engine.compile_dashboard_payload(target_fy)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    except Exception as exc:
        raise _internal_error("metrics.read", exc)

@app.post("/upload")
@app.post("/api/upload")
async def upload(
    request: Request,
    file: UploadFile,
    identity: str | None = Depends(require_identity),
):
    filename = file.filename or ""
    try:
        uploaded = await file.read(MAX_UPLOAD_BYTES + 1)
        _validate_upload(filename, uploaded)
        audit_log(request, "upload.write", identity, filename=Path(filename).name, size=len(uploaded))
        upload_month, warnings = excel_engine.process_upload_db(uploaded, filename)
        excel_engine.invalidate_dashboard_cache()

        from excel_engine import infer_month, get_financial_year
        m_info = infer_month("", upload_month)
        target_fy = get_financial_year(m_info["month"], m_info["year"])
        
        excel_bytes = excel_engine.compile_excel_for_fy(target_fy)
        payload = excel_engine.dashboard_payload(excel_bytes, warnings=warnings, upload_month=upload_month, fy=target_fy)
        payload["financialYear"] = target_fy
        return payload
    except HTTPException:
        raise
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    except Exception as exc:
        raise _internal_error("upload.write", exc)

@app.get("/api/archives")
def get_archives(
    request: Request,
    identity: str | None = Depends(require_identity),
):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        audit_log(request, "archives.read", identity)
        cursor.execute("""
            SELECT financial_year, COUNT(*) as record_count, MAX(last_modified) as last_modified
            FROM amfi_metrics
            GROUP BY financial_year
            ORDER BY financial_year DESC
        """)
        rows = cursor.fetchall()
        
        archives = []
        for r in rows:
            archives.append({
                "financial_year": r["financial_year"],
                "record_count": r["record_count"],
                "last_modified": r["last_modified"],
                "status": "Finalized" if r["record_count"] >= 588 else "In Progress"
            })
        return archives
    except Exception as exc:
        raise _internal_error("archives.read", exc)
    finally:
        conn.close()

@app.get("/api/download")
def download(
    request: Request,
    financial_year: str = Query(None),
    fy: str = Query(None),
    identity: str | None = Depends(require_identity),
):
    target_fy = _validate_financial_year(financial_year or fy)
    if not target_fy:
        raise HTTPException(status_code=400, detail="Query parameter 'financial_year' or 'fy' is required.")
        
    try:
        audit_log(request, "workbook.download", identity, financial_year=target_fy)
        excel_bytes = excel_engine.compile_excel_for_fy(target_fy)
        filename = _safe_download_filename(target_fy)
        headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
        return StreamingResponse(
            io.BytesIO(excel_bytes),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    except Exception as exc:
        raise _internal_error("workbook.download", exc)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host=os.getenv("AMFI_HOST", "127.0.0.1"), port=_env_int("AMFI_PORT", 8000))
