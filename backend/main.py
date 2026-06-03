import io
import sqlite3
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from fastapi import FastAPI, HTTPException, UploadFile, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import RedirectResponse, StreamingResponse
from db import get_db_connection, init_db
import excel_engine

app = FastAPI(title="AMFI Dashboard API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.on_event("startup")
def startup_event():
    init_db()

@app.get("/")
def read_root():
    return RedirectResponse(url="/docs")

@app.get("/dashboard-data")
@app.get("/api/metrics")
def get_metrics(
    financial_year: str = Query(None),
    fy: str = Query(None)
):
    conn = get_db_connection()
    cursor = conn.cursor()
    
    target_fy = financial_year or fy
    if not target_fy:
        cursor.execute("SELECT MAX(financial_year) FROM amfi_metrics")
        row = cursor.fetchone()
        target_fy = row[0] if row else None
        
    conn.close()
    
    if not target_fy:
        return {"sheets": {}, "summary": {}, "warnings": [], "uploadMonth": None}
        
    try:
        excel_bytes = excel_engine.compile_excel_for_fy(target_fy)
        payload = excel_engine.dashboard_payload(excel_bytes, fy=target_fy)
        payload["financialYear"] = target_fy
        return payload
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Internal compiler error: {str(exc)}")

@app.post("/upload")
@app.post("/api/upload")
async def upload(file: UploadFile):
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Invalid file type. Please upload a valid Excel spreadsheet.")
    
    try:
        uploaded = await file.read()
        filename = file.filename or "upload.xlsx"
        
        upload_month, warnings = excel_engine.process_upload_db(uploaded, filename)
        
        from excel_engine import infer_month, get_financial_year
        m_info = infer_month("", upload_month)
        target_fy = get_financial_year(m_info["month"], m_info["year"])
        
        excel_bytes = excel_engine.compile_excel_for_fy(target_fy)
        payload = excel_engine.dashboard_payload(excel_bytes, warnings=warnings, upload_month=upload_month, fy=target_fy)
        payload["financialYear"] = target_fy
        return payload
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Database or processing error: {str(exc)}")

@app.get("/api/archives")
def get_archives():
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
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
        raise HTTPException(status_code=500, detail=str(exc))
    finally:
        conn.close()

@app.get("/api/download")
def download(
    financial_year: str = Query(None),
    fy: str = Query(None)
):
    target_fy = financial_year or fy
    if not target_fy:
        raise HTTPException(status_code=400, detail="Query parameter 'financial_year' or 'fy' is required.")
        
    try:
        excel_bytes = excel_engine.compile_excel_for_fy(target_fy)
        years = target_fy.split("-")
        start_yy = years[0][-2:]
        end_yy = years[1][-2:]
        filename = f"AMFI_MOM DATA - Apr'{start_yy} to Mar{end_yy}.xlsx"
        headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
        return StreamingResponse(
            io.BytesIO(excel_bytes),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
