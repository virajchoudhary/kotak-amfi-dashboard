import io

from fastapi import FastAPI, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import RedirectResponse, StreamingResponse

try:
    from .excel_engine import dashboard_payload, process_upload, template_bytes
except ImportError:
    from excel_engine import dashboard_payload, process_upload, template_bytes


app = FastAPI(title="AMFI Dashboard API")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/")
def read_root():
    return RedirectResponse(url="/docs")


current_workbook = template_bytes()


@app.get("/dashboard-data")
def get_dashboard_data():
    return dashboard_payload(current_workbook)


@app.post("/upload")
async def upload(file: UploadFile):
    global current_workbook
    try:
        uploaded = await file.read()
        current_workbook, payload = process_upload(current_workbook, uploaded, file.filename or "upload.xlsx")
        return payload
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.get("/download")
def download():
    headers = {"Content-Disposition": "attachment; filename=updated-amfi-dashboard.xlsx"}
    return StreamingResponse(
        io.BytesIO(current_workbook),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="0.0.0.0", port=8000)
