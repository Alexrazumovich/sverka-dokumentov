import io
import os
import re
import uuid
from datetime import datetime

import pandas as pd
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from reconciler import reconcile_files

BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR   = os.path.join(BASE_DIR, "static")
RESULTS_DIR  = os.path.join(BASE_DIR, "result_reports")
os.makedirs(RESULTS_DIR, exist_ok=True)

app = FastAPI(title="Сверка документов")
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


@app.middleware("http")
async def no_cache(request, call_next):
    response = await call_next(request)
    if not request.url.path.startswith("/api/"):
        response.headers["Cache-Control"] = "no-store"
    return response


def _read_columns(content: bytes, has_header: bool) -> list[str]:
    if has_header:
        df = pd.read_excel(io.BytesIO(content), nrows=0)
        return [str(c) for c in df.columns]
    else:
        df = pd.read_excel(io.BytesIO(content), header=None, nrows=1)
        return [f"Колонка {i + 1}" for i in range(len(df.columns))]


@app.get("/")
async def root():
    return FileResponse(os.path.join(STATIC_DIR, "index.html"))


@app.get("/en")
async def root_en():
    return FileResponse(os.path.join(STATIC_DIR, "index_en.html"))


@app.get("/help")
async def help_page():
    return FileResponse(os.path.join(STATIC_DIR, "help.html"))


@app.get("/en/help")
async def help_en():
    return FileResponse(os.path.join(STATIC_DIR, "help_en.html"))


@app.post("/api/preview")
async def preview(
    file_a: UploadFile = File(...),
    file_b: UploadFile = File(...),
    has_header_a: str = Form("true"),
    has_header_b: str = Form("true"),
):
    try:
        content_a = await file_a.read()
        content_b = await file_b.read()
        cols_a = _read_columns(content_a, has_header_a.lower() != "false")
        cols_b = _read_columns(content_b, has_header_b.lower() != "false")
        return {
            "columns_a": cols_a,
            "columns_b": cols_b,
            "name_a": file_a.filename,
            "name_b": file_b.filename,
        }
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/api/reconcile")
async def reconcile(
    file_a: UploadFile = File(...),
    file_b: UploadFile = File(...),
    date_col_a:     str = Form(...),
    amount_col_a:   str = Form(...),
    date_col_b:     str = Form(...),
    amount_col_b:   str = Form(...),
    date_tolerance: str = Form("0"),
    has_header_a:   str = Form("true"),
    has_header_b:   str = Form("true"),
    lang:           str = Form("ru"),
):
    try:
        content_a = await file_a.read()
        content_b = await file_b.read()
        result_bytes, summary = reconcile_files(
            content_a, content_b,
            date_col_a, amount_col_a,
            date_col_b, amount_col_b,
            date_tolerance,
            has_header_a.lower() != "false",
            has_header_b.lower() != "false",
            lang if lang in ("ru", "en") else "ru",
        )

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        token = f"{timestamp}_{uuid.uuid4().hex[:8]}"
        filename = f"sverka_{token}.xlsx"
        save_path = os.path.join(RESULTS_DIR, filename)
        with open(save_path, "wb") as f:
            f.write(result_bytes)

        return JSONResponse({
            "summary":      summary,
            "download_url": f"/api/download/{token}",
            "saved_path":   save_path,
        })
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.get("/api/download/{token}")
async def download(token: str):
    if not re.match(r'^[a-zA-Z0-9_-]+$', token):
        raise HTTPException(status_code=400, detail="Некорректный токен")
    path = os.path.join(RESULTS_DIR, f"sverka_{token}.xlsx")
    if not os.path.abspath(path).startswith(os.path.abspath(RESULTS_DIR)):
        raise HTTPException(status_code=400, detail="Некорректный путь")
    if not os.path.exists(path):
        raise HTTPException(status_code=404, detail="Файл не найден")
    return FileResponse(
        path,
        filename=f"sverka_{token}.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    import sys, uvicorn
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    port = int(os.environ.get("PORT", 8001))
    print(f"Results saved to: {RESULTS_DIR}")
    print(f"Open browser: http://localhost:{port}")
    uvicorn.run("main:app", host="0.0.0.0", port=port, reload=False)
