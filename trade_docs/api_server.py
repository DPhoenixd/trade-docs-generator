from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Annotated

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse

from .api_service import analyze_files, download_path, generate_documents, generate_invoice_document, search_fabrics


app = FastAPI(title="P.I/P.L一点腾 API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:5173",
        "http://127.0.0.1:5173",
        "https://dphoenixd.github.io",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/api/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/api/analyze")
async def analyze(files: Annotated[list[UploadFile], File()] = []) -> dict:
    saved: list[Path] = []
    with tempfile.TemporaryDirectory(prefix="pipl-upload-") as tmp:
        tmp_path = Path(tmp)
        for upload in files:
            target = tmp_path / Path(upload.filename or "upload").name
            target.write_bytes(await upload.read())
            saved.append(target)
        try:
            return analyze_files(saved)
        except Exception as exc:  # noqa: BLE001
            raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/generate")
def generate(payload: dict) -> dict:
    try:
        return generate_documents(payload)
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.post("/api/generate/pi")
def generate_pi(payload: dict) -> dict:
    try:
        return generate_invoice_document(payload)
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.get("/api/fabrics/search")
def fabrics_search(q: str = "", limit: int = 20) -> dict:
    try:
        return search_fabrics(q, limit)
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=400, detail=str(exc)) from exc


@app.get("/api/download/{session_id}/{kind}")
def download(session_id: str, kind: str) -> FileResponse:
    if kind not in {"pi", "pl"}:
        raise HTTPException(status_code=404, detail="Unknown file type")
    try:
        path = download_path(session_id, kind)
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    return FileResponse(path, filename=path.name)
