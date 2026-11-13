"""POST /evaluate: multipart `submission` (.pptx), `spec` (JSON), optional `files` (PDFs, EVIDENCE) and `candidate` (JSON)."""
from __future__ import annotations

import json

from fastapi import FastAPI, File, Form, HTTPException, UploadFile

from .engine import BadPackage, evaluate
from .spec import Report, Spec

MAX_BYTES = 60 * 1024 * 1024
app = FastAPI(title="powerpoint-evaluator", version="0.1.0")


async def _read(f: UploadFile) -> bytes:
    data = await f.read(MAX_BYTES + 1)
    if len(data) > MAX_BYTES:
        raise HTTPException(413, f"{f.filename}: file too large")
    return data


@app.get("/health")
def health() -> dict:
    return {"ok": True}


@app.post("/evaluate", response_model=Report)
async def evaluate_endpoint(
    submission: UploadFile = File(...),
    spec: str = Form(..., description="requirements spec as JSON (reference images may be embedded as base64)"),
    files: list[UploadFile] = File(default=[]),
    candidate: str | None = Form(default=None),
) -> Report:
    try:
        sp = Spec.model_validate_json(spec)
        cand = json.loads(candidate) if candidate else {}
    except Exception as e:
        raise HTTPException(422, f"invalid spec or candidate JSON: {type(e).__name__}") from e
    extras = {f.filename or f"file{i}": await _read(f) for i, f in enumerate(files)}
    try:
        return evaluate(sp, await _read(submission), extras, cand)
    except BadPackage as e:
        raise HTTPException(422, str(e)) from e
