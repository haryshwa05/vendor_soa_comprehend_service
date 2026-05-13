from __future__ import annotations

from pathlib import Path
from tempfile import NamedTemporaryFile

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import JSONResponse, StreamingResponse
from pydantic import BaseModel

from msg_to_excel import (
    MsgConversionError,
    extract_msg_to_rows,
    next_available_output_path,
    rows_to_dicts,
    workbook_bytes_from_msg,
)


class ExtractResponse(BaseModel):
    filename: str
    source_type: str
    source_name: str
    source_details: str
    row_count: int
    rows: list[dict[str, str]]


app = FastAPI(title="Vendor SOA Extractor API", version="1.0.0")


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/extract", response_model=ExtractResponse)
async def extract_msg(file: UploadFile = File(...)) -> ExtractResponse:
    temp_path = await save_upload_to_temp(file)
    try:
        result = extract_msg_to_rows(temp_path)
        return ExtractResponse(
            filename=file.filename or Path(temp_path).name,
            source_type=result.source_type,
            source_name=result.source_name,
            source_details=result.source_details,
            row_count=len(result.rows),
            rows=rows_to_dicts(result.rows),
        )
    except MsgConversionError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    finally:
        cleanup_temp_file(temp_path)


@app.post("/extract/xlsx")
async def extract_msg_xlsx(file: UploadFile = File(...)) -> StreamingResponse:
    temp_path = await save_upload_to_temp(file)
    try:
        result, workbook_bytes = workbook_bytes_from_msg(temp_path)
        output_name = next_available_output_path(Path(file.filename or temp_path)).name
        headers = {
            "X-Source-Type": result.source_type,
            "X-Source-Name": result.source_name,
            "X-Source-Details": result.source_details,
            "X-Row-Count": str(len(result.rows)),
            "Content-Disposition": f'attachment; filename="{output_name}"',
        }
        return StreamingResponse(
            iter([workbook_bytes]),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
        )
    except MsgConversionError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    finally:
        cleanup_temp_file(temp_path)


@app.exception_handler(Exception)
async def unhandled_exception_handler(_, exc: Exception) -> JSONResponse:
    return JSONResponse(status_code=500, content={"detail": f"Unexpected error: {exc}"})


async def save_upload_to_temp(file: UploadFile) -> Path:
    suffix = Path(file.filename or "upload.msg").suffix or ".msg"
    with NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        while chunk := await file.read(1024 * 1024):
            tmp.write(chunk)
        return Path(tmp.name)


def cleanup_temp_file(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)
    except OSError:
        pass
