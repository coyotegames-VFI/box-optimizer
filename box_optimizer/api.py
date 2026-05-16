"""FastAPI app for box_optimizer."""

import json
import os
import tempfile
from pathlib import Path
from shutil import copyfileobj
from uuid import uuid4

from fastapi import Depends, FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.responses import FileResponse

from box_optimizer.models import (
    Carton,
    Dimensions,
    OrderLine,
    PackedItem,
    SKU,
    SKUItem,
    UnmatchedSKURecord,
)
from box_optimizer.normalize import normalize_sku
from box_optimizer.workflow import optimize_workbook


app = FastAPI(
    title="box_optimizer",
    version="0.1.0",
    openapi_url="/openapi.json",
)


def _provided_api_key(
    x_api_key: str | None,
    authorization: str | None,
) -> str | None:
    if x_api_key:
        return x_api_key
    if authorization and authorization.lower().startswith("bearer "):
        return authorization.split(" ", 1)[1]
    return authorization


def require_api_key(
    x_api_key: str | None = Header(default=None, alias="X-API-Key"),
    authorization: str | None = Header(default=None),
) -> None:
    """Require an API key when BOX_OPTIMIZER_API_KEY is configured."""
    expected = os.getenv("BOX_OPTIMIZER_API_KEY")
    if not expected:
        return

    provided = _provided_api_key(x_api_key, authorization)
    if provided != expected:
        raise HTTPException(status_code=401, detail="Invalid or missing API key")


def _upload_suffix(upload: UploadFile) -> str:
    suffix = Path(upload.filename or "").suffix.lower()
    return suffix if suffix in {".csv", ".xlsx"} else ".xlsx"


def _save_upload(upload: UploadFile, directory: Path, filename: str) -> Path:
    target = directory / filename
    upload.file.seek(0)
    with open(target, "wb") as output:
        copyfileobj(upload.file, output)
    return target


def _parse_config(config_json: str | None) -> dict:
    if not config_json:
        return {}
    try:
        parsed = json.loads(config_json)
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail="config_json must be valid JSON") from exc
    if not isinstance(parsed, dict):
        raise HTTPException(status_code=400, detail="config_json must be a JSON object")
    return parsed


@app.get("/health")
def health() -> dict:
    """Health check endpoint."""
    return {"status": "ok"}


@app.post("/optimize", dependencies=[Depends(require_api_key)])
def optimize(
    sku_master_file: UploadFile = File(...),
    orders_file: UploadFile = File(...),
    config_json: str | None = Form(default=None),
) -> FileResponse:
    """Optimize uploaded SKU and order workbooks and return the XLSX result."""
    config = _parse_config(config_json)
    run_id = uuid4().hex
    work_dir = Path(tempfile.gettempdir()) / "box_optimizer_api" / run_id
    work_dir.mkdir(parents=True, exist_ok=True)

    sku_master_path = _save_upload(
        sku_master_file,
        work_dir,
        f"sku_master{_upload_suffix(sku_master_file)}",
    )
    orders_path = _save_upload(
        orders_file,
        work_dir,
        f"orders{_upload_suffix(orders_file)}",
    )
    output_path = work_dir / "optimized_shipping_plan.xlsx"

    optimize_workbook(
        sku_master_path=str(sku_master_path),
        orders_path=str(orders_path),
        output_path=str(output_path),
        config=config,
    )

    return FileResponse(
        path=str(output_path),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="optimized_shipping_plan.xlsx",
    )


__all__ = [
    "Carton",
    "Dimensions",
    "OrderLine",
    "PackedItem",
    "SKU",
    "SKUItem",
    "UnmatchedSKURecord",
    "app",
    "normalize_sku",
    "optimize_workbook",
]
