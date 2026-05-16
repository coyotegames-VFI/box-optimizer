"""FastAPI app for box_optimizer."""

import logging
import json
import os
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path
from shutil import copyfileobj
from uuid import uuid4

from fastapi import Depends, FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse

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
from box_optimizer.workflow import inspect_workbook, optimize_workbook


logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("box_optimizer")

app = FastAPI(
    title="box_optimizer",
    version="0.1.0",
    openapi_url="/openapi.json",
)


class RequestStageError(Exception):
    """Internal error wrapper that keeps track of request stage."""

    def __init__(self, stage: str, exc: Exception):
        self.stage = stage
        self.exc = exc
        super().__init__(str(exc))


def _log_event(event: str, **fields) -> None:
    logger.info(event, extra={"box_optimizer": {"event": event, **fields}})


def _provided_api_key(
    x_api_key: str | None,
    authorization: str | None,
) -> str | None:
    if x_api_key:
        return x_api_key.strip()
    if authorization:
        provided = authorization.strip()
        if provided.lower().startswith("bearer "):
            return provided[len("Bearer ") :].strip()
        return provided
    return None


def require_api_key(
    x_api_key: str | None = Header(default=None, alias="X-API-Key"),
    authorization: str | None = Header(default=None),
) -> None:
    """Require an API key when BOX_OPTIMIZER_API_KEY is configured."""
    expected = (os.getenv("BOX_OPTIMIZER_API_KEY") or "").strip()
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
    if "debug" in parsed and not isinstance(parsed["debug"], bool):
        raise HTTPException(status_code=400, detail="debug must be true or false")
    if "max_orders" in parsed and parsed["max_orders"] is not None:
        if not isinstance(parsed["max_orders"], int) or parsed["max_orders"] < 1:
            raise HTTPException(status_code=400, detail="max_orders must be a positive integer")
    if "packing_mode" in parsed and parsed["packing_mode"] not in {"normal", "fast"}:
        raise HTTPException(status_code=400, detail='packing_mode must be "normal" or "fast"')
    return parsed


def _json_error(exc: Exception, stage: str, status_code: int = 500) -> JSONResponse:
    detail = exc.detail if isinstance(exc, HTTPException) else str(exc)
    return JSONResponse(
        status_code=status_code,
        content={
            "detail": detail or "Request failed",
            "stage": stage,
            "error_type": type(exc).__name__,
        },
    )


@app.get("/health")
def health() -> dict:
    """Health check endpoint."""
    return {"status": "ok"}


@app.get("/version")
def version() -> dict:
    """Return lightweight build/version information."""
    return {
        "app": "box_optimizer",
        "version": app.version,
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "git_commit": os.getenv("RAILWAY_GIT_COMMIT_SHA")
        or os.getenv("GIT_COMMIT")
        or "unknown",
    }


@app.post("/inspect", dependencies=[Depends(require_api_key)])
def inspect(
    sku_master_file: UploadFile = File(...),
    orders_file: UploadFile = File(...),
    config_json: str | None = Form(default=None),
):
    """Parse and match uploaded workbooks without packing or writing Excel."""
    started = time.perf_counter()
    stage = "request_received"
    run_id = uuid4().hex
    try:
        _log_event("request_received", endpoint="/inspect", run_id=run_id)
        stage = "config"
        config = _parse_config(config_json)
        if config.get("debug"):
            logger.setLevel(logging.DEBUG)

        work_dir = Path(tempfile.gettempdir()) / "box_optimizer_api" / run_id
        work_dir.mkdir(parents=True, exist_ok=True)

        _log_event(
            "filenames_received",
            endpoint="/inspect",
            sku_master_file=sku_master_file.filename,
            orders_file=orders_file.filename,
        )
        stage = "files_saved"
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
        _log_event("files_saved", endpoint="/inspect", run_id=run_id)

        stage = "inspect"
        result = inspect_workbook(
            sku_master_path=str(sku_master_path),
            orders_path=str(orders_path),
            config=config,
        )
        result["elapsed_seconds"] = round(time.perf_counter() - started, 3)
        _log_event(
            "response_returned",
            endpoint="/inspect",
            run_id=run_id,
            elapsed_seconds=result["elapsed_seconds"],
        )
        return result
    except HTTPException as exc:
        return _json_error(exc, stage, status_code=exc.status_code)
    except Exception as exc:
        logger.exception("inspect failed", extra={"box_optimizer": {"stage": stage}})
        return _json_error(exc, stage)


@app.post("/optimize", dependencies=[Depends(require_api_key)])
def optimize(
    sku_master_file: UploadFile = File(...),
    orders_file: UploadFile = File(...),
    config_json: str | None = Form(default=None),
) -> FileResponse:
    """Optimize uploaded SKU and order workbooks and return the XLSX result."""
    started = time.perf_counter()
    stage = "request_received"
    run_id = uuid4().hex
    try:
        _log_event("request_received", endpoint="/optimize", run_id=run_id)
        stage = "config"
        config = _parse_config(config_json)
        if config.get("debug"):
            logger.setLevel(logging.DEBUG)

        work_dir = Path(tempfile.gettempdir()) / "box_optimizer_api" / run_id
        work_dir.mkdir(parents=True, exist_ok=True)

        _log_event(
            "filenames_received",
            endpoint="/optimize",
            sku_master_file=sku_master_file.filename,
            orders_file=orders_file.filename,
        )
        stage = "files_saved"
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
        _log_event("files_saved", endpoint="/optimize", run_id=run_id)

        stage = "optimize"
        optimize_workbook(
            sku_master_path=str(sku_master_path),
            orders_path=str(orders_path),
            output_path=str(output_path),
            config=config,
        )

        _log_event(
            "response_returned",
            endpoint="/optimize",
            run_id=run_id,
            elapsed_seconds=round(time.perf_counter() - started, 3),
        )
        return FileResponse(
            path=str(output_path),
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename="optimized_shipping_plan.xlsx",
        )
    except HTTPException as exc:
        return _json_error(exc, stage, status_code=exc.status_code)
    except Exception as exc:
        logger.exception("optimize failed", extra={"box_optimizer": {"stage": stage}})
        return _json_error(exc, stage)


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
    "inspect_workbook",
    "optimize_workbook",
]
