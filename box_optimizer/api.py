"""FastAPI app for box_optimizer."""

import base64
import binascii
import logging
import json
import os
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path
from shutil import copyfileobj, rmtree
from typing import Any
from uuid import uuid4

from fastapi import Depends, FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from pydantic import BaseModel, Field

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


class HealthResponse(BaseModel):
    status: str


class VersionResponse(BaseModel):
    app: str
    version: str
    timestamp: str
    git_commit: str


class Base64WorkbookRequest(BaseModel):
    """JSON file transport used by GPT Actions when multipart uploads are unavailable."""

    sku_master_filename: str = Field(..., min_length=1, examples=["sku_master.xlsx"])
    sku_master_base64: str = Field(..., min_length=1, examples=["UEsDB..."])
    orders_filename: str = Field(..., min_length=1, examples=["orders.xlsx"])
    orders_base64: str = Field(..., min_length=1, examples=["UEsDB..."])
    config_json: dict[str, Any] | str | None = Field(
        default=None,
        examples=[
            {
                "debug": True,
                "max_orders": 5,
                "packing_mode": "fast",
                "output_granularity": "order_summary",
                "preserve_region_sheets": False,
            }
        ],
    )


class InspectSummaryResponse(BaseModel):
    sku_items: int
    order_rows: int
    wide_product_columns: int
    order_lines: int
    matched: int
    unmatched: int
    sheets_read: list[str]
    detected_sku_columns: list[str]
    detected_order_columns: list[str]
    detected_product_quantity_columns: list[str]
    matched_rule_keys: list[str]
    unmatched_rule_keys: list[str]
    warnings: list[str]
    elapsed_seconds: float


class OptimizeSummaryResponse(BaseModel):
    output_path: str | None = None
    orders_processed: int | None = None
    boxes_created: int | None = None
    box_types: int | None = None
    unmatched_skus: int | None = None
    warnings: list[str] = Field(default_factory=list)
    warning_count: int | None = None
    multi_box_order_count: int | None = None
    rules_applied_count: int | None = None
    elapsed_seconds: float | None = None


class OptimizeBase64Response(BaseModel):
    filename: str
    content_type: str
    workbook_base64: str
    summary: OptimizeSummaryResponse


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


def _filename_suffix(filename: str, field_name: str) -> str:
    suffix = Path(filename or "").suffix.lower()
    if suffix not in {".csv", ".xlsx"}:
        raise HTTPException(
            status_code=400,
            detail=f"{field_name} must end in .csv or .xlsx",
        )
    return suffix


def _save_upload(upload: UploadFile, directory: Path, filename: str) -> Path:
    target = directory / filename
    upload.file.seek(0)
    with open(target, "wb") as output:
        copyfileobj(upload.file, output)
    return target


def _save_base64_file(
    encoded: str,
    directory: Path,
    filename: str,
    field_name: str,
) -> Path:
    target = directory / filename
    payload = encoded.strip()
    if "," in payload and payload.lower().startswith("data:"):
        payload = payload.split(",", 1)[1]
    try:
        decoded = base64.b64decode(payload, validate=True)
    except (binascii.Error, ValueError) as exc:
        raise HTTPException(status_code=400, detail=f"{field_name} must be valid base64") from exc
    if not decoded:
        raise HTTPException(status_code=400, detail=f"{field_name} decoded to an empty file")
    with open(target, "wb") as output:
        output.write(decoded)
    return target


def _parse_config(config_json: str | dict[str, Any] | None) -> dict:
    if not config_json:
        return {}
    if isinstance(config_json, str):
        try:
            parsed = json.loads(config_json)
        except json.JSONDecodeError as exc:
            raise HTTPException(status_code=400, detail="config_json must be valid JSON") from exc
    else:
        parsed = config_json
    if not isinstance(parsed, dict):
        raise HTTPException(status_code=400, detail="config_json must be a JSON object")
    if "debug" in parsed and not isinstance(parsed["debug"], bool):
        raise HTTPException(status_code=400, detail="debug must be true or false")
    if "max_orders" in parsed and parsed["max_orders"] is not None:
        if not isinstance(parsed["max_orders"], int) or parsed["max_orders"] < 1:
            raise HTTPException(status_code=400, detail="max_orders must be a positive integer")
    if "packing_mode" in parsed and parsed["packing_mode"] not in {"normal", "fast"}:
        raise HTTPException(status_code=400, detail='packing_mode must be "normal" or "fast"')
    if "output_granularity" in parsed and parsed["output_granularity"] not in {"order_summary", "box_detail"}:
        raise HTTPException(status_code=400, detail='output_granularity must be "order_summary" or "box_detail"')
    if "sku_rules" in parsed and not isinstance(parsed["sku_rules"], dict):
        raise HTTPException(status_code=400, detail="sku_rules must be a JSON object")
    if "box_menu" in parsed and not isinstance(parsed["box_menu"], list):
        raise HTTPException(status_code=400, detail="box_menu must be a JSON array")
    if "order_rules" in parsed and not isinstance(parsed["order_rules"], list):
        raise HTTPException(status_code=400, detail="order_rules must be a JSON array")
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


def _work_dir(run_id: str) -> Path:
    work_dir = Path(tempfile.gettempdir()) / "box_optimizer_api" / run_id
    work_dir.mkdir(parents=True, exist_ok=True)
    return work_dir


def _save_base64_request_files(payload: Base64WorkbookRequest, work_dir: Path) -> tuple[Path, Path]:
    sku_suffix = _filename_suffix(payload.sku_master_filename, "sku_master_filename")
    orders_suffix = _filename_suffix(payload.orders_filename, "orders_filename")
    sku_master_path = _save_base64_file(
        payload.sku_master_base64,
        work_dir,
        f"sku_master{sku_suffix}",
        "sku_master_base64",
    )
    orders_path = _save_base64_file(
        payload.orders_base64,
        work_dir,
        f"orders{orders_suffix}",
        "orders_base64",
    )
    return sku_master_path, orders_path


@app.get("/health", response_model=HealthResponse)
def health() -> HealthResponse:
    """Health check endpoint."""
    return {"status": "ok"}


@app.get("/version", response_model=VersionResponse)
def version() -> VersionResponse:
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

        work_dir = _work_dir(run_id)

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


@app.post(
    "/inspect_base64",
    dependencies=[Depends(require_api_key)],
    response_model=InspectSummaryResponse,
)
def inspect_base64(payload: Base64WorkbookRequest):
    """Parse and match base64-encoded workbooks without packing or writing Excel."""
    started = time.perf_counter()
    stage = "request_received"
    run_id = uuid4().hex
    work_dir: Path | None = None
    try:
        _log_event("request_received", endpoint="/inspect_base64", run_id=run_id)
        stage = "config"
        config = _parse_config(payload.config_json)
        if config.get("debug"):
            logger.setLevel(logging.DEBUG)

        work_dir = _work_dir(run_id)
        _log_event(
            "filenames_received",
            endpoint="/inspect_base64",
            sku_master_file=payload.sku_master_filename,
            orders_file=payload.orders_filename,
        )
        stage = "files_saved"
        sku_master_path, orders_path = _save_base64_request_files(payload, work_dir)
        _log_event("files_saved", endpoint="/inspect_base64", run_id=run_id)

        stage = "inspect"
        result = inspect_workbook(
            sku_master_path=str(sku_master_path),
            orders_path=str(orders_path),
            config=config,
        )
        result["elapsed_seconds"] = round(time.perf_counter() - started, 3)
        _log_event(
            "response_returned",
            endpoint="/inspect_base64",
            run_id=run_id,
            elapsed_seconds=result["elapsed_seconds"],
        )
        return result
    except HTTPException as exc:
        return _json_error(exc, stage, status_code=exc.status_code)
    except Exception as exc:
        logger.exception("inspect_base64 failed", extra={"box_optimizer": {"stage": stage}})
        return _json_error(exc, stage)
    finally:
        if work_dir is not None:
            rmtree(work_dir, ignore_errors=True)


@app.post(
    "/optimize_base64",
    dependencies=[Depends(require_api_key)],
    response_model=OptimizeBase64Response,
)
def optimize_base64(payload: Base64WorkbookRequest):
    """Optimize base64-encoded workbooks and return a base64 XLSX workbook."""
    started = time.perf_counter()
    stage = "request_received"
    run_id = uuid4().hex
    work_dir: Path | None = None
    content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    output_filename = "optimized_shipping_plan.xlsx"
    try:
        _log_event("request_received", endpoint="/optimize_base64", run_id=run_id)
        stage = "config"
        config = _parse_config(payload.config_json)
        if config.get("debug"):
            logger.setLevel(logging.DEBUG)

        work_dir = _work_dir(run_id)
        _log_event(
            "filenames_received",
            endpoint="/optimize_base64",
            sku_master_file=payload.sku_master_filename,
            orders_file=payload.orders_filename,
        )
        stage = "files_saved"
        sku_master_path, orders_path = _save_base64_request_files(payload, work_dir)
        output_path = work_dir / output_filename
        _log_event("files_saved", endpoint="/optimize_base64", run_id=run_id)

        stage = "optimize"
        summary = optimize_workbook(
            sku_master_path=str(sku_master_path),
            orders_path=str(orders_path),
            output_path=str(output_path),
            config=config,
        )
        summary["elapsed_seconds"] = round(time.perf_counter() - started, 3)
        workbook_base64 = base64.b64encode(output_path.read_bytes()).decode("ascii")

        _log_event(
            "response_returned",
            endpoint="/optimize_base64",
            run_id=run_id,
            elapsed_seconds=summary["elapsed_seconds"],
        )
        return {
            "filename": output_filename,
            "content_type": content_type,
            "workbook_base64": workbook_base64,
            "summary": summary,
        }
    except HTTPException as exc:
        return _json_error(exc, stage, status_code=exc.status_code)
    except Exception as exc:
        logger.exception("optimize_base64 failed", extra={"box_optimizer": {"stage": stage}})
        return _json_error(exc, stage)
    finally:
        if work_dir is not None:
            rmtree(work_dir, ignore_errors=True)


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

        work_dir = _work_dir(run_id)

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
