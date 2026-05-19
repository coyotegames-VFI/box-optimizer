"""FastAPI app for box_optimizer."""

import base64
import binascii
import html
import logging
import json
import os
import re
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path
from shutil import copyfileobj, rmtree
from typing import Any
from urllib.parse import quote
from uuid import uuid4

from fastapi import Depends, FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
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

DEFAULT_UPLOAD_CONFIG = {
    "debug": True,
    "packing_mode": "fast",
    "output_granularity": "order_summary",
    "preserve_region_sheets": False,
}
JOB_ID_PATTERN = re.compile(r"^[a-f0-9]{32}$")
WORKBOOK_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

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


class JobStatusResponse(BaseModel):
    job_id: str
    status: str
    created_at: str
    expires_at: str
    summary: OptimizeSummaryResponse | None = None
    download_url: str | None = None
    error: str | None = None


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


def _upload_suffix_strict(upload: UploadFile, label: str) -> str:
    suffix = Path(upload.filename or "").suffix.lower()
    if suffix not in {".csv", ".xlsx"}:
        raise HTTPException(status_code=400, detail=f"{label} must be a .csv or .xlsx file")
    return suffix


def _jobs_root() -> Path:
    root = Path(os.getenv("BOX_OPTIMIZER_JOBS_DIR") or Path(tempfile.gettempdir()) / "box_optimizer_jobs")
    root.mkdir(parents=True, exist_ok=True)
    return root


def _job_dir(job_id: str) -> Path:
    if not JOB_ID_PATTERN.fullmatch(job_id):
        raise HTTPException(status_code=404, detail="Job not found")
    return _jobs_root() / job_id


def _job_ttl_seconds() -> float:
    raw = os.getenv("BOX_OPTIMIZER_JOB_TTL_HOURS", "24")
    try:
        hours = float(raw)
    except ValueError:
        hours = 24.0
    return max(hours, 1.0) * 3600


def _cleanup_expired_jobs() -> None:
    root = _jobs_root()
    cutoff = time.time() - _job_ttl_seconds()
    for child in root.iterdir():
        if not child.is_dir():
            continue
        try:
            if child.stat().st_mtime < cutoff:
                rmtree(child, ignore_errors=True)
        except OSError:
            continue


def _upload_access_token() -> str:
    return (os.getenv("BOX_OPTIMIZER_UPLOAD_TOKEN") or "").strip()


def _require_upload_access(upload_token: str | None = None, token: str | None = None) -> str:
    expected = _upload_access_token()
    provided = (upload_token or token or "").strip()
    if expected and provided != expected:
        raise HTTPException(status_code=403, detail="Invalid or missing upload access token")
    return provided


def _token_query(upload_token: str | None) -> str:
    expected = _upload_access_token()
    if not expected:
        return ""
    return f"?upload_token={quote(upload_token or '')}"


def _default_upload_config_text() -> str:
    return json.dumps(DEFAULT_UPLOAD_CONFIG, indent=2)


def _html_page(title: str, body: str, status_code: int = 200) -> HTMLResponse:
    return HTMLResponse(
        status_code=status_code,
        content=f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{html.escape(title)}</title>
  <style>
    body {{ font-family: Arial, sans-serif; margin: 2rem; max-width: 880px; line-height: 1.4; }}
    label {{ display: block; font-weight: 700; margin-top: 1rem; }}
    input[type=file], textarea {{ display: block; width: 100%; margin-top: .35rem; }}
    textarea {{ min-height: 9rem; font-family: Consolas, monospace; }}
    button, .button {{ display: inline-block; margin-top: 1.25rem; padding: .65rem 1rem; background: #174ea6; color: white; border: 0; border-radius: 4px; text-decoration: none; cursor: pointer; }}
    .summary {{ border: 1px solid #ddd; padding: 1rem; border-radius: 6px; background: #fafafa; }}
    .error {{ border: 1px solid #b00020; padding: 1rem; border-radius: 6px; background: #fff4f4; color: #7a0015; }}
    .muted {{ color: #555; }}
    li {{ margin-bottom: .25rem; }}
  </style>
</head>
<body>
{body}
</body>
</html>""",
    )


def _upload_form_html(config_text: str, upload_token: str | None = None, error: str | None = None) -> HTMLResponse:
    safe_config = html.escape(config_text or _default_upload_config_text())
    hidden_token = html.escape(upload_token or "")
    error_html = f'<div class="error"><strong>Something needs attention:</strong> {html.escape(error)}</div>' if error else ""
    return _html_page(
        "Box Optimizer Upload",
        f"""
<h1>Box Optimizer Upload</h1>
<p class="muted">Upload the two campaign files, then run the optimizer. The file names do not need to follow a special pattern.</p>
{error_html}
<form action="/upload" method="post" enctype="multipart/form-data">
  <input type="hidden" name="upload_token" value="{hidden_token}">
  <label for="sku_master_file">Put your SKU file here</label>
  <input id="sku_master_file" name="sku_master_file" type="file" accept=".csv,.xlsx" required>

  <label for="orders_file">Put your orders file here</label>
  <input id="orders_file" name="orders_file" type="file" accept=".csv,.xlsx" required>

  <details>
    <summary>Advanced settings</summary>
    <label for="config_json">Optional packing instructions generated by GPT</label>
    <textarea id="config_json" name="config_json">{safe_config}</textarea>
  </details>

  <button type="submit">Run Optimization</button>
</form>
""",
    )


def _summary_value(summary: dict, key: str) -> str:
    value = summary.get(key)
    return html.escape(str(value if value is not None else ""))


def _result_page(record: dict, upload_token: str | None = None, status_code: int = 200) -> HTMLResponse:
    summary = record.get("summary") or {}
    warnings = summary.get("warnings") or []
    warning_items = "".join(f"<li>{html.escape(str(warning))}</li>" for warning in warnings) or "<li>No workflow warnings.</li>"
    download_url = f"/jobs/{record['job_id']}/download{_token_query(upload_token)}"
    status_url = f"/jobs/{record['job_id']}{_token_query(upload_token)}"
    return _html_page(
        "Box Optimizer Results",
        f"""
<h1>Optimization Results</h1>
<div class="summary">
  <p><strong>Job ID:</strong> {html.escape(record['job_id'])}</p>
  <p><strong>Orders processed:</strong> {_summary_value(summary, 'orders_processed')}</p>
  <p><strong>Boxes created:</strong> {_summary_value(summary, 'boxes_created')}</p>
  <p><strong>Box types used:</strong> {_summary_value(summary, 'box_types')}</p>
  <p><strong>Unmatched SKUs:</strong> {_summary_value(summary, 'unmatched_skus')}</p>
  <p><strong>Warning count:</strong> {_summary_value(summary, 'warning_count')}</p>
  <a class="button" href="{html.escape(download_url)}">Download optimized workbook</a>
  <p><a href="{html.escape(status_url)}">View job status as JSON</a></p>
</div>
<h2>Warnings</h2>
<ul>{warning_items}</ul>
""",
        status_code=status_code,
    )


def _error_page(message: str, upload_token: str | None = None, status_code: int = 400) -> HTMLResponse:
    return _html_page(
        "Box Optimizer Error",
        f"""
<h1>Optimization could not finish</h1>
<div class="error">{html.escape(message)}</div>
<p><a href="/upload{html.escape(_token_query(upload_token))}">Return to upload page</a></p>
""",
        status_code=status_code,
    )


def _write_job_record(job_dir: Path, record: dict) -> None:
    (job_dir / "summary.json").write_text(json.dumps(record, indent=2), encoding="utf-8")


def _read_job_record(job_id: str) -> dict:
    job_dir = _job_dir(job_id)
    summary_path = job_dir / "summary.json"
    if not summary_path.exists():
        raise HTTPException(status_code=404, detail="Job not found")
    return json.loads(summary_path.read_text(encoding="utf-8"))


def _job_expiration(created_at_seconds: float) -> str:
    return datetime.fromtimestamp(created_at_seconds + _job_ttl_seconds(), timezone.utc).isoformat()


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


@app.get("/upload", response_class=HTMLResponse, include_in_schema=False)
def upload_page(
    upload_token: str | None = None,
    token: str | None = None,
    job_config: str | None = None,
):
    """Return an employee-friendly upload page for daily optimizer use."""
    try:
        provided_token = _require_upload_access(upload_token=upload_token, token=token)
    except HTTPException as exc:
        return _error_page(str(exc.detail), status_code=exc.status_code)
    config_text = job_config or _default_upload_config_text()
    return _upload_form_html(config_text=config_text, upload_token=provided_token)


@app.post("/upload", response_class=HTMLResponse, include_in_schema=False)
def upload_workbooks(
    sku_master_file: UploadFile = File(...),
    orders_file: UploadFile = File(...),
    config_json: str | None = Form(default=None),
    upload_token: str | None = Form(default=None),
):
    """Run the existing optimizer from the employee upload page and return an HTML result."""
    started = time.perf_counter()
    job_id = uuid4().hex
    created_seconds = time.time()
    job_dir = _jobs_root() / job_id
    try:
        provided_token = _require_upload_access(upload_token=upload_token)
    except HTTPException as exc:
        return _error_page(str(exc.detail), status_code=exc.status_code)

    try:
        _cleanup_expired_jobs()
        job_dir.mkdir(parents=True, exist_ok=False)
        config = _parse_config(config_json or DEFAULT_UPLOAD_CONFIG)
        if config.get("debug"):
            logger.setLevel(logging.DEBUG)

        sku_master_path = _save_upload(
            sku_master_file,
            job_dir,
            f"sku_master{_upload_suffix_strict(sku_master_file, 'SKU file')}",
        )
        orders_path = _save_upload(
            orders_file,
            job_dir,
            f"orders{_upload_suffix_strict(orders_file, 'Orders file')}",
        )
        output_path = job_dir / "optimized_shipping_plan.xlsx"
        summary = optimize_workbook(
            sku_master_path=str(sku_master_path),
            orders_path=str(orders_path),
            output_path=str(output_path),
            config=config,
        )
        summary["elapsed_seconds"] = round(time.perf_counter() - started, 3)
        record = {
            "job_id": job_id,
            "status": "completed",
            "created_at": datetime.fromtimestamp(created_seconds, timezone.utc).isoformat(),
            "expires_at": _job_expiration(created_seconds),
            "summary": summary,
            "download_url": f"/jobs/{job_id}/download",
            "error": None,
        }
        _write_job_record(job_dir, record)
        return _result_page(record, upload_token=provided_token)
    except HTTPException as exc:
        record = {
            "job_id": job_id,
            "status": "failed",
            "created_at": datetime.fromtimestamp(created_seconds, timezone.utc).isoformat(),
            "expires_at": _job_expiration(created_seconds),
            "summary": None,
            "download_url": None,
            "error": str(exc.detail),
        }
        job_dir.mkdir(parents=True, exist_ok=True)
        _write_job_record(job_dir, record)
        return _error_page(str(exc.detail), upload_token=provided_token, status_code=exc.status_code)
    except Exception as exc:
        logger.exception("upload optimization failed")
        message = f"The optimizer could not finish this job: {exc}"
        record = {
            "job_id": job_id,
            "status": "failed",
            "created_at": datetime.fromtimestamp(created_seconds, timezone.utc).isoformat(),
            "expires_at": _job_expiration(created_seconds),
            "summary": None,
            "download_url": None,
            "error": message,
        }
        job_dir.mkdir(parents=True, exist_ok=True)
        _write_job_record(job_dir, record)
        return _error_page(message, upload_token=provided_token, status_code=500)


@app.get("/jobs/{job_id}", response_model=JobStatusResponse)
def job_status(job_id: str, upload_token: str | None = None, token: str | None = None):
    """Return compact status and summary for an upload job."""
    _require_upload_access(upload_token=upload_token, token=token)
    _cleanup_expired_jobs()
    return _read_job_record(job_id)


@app.get("/jobs/{job_id}/download")
def job_download(job_id: str, upload_token: str | None = None, token: str | None = None) -> FileResponse:
    """Download the optimized XLSX output for an upload job."""
    _require_upload_access(upload_token=upload_token, token=token)
    _cleanup_expired_jobs()
    job_dir = _job_dir(job_id)
    record = _read_job_record(job_id)
    if record.get("status") != "completed":
        raise HTTPException(status_code=409, detail="Job is not complete")
    output_path = job_dir / "optimized_shipping_plan.xlsx"
    if not output_path.exists():
        raise HTTPException(status_code=404, detail="Optimized workbook not found")
    return FileResponse(
        path=str(output_path),
        media_type=WORKBOOK_CONTENT_TYPE,
        filename="optimized_shipping_plan.xlsx",
    )


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
    content_type = WORKBOOK_CONTENT_TYPE
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
            media_type=WORKBOOK_CONTENT_TYPE,
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
