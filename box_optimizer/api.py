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
from urllib.parse import quote, urlencode
from uuid import uuid4

from fastapi import Depends, FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse
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
from box_optimizer.rate_sources import (
    ACTIVE_RATE_SHEET_FILENAME,
    RATE_ADMIN_TOKEN_ENV,
    RATE_SYNC_TOKEN_ENV,
    RateSheetValidationError,
    active_rate_sheet_path,
    rate_sheet_metadata,
    rate_sheet_metadata_path,
    rate_sheet_root,
    rate_sheet_status,
    sha256_file,
    validate_rate_sheet,
)
from box_optimizer.workflow import inspect_workbook, optimize_workbook


logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("box_optimizer")

DEFAULT_UPLOAD_CONFIG = {
    "debug": True,
    "packing_mode": "fast",
    "output_granularity": "order_summary",
    "preserve_region_sheets": False,
}
ADMIN_UPLOAD_TOKEN_ENV = "BOX_OPTIMIZER_ADMIN_UPLOAD_TOKEN"
POWER_UPLOAD_CONFIG = {
    "debug": True,
    "packing_mode": "balanced",
    "max_optimization_seconds": 300,
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
    chargeable_weight_plan_selected_count: int | None = None
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
    if "packing_mode" in parsed and parsed["packing_mode"] not in {"normal", "fast", "balanced"}:
        raise HTTPException(status_code=400, detail='packing_mode must be "normal", "fast", or "balanced"')
    if "max_optimization_seconds" in parsed and parsed["max_optimization_seconds"] is not None:
        if not isinstance(parsed["max_optimization_seconds"], (int, float)) or parsed["max_optimization_seconds"] <= 0:
            raise HTTPException(status_code=400, detail="max_optimization_seconds must be a positive number")
    for field in ["balanced_max_items_for_deep_search", "balanced_max_item_quantity_for_recombine"]:
        if field in parsed and parsed[field] is not None:
            if not isinstance(parsed[field], int) or parsed[field] < 1:
                raise HTTPException(status_code=400, detail=f"{field} must be a positive integer")
    if "balanced_min_remaining_seconds" in parsed and parsed["balanced_min_remaining_seconds"] is not None:
        if not isinstance(parsed["balanced_min_remaining_seconds"], (int, float)) or parsed["balanced_min_remaining_seconds"] < 0:
            raise HTTPException(status_code=400, detail="balanced_min_remaining_seconds must be a non-negative number")
    if "bundle_footprint_tolerance_cm" in parsed and parsed["bundle_footprint_tolerance_cm"] is not None:
        if not isinstance(parsed["bundle_footprint_tolerance_cm"], (int, float)) or parsed["bundle_footprint_tolerance_cm"] < 0:
            raise HTTPException(status_code=400, detail="bundle_footprint_tolerance_cm must be a non-negative number")
    if "chargeable_weight_split_savings_threshold_kg" in parsed and parsed["chargeable_weight_split_savings_threshold_kg"] is not None:
        if not isinstance(parsed["chargeable_weight_split_savings_threshold_kg"], (int, float)) or parsed["chargeable_weight_split_savings_threshold_kg"] < 0:
            raise HTTPException(status_code=400, detail="chargeable_weight_split_savings_threshold_kg must be a non-negative number")
    if "chargeable_weight_split_savings_threshold_pct" in parsed and parsed["chargeable_weight_split_savings_threshold_pct"] is not None:
        if not isinstance(parsed["chargeable_weight_split_savings_threshold_pct"], (int, float)) or parsed["chargeable_weight_split_savings_threshold_pct"] < 0:
            raise HTTPException(status_code=400, detail="chargeable_weight_split_savings_threshold_pct must be a non-negative number")
    if "chargeable_weight_split_two_extra_box_threshold_kg" in parsed and parsed["chargeable_weight_split_two_extra_box_threshold_kg"] is not None:
        if not isinstance(parsed["chargeable_weight_split_two_extra_box_threshold_kg"], (int, float)) or parsed["chargeable_weight_split_two_extra_box_threshold_kg"] < 0:
            raise HTTPException(status_code=400, detail="chargeable_weight_split_two_extra_box_threshold_kg must be a non-negative number")
    if "max_extra_boxes_per_order" in parsed and parsed["max_extra_boxes_per_order"] is not None:
        if not isinstance(parsed["max_extra_boxes_per_order"], int) or parsed["max_extra_boxes_per_order"] < 0:
            raise HTTPException(status_code=400, detail="max_extra_boxes_per_order must be a non-negative integer")
    if "oversized_vendor_box_ids" in parsed and parsed["oversized_vendor_box_ids"] is not None:
        if not isinstance(parsed["oversized_vendor_box_ids"], list):
            raise HTTPException(status_code=400, detail="oversized_vendor_box_ids must be a JSON array")
    if "oversized_vendor_box_chargeable_threshold_kg" in parsed and parsed["oversized_vendor_box_chargeable_threshold_kg"] is not None:
        if not isinstance(parsed["oversized_vendor_box_chargeable_threshold_kg"], (int, float)) or parsed["oversized_vendor_box_chargeable_threshold_kg"] < 0:
            raise HTTPException(status_code=400, detail="oversized_vendor_box_chargeable_threshold_kg must be a non-negative number")
    if "oversized_max_extra_boxes_per_order" in parsed and parsed["oversized_max_extra_boxes_per_order"] is not None:
        if not isinstance(parsed["oversized_max_extra_boxes_per_order"], int) or parsed["oversized_max_extra_boxes_per_order"] < 0:
            raise HTTPException(status_code=400, detail="oversized_max_extra_boxes_per_order must be a non-negative integer")
    for field in [
        "non_preferred_extra_box_savings_threshold_kg",
        "non_preferred_extra_box_savings_threshold_pct",
        "non_preferred_two_extra_box_savings_threshold_kg",
        "non_preferred_two_extra_box_savings_threshold_pct",
        "company_protection_max_rate_weight_kg",
        "company_protection_min_margin_delta",
        "repeat_retail_min_optimization_seconds",
        "repeat_retail_min_savings_threshold_kg",
        "repeat_retail_min_savings_threshold_pct",
        "repeat_retail_max_margin_giveback",
        "repeat_retail_min_customer_savings",
        "vendor_box_fit_tolerance_cm",
        "vendor_box_fit_tolerance_max_cm",
        "vendor_box_fit_tolerance_max_chargeable_increase_kg",
    ]:
        if field in parsed and parsed[field] is not None:
            if not isinstance(parsed[field], (int, float)) or parsed[field] < 0:
                raise HTTPException(status_code=400, detail=f"{field} must be a non-negative number")
    if "vendor_box_fit_tolerance_guardrail" in parsed and not isinstance(parsed["vendor_box_fit_tolerance_guardrail"], bool):
        raise HTTPException(status_code=400, detail="vendor_box_fit_tolerance_guardrail must be true or false")
    if "vendor_box_fit_mode" in parsed and parsed["vendor_box_fit_mode"] not in {"auto", "off", "on"}:
        raise HTTPException(status_code=400, detail='vendor_box_fit_mode must be "auto", "off", or "on"')
    if (
        "vendor_box_fit_tolerance_max_cm" in parsed
        and parsed["vendor_box_fit_tolerance_max_cm"] is not None
        and parsed["vendor_box_fit_tolerance_max_cm"] > 2
    ):
        raise HTTPException(status_code=400, detail="vendor_box_fit_tolerance_max_cm must be 2 cm or less")
    for field in [
        "repeat_retail_min_repeated_units",
        "repeat_retail_max_extra_boxes_per_order",
        "repeat_retail_max_candidate_boxes",
    ]:
        if field in parsed and parsed[field] is not None:
            if not isinstance(parsed[field], int) or parsed[field] < 0:
                raise HTTPException(status_code=400, detail=f"{field} must be a non-negative integer")
    if "repeat_retail_batch_planning_enabled" in parsed and not isinstance(parsed["repeat_retail_batch_planning_enabled"], bool):
        raise HTTPException(status_code=400, detail="repeat_retail_batch_planning_enabled must be true or false")
    if "repeat_retail_batch_sizes" in parsed and parsed["repeat_retail_batch_sizes"] is not None:
        if (
            not isinstance(parsed["repeat_retail_batch_sizes"], list)
            or any(not isinstance(size, int) or size < 1 for size in parsed["repeat_retail_batch_sizes"])
        ):
            raise HTTPException(status_code=400, detail="repeat_retail_batch_sizes must be a JSON array of positive integers")
    if "company_protection_extra_box_guardrail" in parsed and not isinstance(parsed["company_protection_extra_box_guardrail"], bool):
        raise HTTPException(status_code=400, detail="company_protection_extra_box_guardrail must be true or false")
    if "company_protection_rate_bands" in parsed and parsed["company_protection_rate_bands"] is not None:
        if not isinstance(parsed["company_protection_rate_bands"], dict):
            raise HTTPException(status_code=400, detail="company_protection_rate_bands must be a JSON object")
    if "company_protection_zone_markups" in parsed and parsed["company_protection_zone_markups"] is not None:
        if not isinstance(parsed["company_protection_zone_markups"], dict):
            raise HTTPException(status_code=400, detail="company_protection_zone_markups must be a JSON object")
    if "company_protection_country_zones" in parsed and parsed["company_protection_country_zones"] is not None:
        if not isinstance(parsed["company_protection_country_zones"], dict):
            raise HTTPException(status_code=400, detail="company_protection_country_zones must be a JSON object")
    if "output_granularity" in parsed and parsed["output_granularity"] not in {"order_summary", "box_detail"}:
        raise HTTPException(status_code=400, detail='output_granularity must be "order_summary" or "box_detail"')
    if "sku_rules" in parsed and not isinstance(parsed["sku_rules"], dict):
        raise HTTPException(status_code=400, detail="sku_rules must be a JSON object")
    if "separate_playmat_charge_skus" in parsed and parsed["separate_playmat_charge_skus"] is not None:
        if not isinstance(parsed["separate_playmat_charge_skus"], list):
            raise HTTPException(status_code=400, detail="separate_playmat_charge_skus must be a JSON array")
    if "box_menu" in parsed and not isinstance(parsed["box_menu"], list):
        raise HTTPException(status_code=400, detail="box_menu must be a JSON array")
    if "order_rules" in parsed and not isinstance(parsed["order_rules"], list):
        raise HTTPException(status_code=400, detail="order_rules must be a JSON array")
    return parsed


def _campaign_download_filename(config: dict) -> str:
    campaign = config.get("campaign") if isinstance(config.get("campaign"), dict) else {}
    raw_name = str(
        campaign.get("name")
        or campaign.get("project_name")
        or campaign.get("short_name")
        or campaign.get("code")
        or ""
    ).strip()
    safe_name = re.sub(r'[\\/:*?"<>|]+', "", raw_name)
    safe_name = re.sub(r"\s+", " ", safe_name).strip()
    generated_date = datetime.now().date().isoformat()
    return f"{safe_name} - {generated_date}.xlsx" if safe_name else "optimized_shipping_plan.xlsx"


def _form_lines(value: str | None) -> list[str]:
    return [line.strip() for line in (value or "").splitlines() if line.strip()]


def _form_skus(value: str | None) -> list[str]:
    normalized = []
    seen = set()
    for part in re.split(r"[\r\n,]+", value or ""):
        sku = normalize_sku(part)
        if sku and sku not in seen:
            normalized.append(sku)
            seen.add(sku)
    return normalized


def _form_float(value: str | float | int | None, default: float) -> float:
    if value in (None, ""):
        return default
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def _base_config_for_packing_choice(choice: str | None) -> dict[str, Any]:
    mode = (choice or "railway_fast").strip()
    if mode == "local_power_300":
        mode = "local_power_balanced_300"
    config: dict[str, Any] = {
        "debug": True,
        "output_granularity": "order_summary",
        "preserve_region_sheets": False,
    }
    if mode == "railway_balanced_30":
        config.update({"packing_mode": "balanced", "max_optimization_seconds": 30})
    elif mode == "railway_balanced_60":
        config.update({"packing_mode": "balanced", "max_optimization_seconds": 60})
    elif mode == "local_power_balanced_300":
        config.update({"packing_mode": "balanced", "max_optimization_seconds": 300})
    else:
        config["packing_mode"] = "fast"
    return config


def _invoice_config_for_selection(selection: str | None) -> dict[str, Any]:
    value = (selection or "none").strip()
    if value == "invoice_us":
        return {"include_invoice": True, "invoice_variant": "US"}
    if value == "invoice_cn":
        return {"include_invoice": True, "invoice_variant": "CN"}
    return {"include_invoice": False}


def _merge_config(base: dict[str, Any], override: dict[str, Any]) -> dict[str, Any]:
    merged = dict(base)
    for key, value in override.items():
        if isinstance(value, dict) and isinstance(merged.get(key), dict):
            merged[key] = _merge_config(merged[key], value)
        else:
            merged[key] = value
    return merged


def _structured_upload_config(
    *,
    campaign_name: str | None = None,
    campaign_code: str | None = None,
    campaign_notes: str | None = None,
    packing_mode_choice: str | None = None,
    ship_as_is_skus: str | None = None,
    ship_as_is_exception_skus: str | None = None,
    ship_as_is_box_type: str | None = None,
    separate_playmat_charge_skus: str | None = None,
    no_padding_skus: str | None = None,
    wrap_around_skus: str | None = None,
    wrapped_height_cm: str | float | int | None = None,
    compressible_skus: str | None = None,
    compressed_height_ratio: str | float | int | None = None,
    compressed_volume_ratio: str | float | int | None = None,
    invoice_selection: str | None = None,
) -> dict[str, Any]:
    config = _base_config_for_packing_choice(packing_mode_choice)
    config.update(_invoice_config_for_selection(invoice_selection))
    campaign = {}
    if campaign_name and campaign_name.strip():
        campaign["name"] = campaign_name.strip()
    if campaign_code and campaign_code.strip():
        campaign["code"] = campaign_code.strip()
    if campaign_notes and campaign_notes.strip():
        campaign["notes"] = campaign_notes.strip()
    if campaign:
        config["campaign"] = campaign

    sku_rules: dict[str, dict[str, Any]] = {}
    ship_as_is_label = str(ship_as_is_box_type or "").strip()
    for sku in _form_lines(ship_as_is_skus):
        rule = {
            "prepacked": True,
            "no_padding": True,
            "ships_alone": True,
            "can_mix_with_other_items": False,
            "box_type": f"{sku} shipping carton",
        }
        if ship_as_is_label:
            rule["label_box_type"] = ship_as_is_label
        sku_rules[sku] = rule
    exception_skus = _form_skus(ship_as_is_exception_skus)
    if exception_skus:
        config["ship_as_is_exception_skus"] = exception_skus
    playmat_skus = _form_lines(separate_playmat_charge_skus)
    if playmat_skus:
        config["separate_playmat_charge_skus"] = playmat_skus
        for sku in playmat_skus:
            sku_rules[sku] = _merge_config(
                sku_rules.get(sku, {}),
                {
                    "prepacked": True,
                    "no_padding": True,
                    "ships_alone": True,
                    "can_mix_with_other_items": False,
                    "separate_playmat_charge": True,
                    "box_type": f"{sku} separate playmat parcel",
                },
            )
    for sku in _form_lines(no_padding_skus):
        sku_rules[sku] = _merge_config(sku_rules.get(sku, {}), {"no_padding": True})
    wrapped_height = _form_float(wrapped_height_cm, 4.0)
    for sku in _form_lines(wrap_around_skus):
        sku_rules[sku] = _merge_config(
            sku_rules.get(sku, {}),
            {
                "wrap_around_largest_item": True,
                "wrapped_height_cm": wrapped_height,
                "no_padding": True,
            },
        )
    height_ratio = _form_float(compressed_height_ratio, 0.6)
    volume_ratio = _form_float(compressed_volume_ratio, 0.75)
    for sku in _form_lines(compressible_skus):
        sku_rules[sku] = _merge_config(
            sku_rules.get(sku, {}),
            {
                "compressible": True,
                "compressed_height_ratio": height_ratio,
                "compressed_volume_ratio": volume_ratio,
            },
        )
    if sku_rules:
        config["sku_rules"] = sku_rules
    return config


def _final_upload_config(manual_config_json: str | None, structured_config: dict[str, Any]) -> dict[str, Any]:
    manual_config = _parse_config(manual_config_json) if manual_config_json and manual_config_json.strip() else {}
    return _parse_config(_merge_config(structured_config, manual_config))

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


def _rate_sheet_root() -> Path:
    return rate_sheet_root()


def _active_rate_sheet_path() -> Path:
    return active_rate_sheet_path()


def _rate_sheet_metadata_path() -> Path:
    return rate_sheet_metadata_path()


def _rate_sheet_metadata() -> dict[str, Any]:
    return rate_sheet_metadata()


def _sha256_file(path: Path) -> str:
    return sha256_file(path)


def _rate_sheet_status() -> dict[str, Any]:
    return rate_sheet_status()


def _rate_sheet_status_html() -> str:
    status = _rate_sheet_status()
    loaded_at = f"<br>Last loaded: {html.escape(status['uploaded_at'])}" if status["uploaded_at"] else ""
    checksum = f"<br>Checksum: <code>{html.escape(status.get('checksum_short') or '')}</code>" if status.get("checksum_short") else ""
    filename = f"<br>Filename: {html.escape(status['filename'])}" if status["filename"] else ""
    return f"<p class=\"help\"><strong>{html.escape(status['message'])}</strong>{filename}{loaded_at}{checksum}</p>"


def _validate_rate_sheet(path: Path) -> dict[str, Any]:
    try:
        return validate_rate_sheet(path)
    except RateSheetValidationError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


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


def _admin_upload_token() -> str:
    return (os.getenv(ADMIN_UPLOAD_TOKEN_ENV) or "").strip()


def _rate_admin_token() -> str:
    return (os.getenv(RATE_ADMIN_TOKEN_ENV) or "").strip()


def _rate_sync_token() -> str:
    return (os.getenv(RATE_SYNC_TOKEN_ENV) or "").strip()


def _require_upload_access(upload_token: str | None = None, token: str | None = None) -> str:
    allowed_tokens = [value for value in [_upload_access_token(), _admin_upload_token()] if value]
    provided = (upload_token or token or "").strip()
    if allowed_tokens and provided not in allowed_tokens:
        raise HTTPException(status_code=403, detail="Invalid or missing upload access token")
    return provided


def _workbook_output_mode_for_token(provided_token: str) -> str:
    admin_token = _admin_upload_token()
    if admin_token:
        return "admin" if provided_token == admin_token else "worker"
    return "admin"


def _require_rate_upload_access(
    upload_token: str | None = None,
    rate_admin_token: str | None = None,
    token: str | None = None,
) -> str:
    admin_expected = _rate_admin_token()
    provided = (rate_admin_token or token or upload_token or "").strip()
    if admin_expected:
        if provided != admin_expected:
            raise HTTPException(status_code=403, detail="Invalid or missing rate admin token")
        return provided
    return _require_upload_access(upload_token=upload_token, token=token)


def _require_rate_download_access(upload_token: str | None = None, token: str | None = None) -> str:
    provided = (upload_token or token or "").strip()
    allowed_tokens = [value for value in [_rate_sync_token(), _rate_admin_token()] if value]
    if allowed_tokens:
        if provided not in allowed_tokens:
            raise HTTPException(status_code=403, detail="Invalid or missing rate sheet download token")
        return provided
    return _require_upload_access(upload_token=upload_token, token=token)


def _token_query(upload_token: str | None) -> str:
    if not (_upload_access_token() or _admin_upload_token()):
        return ""
    return f"?upload_token={quote(upload_token or '')}"


def _default_upload_config_text() -> str:
    return json.dumps(DEFAULT_UPLOAD_CONFIG, indent=2)


def _power_upload_config_text() -> str:
    return json.dumps(POWER_UPLOAD_CONFIG, indent=2)


def _power_upload_enabled() -> bool:
    return os.getenv("BOX_OPTIMIZER_ENABLE_POWER_UPLOAD", "").strip().lower() in {"1", "true", "yes", "on"}


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
    body {{ font-family: Arial, sans-serif; margin: 2rem; max-width: 980px; line-height: 1.4; }}
    label {{ display: block; font-weight: 700; margin-top: 1rem; }}
    input[type=file], input[type=text], input[type=number], select, textarea {{ display: block; width: 100%; margin-top: .35rem; box-sizing: border-box; }}
    input[type=checkbox] {{ margin-right: .35rem; }}
    textarea {{ min-height: 9rem; font-family: Consolas, monospace; }}
    .short {{ max-width: 14rem; }}
    button, .button {{ display: inline-block; margin-top: 1.25rem; padding: .65rem 1rem; background: #174ea6; color: white; border: 0; border-radius: 4px; text-decoration: none; cursor: pointer; }}
    .summary {{ border: 1px solid #ddd; padding: 1rem; border-radius: 6px; background: #fafafa; }}
    .section {{ border-top: 1px solid #ddd; margin-top: 1.25rem; padding-top: 1rem; }}
    .error {{ border: 1px solid #b00020; padding: 1rem; border-radius: 6px; background: #fff4f4; color: #7a0015; }}
    .hidden {{ display: none; }}
    .muted {{ color: #555; }}
    .help {{ color: #555; font-size: .95rem; margin: .25rem 0 .75rem; }}
    .inline-label {{ display: flex; align-items: center; gap: .35rem; font-weight: 400; margin-top: .5rem; }}
    .grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 1rem; }}
    li {{ margin-bottom: .25rem; }}
  </style>
</head>
<body>
{body}
</body>
</html>""",
    )


def _campaign_intake_html(default_packing_mode_choice: str = "railway_fast") -> str:
    selected = {
        "railway_fast": "",
        "railway_balanced_30": "",
        "railway_balanced_60": "",
        "local_power_balanced_300": "",
    }
    if default_packing_mode_choice == "local_power_300":
        default_packing_mode_choice = "local_power_balanced_300"
    selected[default_packing_mode_choice if default_packing_mode_choice in selected else "railway_fast"] = " selected"
    html_text = """
  <div class="section" id="campaign_intake">
    <h2>Campaign setup</h2>
    <p class="help">Use these fields for common campaign rules. The page will generate Advanced settings JSON below. SKU entries must match standardized uploaded SKUs or item names; language variants must be listed as separate SKUs.</p>

    <div class="grid">
      <div>
        <label for="campaign_name">Campaign Name</label>
        <input id="campaign_name" name="campaign_name" type="text" placeholder="Example: Vanguard US">
      </div>
      <div>
        <label for="campaign_code">Campaign Code / Project Name</label>
        <input id="campaign_code" name="campaign_code" type="text" placeholder="Optional">
      </div>
    </div>
    <label for="campaign_notes">Notes</label>
    <textarea id="campaign_notes" name="campaign_notes" placeholder="Optional internal notes for this run."></textarea>

    <h3>Packing mode</h3>
    <label for="packing_mode_choice">Run type</label>
    <select id="packing_mode_choice" name="packing_mode_choice">
      <option value="railway_fast"__SEL_RAILWAY_FAST__>Railway Fast - recommended for daily team use</option>
      <option value="railway_balanced_30"__SEL_RAILWAY_BALANCED_30__>Railway Balanced - 30 second accuracy mode</option>
      <option value="railway_balanced_60"__SEL_RAILWAY_BALANCED_60__>Railway Balanced - 60 second accuracy mode</option>
      <option value="local_power_balanced_300"__SEL_LOCAL_POWER_300__>Local Power Balanced - 300 second local run</option>
    </select>
    <p class="help">Local Power Balanced is intended for a local machine. The optimization seconds value is a search budget, not a guaranteed total page runtime.</p>

    <h3>Common SKU rules</h3>
    <p class="help">Enter one standardized SKU or exact item name per line. Do not combine English/German versions under one rule unless the uploaded file uses the same SKU for both.</p>

    <label for="ship_as_is_skus">Ship-as-is / do not touch packages</label>
    <p class="help">Use when the item is already in its final shipping carton. The optimizer will not place it inside another box or combine it with add-ons.</p>
    <textarea id="ship_as_is_skus" name="ship_as_is_skus" placeholder="SKU-001&#10;SKU-002"></textarea>
    <label for="ship_as_is_box_type">Friendly box type label for ship-as-is items</label>
    <input id="ship_as_is_box_type" name="ship_as_is_box_type" type="text" placeholder="Example: Final factory shipping carton">
    <label for="ship_as_is_exception_skus">Ship as Is Exceptions</label>
    <p class="help">Small pre-approved SKUs that may be inserted into a single ship-as-is carton when all other order items are listed here.</p>
    <textarea id="ship_as_is_exception_skus" name="ship_as_is_exception_skus" placeholder="PROMO-CARD, THANK-YOU-NOTE"></textarea>

    <label for="separate_playmat_charge_skus">Separate Playmat Charge SKUs</label>
    <p class="help">Use for playmat SKUs that ship as their own no-touch parcel and receive the fixed $6 package weight charge per unit.</p>
    <textarea id="separate_playmat_charge_skus" name="separate_playmat_charge_skus" placeholder="PLAYMAT-001&#10;PLAYMAT-002"></textarea>

    <label for="no_padding_skus">No-padding items</label>
    <p class="help">Use when the item does not need item-level padding. It may still share a box with other items unless another rule prevents mixing.</p>
    <textarea id="no_padding_skus" name="no_padding_skus" placeholder="SKU-003"></textarea>

    <label for="wrap_skus">Foldable / wrap-around items</label>
    <p class="help">Use for playmats or similar flat items that can fold or wrap around the largest item in the order. These items also receive no item-level padding by default. This reduces the item footprint and uses a wrapped height.</p>
    <textarea id="wrap_skus" name="wrap_around_skus" placeholder="PLAYMAT-001"></textarea>
    <label for="wrapped_height_cm">Wrapped height cm</label>
    <input id="wrapped_height_cm" name="wrapped_height_cm" class="short" type="number" min="1" step="0.5" value="4">
    <label class="inline-label" for="wrap_no_padding"><input id="wrap_no_padding" name="wrap_no_padding" type="checkbox" checked> Do not add padding to foldable / wrap-around items</label>

    <label for="compressible_skus">Compressible items</label>
    <p class="help">Use for soft goods such as plush, blankets, fabric, or foam items that can safely compress. This reduces modeled packed size and may lower dimensional weight.</p>
    <textarea id="compressible_skus" name="compressible_skus" placeholder="PLUSH-001"></textarea>
    <div class="grid">
      <div>
        <label for="compressed_height_ratio">Compressed height ratio</label>
        <input id="compressed_height_ratio" name="compressed_height_ratio" class="short" type="number" min="0.1" max="1" step="0.05" value="0.6">
      </div>
      <div>
        <label for="compressed_volume_ratio">Compressed volume ratio</label>
        <input id="compressed_volume_ratio" name="compressed_volume_ratio" class="short" type="number" min="0.1" max="1" step="0.05" value="0.75">
      </div>
    </div>


    <button type="button" id="generate_config_button">Update Advanced settings from form</button>
  </div>
  <script>
    (function () {
      function lines(id) {
        return (document.getElementById(id).value || "")
          .split(/\r?\n/)
          .map(function (line) { return line.trim(); })
          .filter(Boolean);
      }

      function skuValues(id) {
        var seen = {};
        return (document.getElementById(id).value || "")
          .split(/[\r\n,]+/)
          .map(function (line) { return line.trim().toUpperCase(); })
          .filter(function (sku) {
            if (!sku || seen[sku]) return false;
            seen[sku] = true;
            return true;
          });
      }

      function numberValue(id, fallback) {
        var value = parseFloat(document.getElementById(id).value);
        return Number.isFinite(value) ? value : fallback;
      }

      function applyRule(config, sku, rule) {
        config.sku_rules = config.sku_rules || {};
        config.sku_rules[sku] = Object.assign({}, config.sku_rules[sku] || {}, rule);
      }

      function applyPackingMode(config) {
        var choice = document.getElementById("packing_mode_choice").value;
        config.debug = true;
        config.output_granularity = "order_summary";
        config.preserve_region_sheets = false;
        if (choice === "railway_fast") {
          config.packing_mode = "fast";
          delete config.max_optimization_seconds;
          return;
        }
        config.packing_mode = "balanced";
        if (choice === "railway_balanced_30") {
          config.max_optimization_seconds = 30;
        } else if (choice === "railway_balanced_60") {
          config.max_optimization_seconds = 60;
        } else {
          config.max_optimization_seconds = 300;
        }
      }

      function showConfigError(message) {
        var errorBox = document.getElementById("config_generation_error");
        if (!errorBox) return;
        errorBox.textContent = message;
        errorBox.classList.remove("hidden");
      }

      function clearConfigError() {
        var errorBox = document.getElementById("config_generation_error");
        if (!errorBox) return;
        errorBox.textContent = "";
        errorBox.classList.add("hidden");
      }

      function updateConfig() {
        clearConfigError();
        var configBox = document.getElementById("config_json");
        var config = {};
        try {
          config = JSON.parse(configBox.value || "{}");
        } catch (error) {
          console.error("Advanced settings JSON parse failed", error);
          showConfigError("Advanced settings JSON is not valid. Fix it before running optimization.");
          return false;
        }

        try {
          applyPackingMode(config);

          var campaignName = document.getElementById("campaign_name").value.trim();
          var campaignCode = document.getElementById("campaign_code").value.trim();
          var campaignNotes = document.getElementById("campaign_notes").value.trim();
          if (campaignName || campaignCode || campaignNotes) {
            config.campaign = {};
            if (campaignName) config.campaign.name = campaignName;
            if (campaignCode) config.campaign.code = campaignCode;
            if (campaignNotes) config.campaign.notes = campaignNotes;
          }

          var shipBoxType = document.getElementById("ship_as_is_box_type").value.trim();
          lines("ship_as_is_skus").forEach(function (sku) {
            var rule = {
              prepacked: true,
              no_padding: true,
              ships_alone: true,
              can_mix_with_other_items: false,
              box_type: sku + " shipping carton"
            };
            if (shipBoxType) rule.label_box_type = shipBoxType;
            applyRule(config, sku, rule);
          });

          var shipAsIsExceptionSkus = skuValues("ship_as_is_exception_skus");
          if (shipAsIsExceptionSkus.length) {
            config.ship_as_is_exception_skus = shipAsIsExceptionSkus;
          } else {
            delete config.ship_as_is_exception_skus;
          }

          var separatePlaymatSkus = lines("separate_playmat_charge_skus");
          if (separatePlaymatSkus.length) {
            config.separate_playmat_charge_skus = separatePlaymatSkus;
          } else {
            delete config.separate_playmat_charge_skus;
          }
          separatePlaymatSkus.forEach(function (sku) {
            applyRule(config, sku, {
              prepacked: true,
              no_padding: true,
              ships_alone: true,
              can_mix_with_other_items: false,
              separate_playmat_charge: true,
              box_type: sku + " separate playmat parcel"
            });
          });

          lines("no_padding_skus").forEach(function (sku) {
            applyRule(config, sku, { no_padding: true });
          });

          var wrapRule = {
            wrap_around_largest_item: true,
            wrapped_height_cm: numberValue("wrapped_height_cm", 4)
          };
          if (document.getElementById("wrap_no_padding").checked) {
            wrapRule.no_padding = true;
          }
          lines("wrap_skus").forEach(function (sku) {
            applyRule(config, sku, wrapRule);
          });

          var compressibleRule = {
            compressible: true,
            compressed_height_ratio: numberValue("compressed_height_ratio", 0.6),
            compressed_volume_ratio: numberValue("compressed_volume_ratio", 0.75)
          };
          lines("compressible_skus").forEach(function (sku) {
            applyRule(config, sku, compressibleRule);
          });


          configBox.value = JSON.stringify(config, null, 2);
          return true;
        } catch (error) {
          console.error("Advanced settings generation failed", error);
          showConfigError("Advanced settings could not be generated from the form. The upload will continue with the current Advanced settings text.");
          return true;
        }
      }
      document.getElementById("generate_config_button").addEventListener("click", function () {
        updateConfig();
      });
      document.querySelector("form").addEventListener("submit", function (event) {
        if (document.getElementById("campaign_intake")) {
          var canSubmit = updateConfig();
          if (!canSubmit) {
            event.preventDefault();
          }
        }
      });    })();
  </script>
"""
    return (
        html_text
        .replace("__SEL_RAILWAY_FAST__", selected["railway_fast"])
        .replace("__SEL_RAILWAY_BALANCED_30__", selected["railway_balanced_30"])
        .replace("__SEL_RAILWAY_BALANCED_60__", selected["railway_balanced_60"])
        .replace("__SEL_LOCAL_POWER_300__", selected["local_power_balanced_300"])
    )


def _upload_form_html(
    config_text: str,
    upload_token: str | None = None,
    error: str | None = None,
    *,
    title: str = "Box Optimizer Upload",
    action: str = "/upload",
    intro: str = "Upload the two campaign files, then run the optimizer. The file names do not need to follow a special pattern.",
    notice: str = "",
    show_campaign_intake: bool = False,
    default_packing_mode_choice: str = "railway_fast",
    status_code: int = 200,
) -> HTMLResponse:
    safe_config = html.escape(config_text or _default_upload_config_text())
    hidden_token = html.escape(upload_token or "")
    error_html = f'<div class="error"><strong>Something needs attention:</strong> {html.escape(error)}</div>' if error else ""
    notice_html = f'<div class="summary"><strong>{html.escape(notice)}</strong></div>' if notice else ""
    campaign_intake_html = _campaign_intake_html(default_packing_mode_choice) if show_campaign_intake else ""
    action_target = action
    if upload_token:
        separator = "&" if "?" in action_target else "?"
        action_target = f"{action_target}{separator}upload_token={quote(upload_token)}"
    safe_action = html.escape(action_target)
    rate_action = "/rates/upload"
    if upload_token:
        rate_action = f"{rate_action}?upload_token={quote(upload_token)}"
    rate_status_html = _rate_sheet_status_html()
    return _html_page(
        title,
        f"""
<h1>{html.escape(title)}</h1>
<p class="muted">{html.escape(intro)}</p>
{notice_html}
{error_html}
<div id="config_generation_error" class="error hidden" role="alert"></div>
<form action="{safe_action}" method="post" enctype="multipart/form-data">
  <input type="hidden" name="upload_token" value="{hidden_token}">
  <label for="sku_master_file">Put your SKU file here</label>
  <input id="sku_master_file" name="sku_master_file" type="file" accept=".csv,.xlsx" required>

  <label for="orders_file">Put your orders file here</label>
  <input id="orders_file" name="orders_file" type="file" accept=".csv,.xlsx" required>

{campaign_intake_html}

  <fieldset>
    <legend>Add invoice</legend>
    <label><input type="radio" name="invoice_selection" value="none" checked> No</label>
    <label><input type="radio" name="invoice_selection" value="invoice_us"> Yes - US Pay To</label>
    <label><input type="radio" name="invoice_selection" value="invoice_cn"> Yes - CN Pay To</label>
  </fieldset>

  <details>
    <summary>Advanced settings</summary>
    <label for="config_json">Optional packing instructions generated by GPT</label>
    <div class="summary">
      <p><strong>Packing Mode options</strong></p>
      <ul>
        <li><code>"fast"</code> is recommended for Railway uploads and guaranteed completion.</li>
        <li><code>"balanced"</code> gives better quote accuracy, but is slower and best for local runs or smaller files.</li>
      </ul>
      <p><strong>Optional testing limit</strong></p>
      <ul>
        <li>Add <code>"max_orders": 5</code> for a small test run.</li>
        <li>Remove or omit <code>max_orders</code> for the full file.</li>
      </ul>
      <p><strong>Optimization time</strong></p>
      <ul>
        <li>For balanced on Railway, use <code>"max_optimization_seconds": 30</code> or <code>60</code>.</li>
        <li>For local/offline balanced runs, <code>"max_optimization_seconds": 300</code> is useful for quote accuracy.</li>
      </ul>
      <p><strong>Output granularity</strong></p>
      <ul>
        <li><code>"order_summary"</code> is the recommended default.</li>
        <li><code>"box_detail"</code> gives more detail when supported.</li>
      </ul>
    </div>
    <textarea id="config_json" name="config_json">{safe_config}</textarea>
    <details id="rate_sheet_management">
      <summary>Rate Sheet Management</summary>
      <p class="help">Load a new Excel rate sheet without running optimization.</p>
      {rate_status_html}
      <label for="rate_sheet_file">Select new rate sheet</label>
      <input id="rate_sheet_file" name="rate_sheet_file" type="file" accept=".xlsx" form="rate_sheet_upload_form">
      <label for="rate_admin_token">Rate admin token</label>
      <input id="rate_admin_token" name="rate_admin_token" type="password" autocomplete="off" form="rate_sheet_upload_form">
      <button type="submit" form="rate_sheet_upload_form">Load Rates</button>
    </details>
  </details>

  <button type="submit">Run Optimization</button>
</form>
<form id="rate_sheet_upload_form" action="{html.escape(rate_action)}" method="post" enctype="multipart/form-data">
  <input type="hidden" name="upload_token" value="{hidden_token}">
</form>
""",
        status_code=status_code,
    )


def _summary_value(summary: dict, key: str) -> str:
    value = summary.get(key)
    return html.escape(str(value if value is not None else ""))


def _format_elapsed_seconds(value: object) -> str:
    try:
        seconds = float(value)
    except (TypeError, ValueError):
        return ""
    if seconds < 60:
        return f"{seconds:.1f} seconds"
    minutes = int(seconds // 60)
    remaining_seconds = int(round(seconds % 60))
    if remaining_seconds == 60:
        minutes += 1
        remaining_seconds = 0
    return f"{minutes} min {remaining_seconds:02d} sec"


def _result_page(record: dict, upload_token: str | None = None, status_code: int = 200) -> HTMLResponse:
    summary = record.get("summary") or {}
    warnings = summary.get("warnings") or []
    warning_items = "".join(f"<li>{html.escape(str(warning))}</li>" for warning in warnings) or "<li>No workflow warnings.</li>"
    download_url = f"/jobs/{record['job_id']}/download{_token_query(upload_token)}"
    status_url = f"/jobs/{record['job_id']}{_token_query(upload_token)}"
    completed_time = _format_elapsed_seconds(summary.get("elapsed_seconds"))
    completed_time_html = (
        f"  <p><strong>Time to completed:</strong> {html.escape(completed_time)}</p>\n"
        if completed_time
        else ""
    )
    return _html_page(
        "Box Optimizer Results",
        f"""
<h1>Optimization Results</h1>
<div class="summary">
  <p><strong>Job ID:</strong> {html.escape(record['job_id'])}</p>
{completed_time_html}  <p><strong>Status:</strong> {html.escape(str(record.get('status') or ''))}</p>
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


def _error_page(
    message: str,
    upload_token: str | None = None,
    status_code: int = 400,
    return_path: str = "/upload",
    received_post: bool | None = None,
) -> HTMLResponse:
    received_post_html = ""
    if received_post is not None:
        received_post_html = f"<p><strong>FastAPI received POST:</strong> {'yes' if received_post else 'no'}</p>"
    return _html_page(
        "Box Optimizer Error",
        f"""
<h1>Optimization could not finish</h1>
<div class="error">{html.escape(message)}</div>
{received_post_html}
<p><a href="{html.escape(return_path + _token_query(upload_token))}">Return to upload page</a></p>
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
    mode: str | None = None,
):
    """Return the single employee-friendly upload page for Railway and local use."""
    try:
        provided_token = _require_upload_access(upload_token=upload_token, token=token)
    except HTTPException as exc:
        return _error_page(str(exc.detail), status_code=exc.status_code)
    selected_mode = mode if mode in {"railway_fast", "railway_balanced_30", "railway_balanced_60", "local_power_300", "local_power_balanced_300"} else "railway_fast"
    config_text = job_config or (_power_upload_config_text() if selected_mode in {"local_power_300", "local_power_balanced_300"} else _default_upload_config_text())
    return _upload_form_html(
        config_text=config_text,
        upload_token=provided_token,
        show_campaign_intake=True,
        default_packing_mode_choice=selected_mode,
    )


def _run_upload_job(
    *,
    sku_master_file: UploadFile,
    orders_file: UploadFile,
    config_json: str | None,
    upload_token: str | None,
    default_config: dict[str, Any],
    final_config: dict[str, Any] | None = None,
    return_path: str,
) -> HTMLResponse:
    started = time.perf_counter()
    job_id = uuid4().hex
    created_seconds = time.time()
    job_dir = _jobs_root() / job_id
    try:
        provided_token = _require_upload_access(upload_token=upload_token)
    except HTTPException as exc:
        return _error_page(str(exc.detail), upload_token=upload_token, status_code=exc.status_code, return_path=return_path)

    try:
        _cleanup_expired_jobs()
        job_dir.mkdir(parents=True, exist_ok=False)
        if final_config is None:
            submitted_config = _parse_config(config_json) if config_json is not None else {}
            config = {**default_config, **submitted_config}
        else:
            config = final_config
        config = {**config, "workbook_output_mode": _workbook_output_mode_for_token(provided_token)}
        if config.get("debug"):
            logger.setLevel(logging.DEBUG)
        output_filename = _campaign_download_filename(config)

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
            "filename": output_filename,
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
            "summary": {},
            "download_url": None,
            "error": str(exc.detail),
        }
        try:
            job_dir.mkdir(parents=True, exist_ok=True)
            _write_job_record(job_dir, record)
        except OSError:
            pass
        return _error_page(str(exc.detail), upload_token=provided_token, status_code=exc.status_code, return_path=return_path, received_post=True)
    except Exception as exc:
        logger.exception("upload optimization failed")
        message = str(exc) or "Optimization failed."
        record = {
            "job_id": job_id,
            "status": "failed",
            "created_at": datetime.fromtimestamp(created_seconds, timezone.utc).isoformat(),
            "expires_at": _job_expiration(created_seconds),
            "summary": {},
            "download_url": None,
            "error": message,
        }
        try:
            job_dir.mkdir(parents=True, exist_ok=True)
            _write_job_record(job_dir, record)
        except OSError:
            pass
        return _error_page(message, upload_token=provided_token, status_code=500, return_path=return_path, received_post=True)


@app.post("/upload", response_class=HTMLResponse, include_in_schema=False)
def upload_workbooks(
    sku_master_file: UploadFile | None = File(default=None),
    orders_file: UploadFile | None = File(default=None),
    config_json: str | None = Form(default=None),
    upload_token: str | None = Form(default=None),
    campaign_name: str | None = Form(default=None),
    campaign_code: str | None = Form(default=None),
    campaign_notes: str | None = Form(default=None),
    packing_mode_choice: str | None = Form(default="railway_fast"),
    ship_as_is_skus: str | None = Form(default=None),
    ship_as_is_exception_skus: str | None = Form(default=None),
    ship_as_is_box_type: str | None = Form(default=None),
    separate_playmat_charge_skus: str | None = Form(default=None),
    no_padding_skus: str | None = Form(default=None),
    wrap_around_skus: str | None = Form(default=None),
    wrapped_height_cm: str | None = Form(default=None),
    compressible_skus: str | None = Form(default=None),
    compressed_height_ratio: str | None = Form(default=None),
    compressed_volume_ratio: str | None = Form(default=None),
    invoice_selection: str | None = Form(default=None),
):
    """Run the optimizer from the single upload page and build structured rules server-side."""
    if sku_master_file is None or orders_file is None:
        return _error_page(
            "FastAPI received the upload request, but both SKU and orders files are required.",
            upload_token=upload_token,
            status_code=400,
            return_path="/upload",
            received_post=True,
        )
    structured_config = _structured_upload_config(
        campaign_name=campaign_name,
        campaign_code=campaign_code,
        campaign_notes=campaign_notes,
        packing_mode_choice=packing_mode_choice,
        ship_as_is_skus=ship_as_is_skus,
        ship_as_is_exception_skus=ship_as_is_exception_skus,
        ship_as_is_box_type=ship_as_is_box_type,
        separate_playmat_charge_skus=separate_playmat_charge_skus,
        no_padding_skus=no_padding_skus,
        wrap_around_skus=wrap_around_skus,
        wrapped_height_cm=wrapped_height_cm,
        compressible_skus=compressible_skus,
        compressed_height_ratio=compressed_height_ratio,
        compressed_volume_ratio=compressed_volume_ratio,
        invoice_selection=invoice_selection,
    )
    try:
        final_config = _final_upload_config(config_json, structured_config)
        if invoice_selection is not None:
            final_config = _parse_config(_merge_config(final_config, _invoice_config_for_selection(invoice_selection)))
    except HTTPException as exc:
        return _error_page(
            str(exc.detail),
            upload_token=upload_token,
            status_code=exc.status_code,
            return_path="/upload",
            received_post=True,
        )
    return _run_upload_job(
        sku_master_file=sku_master_file,
        orders_file=orders_file,
        config_json=config_json,
        upload_token=upload_token,
        default_config=DEFAULT_UPLOAD_CONFIG,
        final_config=final_config,
        return_path="/upload",
    )


@app.post("/rates/upload", response_class=HTMLResponse, include_in_schema=False)
def upload_rate_sheet(
    rate_sheet_file: UploadFile = File(...),
    upload_token: str | None = Form(default=None),
    rate_admin_token: str | None = Form(default=None),
):
    """Load a new active rate sheet without running optimization."""
    page_upload_token = ""
    try:
        page_upload_token = _require_upload_access(upload_token=upload_token)
    except HTTPException:
        page_upload_token = ""
    try:
        _require_rate_upload_access(upload_token=upload_token, rate_admin_token=rate_admin_token)
    except HTTPException as exc:
        return _error_page(str(exc.detail), upload_token=upload_token, status_code=exc.status_code, return_path="/upload")

    if Path(rate_sheet_file.filename or "").suffix.lower() != ".xlsx":
        return _upload_form_html(
            _default_upload_config_text(),
            upload_token=page_upload_token,
            error="Rate sheet upload accepts .xlsx files only.",
            show_campaign_intake=True,
            status_code=400,
        )

    root = _rate_sheet_root()
    temp_path = root / f"candidate_{uuid4().hex}.xlsx"
    try:
        _save_upload(rate_sheet_file, root, temp_path.name)
        validation = _validate_rate_sheet(temp_path)
        active_path = _active_rate_sheet_path()
        temp_path.replace(active_path)
        size_bytes = active_path.stat().st_size
        sha256 = _sha256_file(active_path)
        metadata = {
            "original_filename": Path(rate_sheet_file.filename or ACTIVE_RATE_SHEET_FILENAME).name,
            "saved_filename": ACTIVE_RATE_SHEET_FILENAME,
            "storage_path": str(active_path),
            "uploaded_at": datetime.now(timezone.utc).isoformat(),
            "size_bytes": size_bytes,
            "sha256": sha256,
            "validation": validation,
            "source": "active_upload",
            "app_version": app.version,
        }
        _rate_sheet_metadata_path().write_text(json.dumps(metadata, indent=2), encoding="utf-8")
    except HTTPException as exc:
        try:
            temp_path.unlink()
        except OSError:
            pass
        return _upload_form_html(
            _default_upload_config_text(),
            upload_token=page_upload_token,
            error=str(exc.detail),
            show_campaign_intake=True,
            status_code=exc.status_code,
        )
    except Exception:
        logger.exception("rate sheet upload failed")
        try:
            temp_path.unlink()
        except OSError:
            pass
        return _upload_form_html(
            _default_upload_config_text(),
            upload_token=page_upload_token,
            error="Rate sheet upload failed before it could be saved.",
            show_campaign_intake=True,
            status_code=500,
        )

    return _upload_form_html(
        _default_upload_config_text(),
        upload_token=page_upload_token,
        notice="Rate sheet loaded successfully.",
        show_campaign_intake=True,
    )


@app.get("/rates/current", response_class=JSONResponse, include_in_schema=False)
def current_rate_sheet(
    upload_token: str | None = None,
    token: str | None = None,
):
    """Return metadata for the current active rate sheet."""
    _require_upload_access(upload_token=upload_token, token=token)
    return JSONResponse(_rate_sheet_status())


@app.get("/rates/download", include_in_schema=False)
def download_rate_sheet(
    upload_token: str | None = None,
    token: str | None = None,
):
    """Download the current active rate sheet."""
    _require_rate_download_access(upload_token=upload_token, token=token)
    active_path = _active_rate_sheet_path()
    if not active_path.exists():
        raise HTTPException(status_code=404, detail="No active rate sheet loaded")
    metadata = _rate_sheet_metadata()
    filename = str(metadata.get("original_filename") or ACTIVE_RATE_SHEET_FILENAME)
    return FileResponse(active_path, media_type=WORKBOOK_CONTENT_TYPE, filename=filename)


@app.get("/power-upload", include_in_schema=False)
def power_upload_page(
    upload_token: str | None = None,
    token: str | None = None,
    job_config: str | None = None,
):
    """Compatibility shortcut to the single upload page with local power defaults."""
    params = {"mode": "local_power_balanced_300"}
    provided_token = upload_token or token
    if provided_token:
        params["upload_token"] = provided_token
    if job_config:
        params["job_config"] = job_config
    return RedirectResponse(url=f"/upload?{urlencode(params)}", status_code=303)


@app.post("/power-upload", include_in_schema=False)
def power_upload_workbooks():
    """Deprecated compatibility route; use /upload for all local runs."""
    return RedirectResponse(url="/upload", status_code=307)

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
        filename=str(record.get("filename") or "optimized_shipping_plan.xlsx"),
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
    try:
        _log_event("request_received", endpoint="/optimize_base64", run_id=run_id)
        stage = "config"
        config = _parse_config(payload.config_json)
        if config.get("debug"):
            logger.setLevel(logging.DEBUG)
        output_filename = _campaign_download_filename(config)

        work_dir = _work_dir(run_id)
        _log_event(
            "filenames_received",
            endpoint="/optimize_base64",
            sku_master_file=payload.sku_master_filename,
            orders_file=payload.orders_filename,
        )
        stage = "files_saved"
        sku_master_path, orders_path = _save_base64_request_files(payload, work_dir)
        output_path = work_dir / "optimized_shipping_plan.xlsx"
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
        output_filename = _campaign_download_filename(config)

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
            filename=output_filename,
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
