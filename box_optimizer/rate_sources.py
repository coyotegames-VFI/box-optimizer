"""Managed customer rate sheet storage and validation helpers."""

from __future__ import annotations

import hashlib
import json
import os
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode, urljoin
from urllib.request import Request, urlopen
from xml.etree import ElementTree


ACTIVE_RATE_SHEET_FILENAME = "current_rate_sheet.xlsx"
RATE_SHEET_METADATA_FILENAME = "current_rate_sheet.json"
RATE_SYNC_URL_ENV = "BOX_OPTIMIZER_RATE_SYNC_URL"
RATE_SYNC_TOKEN_ENV = "BOX_OPTIMIZER_RATE_SYNC_TOKEN"
RATE_ADMIN_TOKEN_ENV = "BOX_OPTIMIZER_RATE_ADMIN_TOKEN"


class RateSheetValidationError(ValueError):
    """Raised when an uploaded or active rate sheet fails minimum validation."""


def rate_sheet_root() -> Path:
    configured_root = os.getenv("BOX_OPTIMIZER_RATE_SHEET_DIR")
    railway_mount = os.getenv("RAILWAY_VOLUME_MOUNT_PATH")
    if configured_root:
        root = Path(configured_root)
    elif railway_mount:
        root = Path(railway_mount) / "rates"
    else:
        root = Path.cwd() / "runtime" / "rates"
    root.mkdir(parents=True, exist_ok=True)
    return root


def active_rate_sheet_path() -> Path:
    return rate_sheet_root() / ACTIVE_RATE_SHEET_FILENAME


def rate_sheet_metadata_path() -> Path:
    return rate_sheet_root() / RATE_SHEET_METADATA_FILENAME


def rate_sheet_metadata() -> dict[str, Any]:
    path = rate_sheet_metadata_path()
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def validate_rate_sheet(path: Path) -> dict[str, Any]:
    try:
        with zipfile.ZipFile(path) as archive:
            workbook_xml = archive.read("xl/workbook.xml")
    except (KeyError, zipfile.BadZipFile, OSError) as exc:
        raise RateSheetValidationError("Rate sheet must be a readable .xlsx workbook.") from exc

    try:
        root = ElementTree.fromstring(workbook_xml)
    except ElementTree.ParseError as exc:
        raise RateSheetValidationError("Rate sheet workbook metadata could not be read.") from exc
    ns = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    sheet_names = [
        str(sheet.attrib.get("name") or "")
        for sheet in root.findall("main:sheets/main:sheet", ns)
    ]
    if "Zone Key" not in sheet_names or len(sheet_names) < 2:
        raise RateSheetValidationError("Rate sheet must include a Zone Key sheet and a country/zone mapping sheet.")
    return {
        "zone_key": True,
        "mapping_sheet_count": max(len(sheet_names) - 1, 0),
        "sheet_names": sheet_names,
    }


def rate_sheet_status() -> dict[str, Any]:
    active_path = active_rate_sheet_path()
    metadata = rate_sheet_metadata()
    if not active_path.exists():
        return {
            "active": False,
            "message": "No active rate sheet loaded",
            "filename": "",
            "uploaded_at": "",
            "size_bytes": 0,
            "sha256": "",
            "checksum": "",
            "storage_path": "",
            "validation": {},
            "source": "missing",
        }
    uploaded_at = str(metadata.get("uploaded_at") or metadata.get("loaded_at") or "")
    sha256 = str(metadata.get("sha256") or sha256_file(active_path))
    size_bytes = int(metadata.get("size_bytes") or active_path.stat().st_size)
    return {
        "active": True,
        "message": "Current active rate sheet loaded",
        "filename": str(metadata.get("original_filename") or active_path.name),
        "saved_filename": str(metadata.get("saved_filename") or active_path.name),
        "uploaded_at": uploaded_at,
        "size_bytes": size_bytes,
        "sha256": sha256,
        "checksum": sha256,
        "checksum_short": sha256[:12],
        "storage_path": str(active_path),
        "validation": metadata.get("validation") or {},
        "source": str(metadata.get("source") or "active_upload"),
    }


def _sync_request_url(base_url: str, endpoint: str) -> str:
    url = urljoin(f"{base_url.rstrip('/')}/", endpoint.lstrip("/"))
    token = (
        os.getenv(RATE_SYNC_TOKEN_ENV)
        or os.getenv(RATE_ADMIN_TOKEN_ENV)
        or (
            os.getenv("BOX_OPTIMIZER_UPLOAD_TOKEN")
            if not os.getenv(RATE_SYNC_TOKEN_ENV) and not os.getenv(RATE_ADMIN_TOKEN_ENV)
            else ""
        )
    )
    if token:
        separator = "&" if "?" in url else "?"
        url = f"{url}{separator}{urlencode({'upload_token': token})}"
    return url


def _remote_json(url: str, timeout_seconds: float) -> dict[str, Any]:
    request = Request(url, headers={"Accept": "application/json"})
    with urlopen(request, timeout=timeout_seconds) as response:
        return json.loads(response.read().decode("utf-8"))


def _remote_bytes(url: str, timeout_seconds: float) -> bytes:
    request = Request(url, headers={"Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"})
    with urlopen(request, timeout=timeout_seconds) as response:
        return response.read()


def sync_rate_sheet_from_remote(timeout_seconds: float = 10.0) -> dict[str, Any]:
    """Sync local active rate sheet cache from a configured remote source."""
    base_url = str(os.getenv(RATE_SYNC_URL_ENV) or "").strip()
    if not base_url:
        return {"attempted": False, "status": "skipped", "message": ""}

    current_url = _sync_request_url(base_url, "/rates/current")
    download_url = _sync_request_url(base_url, "/rates/download")
    try:
        remote_status = _remote_json(current_url, timeout_seconds)
    except (OSError, HTTPError, URLError, TimeoutError, json.JSONDecodeError) as exc:
        return {
            "attempted": True,
            "status": "failed",
            "message": "Rate Sheet Sync: failed, using local cached/default rate sheet.",
            "detail": str(exc),
        }

    if not remote_status.get("active"):
        return {
            "attempted": True,
            "status": "no_active",
            "message": "Rate Sheet Sync: Railway has no active rate sheet; using local cached/default rate sheet.",
        }

    remote_sha = str(remote_status.get("sha256") or remote_status.get("checksum") or "").strip()
    local_status = rate_sheet_status()
    if remote_sha and local_status.get("active") and str(local_status.get("sha256") or "") == remote_sha:
        return {
            "attempted": True,
            "status": "up_to_date",
            "message": "Rate Sheet Sync: up to date from Railway.",
            "filename": remote_status.get("filename", ""),
            "sha256": remote_sha,
        }

    active_path = active_rate_sheet_path()
    temp_path = active_path.with_name(f"candidate_sync_{os.getpid()}_{ACTIVE_RATE_SHEET_FILENAME}")
    try:
        workbook_bytes = _remote_bytes(download_url, timeout_seconds)
        temp_path.write_bytes(workbook_bytes)
        validation = validate_rate_sheet(temp_path)
        sha256 = sha256_file(temp_path)
        if remote_sha and sha256 != remote_sha:
            raise RateSheetValidationError("Downloaded rate sheet checksum did not match Railway metadata.")
        temp_path.replace(active_path)
        metadata = {
            "original_filename": str(remote_status.get("filename") or ACTIVE_RATE_SHEET_FILENAME),
            "saved_filename": ACTIVE_RATE_SHEET_FILENAME,
            "storage_path": str(active_path),
            "uploaded_at": str(remote_status.get("uploaded_at") or ""),
            "synced_at": datetime.now(timezone.utc).isoformat(),
            "size_bytes": active_path.stat().st_size,
            "sha256": sha256,
            "validation": validation,
            "source": "remote_sync",
            "remote_source": str(remote_status.get("source") or ""),
            "remote_url": base_url.rstrip("/"),
        }
        rate_sheet_metadata_path().write_text(json.dumps(metadata, indent=2), encoding="utf-8")
    except (OSError, HTTPError, URLError, TimeoutError, RateSheetValidationError) as exc:
        try:
            temp_path.unlink()
        except OSError:
            pass
        return {
            "attempted": True,
            "status": "failed",
            "message": "Rate Sheet Sync: failed, using local cached/default rate sheet.",
            "detail": str(exc),
        }

    return {
        "attempted": True,
        "status": "downloaded",
        "message": "Rate Sheet Sync: downloaded current Railway rate sheet.",
        "filename": metadata["original_filename"],
        "sha256": sha256,
    }
