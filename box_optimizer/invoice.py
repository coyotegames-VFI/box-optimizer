"""Generated invoice sheet support."""

from __future__ import annotations

import json
import os
import re
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Mapping


INVOICE_WARNING_TEXT = (
    "Warning: Some packages are missing scans or final costs. Review before sending invoice."
)

PAY_TO_INCOMPLETE_WARNING_TEXT = "Warning: Pay To config is incomplete."

US_PAY_TO_FIELDS = [
    ("Account holder", "account_holder", "INVOICE_PAY_TO_US_ACCOUNT_HOLDER"),
    ("Bank name", "bank_name", "INVOICE_PAY_TO_US_BANK_NAME"),
    ("Account number", "account_number", "INVOICE_PAY_TO_US_ACCOUNT_NUMBER"),
    ("Routing number", "routing_number", "INVOICE_PAY_TO_US_ROUTING_NUMBER"),
    ("SWIFT code", "swift_code", "INVOICE_PAY_TO_US_SWIFT_CODE"),
    ("IBAN", "iban", "INVOICE_PAY_TO_US_IBAN"),
    ("Address", "address", "INVOICE_PAY_TO_US_ADDRESS"),
    ("Email", "email", "INVOICE_PAY_TO_US_EMAIL"),
    ("Extra", "extra_lines", "INVOICE_PAY_TO_US_EXTRA_LINES"),
]

CN_PAY_TO_FIELDS = [
    ("Account holder", "account_holder", "INVOICE_PAY_TO_CN_ACCOUNT_HOLDER"),
    ("Bank name", "bank_name", "INVOICE_PAY_TO_CN_BANK_NAME"),
    ("Account number", "account_number", "INVOICE_PAY_TO_CN_ACCOUNT_NUMBER"),
    ("SWIFT code", "swift_code", "INVOICE_PAY_TO_CN_SWIFT_CODE"),
    ("CNAPS", "cnaps", "INVOICE_PAY_TO_CN_CNAPS"),
    ("Beneficiary address", "beneficiary_address", "INVOICE_PAY_TO_CN_BENEFICIARY_ADDRESS"),
    ("Postal code", "postal_code", "INVOICE_PAY_TO_CN_POSTAL_CODE"),
    ("Phone", "phone", "INVOICE_PAY_TO_CN_PHONE"),
    ("Extra", "extra_lines", "INVOICE_PAY_TO_CN_EXTRA_LINES"),
]

PAY_TO_DISPLAY_LINES_ENV = {
    "US": "INVOICE_PAY_TO_US_DISPLAY_LINES",
    "CN": "INVOICE_PAY_TO_CN_DISPLAY_LINES",
}


@dataclass(frozen=True)
class InvoicePayload:
    variant: str
    invoice_number: str
    invoice_date: date = field(default_factory=date.today)
    campaign_name: str = ""
    bill_to: str = ""
    email: str = ""
    address_lines: tuple[str, ...] = ()
    inbound_fee: object = ""
    ship_order_count: int = 0
    cost_summary_sheet_name: str = "Cost Summary"
    final_cost_column: str = ""
    final_weight_column: str = ""
    scan_note_column: str = ""
    cost_summary_data_start_row: int = 2
    cost_summary_data_end_row: int = 1
    pay_to_lines: tuple[tuple[str, str], ...] = ()
    pay_to_display_lines: bool = False
    pay_to_incomplete: bool = False


def invoice_enabled(config: Mapping | None) -> bool:
    value = (config or {}).get("include_invoice")
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().lower() in {"true", "1", "yes", "y", "on"}
    return bool(value)


def normalize_invoice_variant(value: object) -> str | None:
    variant = str(value or "US").strip().upper()
    return variant if variant in {"US", "CN"} else None


def invoice_number(short_campaign_slug: object, generation_date: date | None = None) -> str:
    slug = str(short_campaign_slug or "VFI").strip() or "VFI"
    generation_date = generation_date or date.today()
    return f"{slug}-{generation_date:%m%d%Y}"


def first_present(mapping: Mapping | None, keys: list[str]) -> str:
    mapping = mapping or {}
    for key in keys:
        value = str(mapping.get(key, "") or "").strip()
        if value:
            return value
    return ""


def parse_money(value: object) -> float | str:
    text = str(value or "").strip()
    if not text:
        return ""
    match = re.search(r"-?(?:\d+(?:\.\d*)?|\.\d+)", text.replace(",", ""))
    if not match:
        return ""
    number = float(match.group(0))
    truncated = int(number * 100) / 100
    return round(truncated, 2)


def _load_local_invoice_config(path: Path | None = None) -> dict:
    config_path = path or Path("local_reference") / "invoice_config.json"
    if not config_path.exists():
        return {}
    try:
        data = json.loads(config_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return data if isinstance(data, dict) else {}


def _variant_config(source: Mapping | None, variant: str) -> Mapping:
    if not isinstance(source, Mapping):
        return {}
    direct = source.get(variant) or source.get(variant.lower())
    if isinstance(direct, Mapping):
        return direct
    pay_to = source.get("pay_to") or source.get("invoice_pay_to")
    if isinstance(pay_to, Mapping):
        nested = pay_to.get(variant) or pay_to.get(variant.lower())
        if isinstance(nested, Mapping):
            return nested
    return source


def _string_list(value: object) -> list[str]:
    if not isinstance(value, list):
        return []
    return [str(line) for line in value]


def _strict_string_list(value: object) -> list[str] | None:
    if not isinstance(value, list):
        return None
    if not all(isinstance(line, str) for line in value):
        return None
    return list(value)


def _env_display_lines(variant: str, environ: Mapping[str, str]) -> list[str] | None:
    env_name = PAY_TO_DISPLAY_LINES_ENV.get(variant)
    if not env_name or env_name not in environ:
        return None
    try:
        parsed = json.loads(environ.get(env_name) or "")
    except json.JSONDecodeError:
        return None
    return _strict_string_list(parsed)


def _display_lines_from_blocks(value: object) -> list[str]:
    if not isinstance(value, list):
        return []
    lines: list[str] = []
    for block in value:
        if not isinstance(block, Mapping):
            continue
        if lines:
            lines.append("")
        title = str(block.get("title", "") or "")
        if title:
            lines.append(title)
        for line in _string_list(block.get("lines")):
            lines.append(line)
    return lines


def _configured_display_lines(*sources: Mapping) -> list[str]:
    for source in sources:
        display_lines = _string_list(source.get("display_lines"))
        if display_lines:
            return display_lines
        block_lines = _display_lines_from_blocks(source.get("payment_blocks"))
        if block_lines:
            return block_lines
    return []


def pay_to_lines(
    variant: str,
    *,
    config: Mapping | None = None,
    environ: Mapping[str, str] | None = None,
    local_config_path: Path | None = None,
) -> tuple[tuple[tuple[str, str], ...], bool]:
    fields = US_PAY_TO_FIELDS if variant == "US" else CN_PAY_TO_FIELDS
    environ = environ or os.environ
    local_config = _variant_config(_load_local_invoice_config(local_config_path), variant)
    runtime_config = _variant_config((config or {}).get("invoice_pay_to"), variant)
    env_display_lines = _env_display_lines(variant, environ)
    if env_display_lines is not None:
        return tuple((line, "") for line in env_display_lines), not any(
            line.strip() for line in env_display_lines
        )

    display_lines = _configured_display_lines(runtime_config, local_config)
    if any(line.strip() for line in display_lines):
        return tuple((line, "") for line in display_lines), False

    lines: list[tuple[str, str]] = []
    missing_required = False
    for label, key, env_name in fields:
        value = str(local_config.get(key, "") or "").strip()
        value = str(runtime_config.get(key, value) or "").strip()
        value = str(environ.get(env_name, value) or "").strip()
        if key == "extra_lines":
            for line in re.split(r"\r?\n|\|", value):
                if line.strip():
                    lines.append(("", line.strip()))
            continue
        if not value:
            missing_required = True
        lines.append((label, value))
    return tuple(lines), missing_required


def pay_to_uses_display_lines(
    variant: str,
    *,
    config: Mapping | None = None,
    environ: Mapping[str, str] | None = None,
    local_config_path: Path | None = None,
) -> bool:
    environ = environ or os.environ
    env_display_lines = _env_display_lines(variant, environ)
    if env_display_lines is not None:
        return True

    local_config = _variant_config(_load_local_invoice_config(local_config_path), variant)
    runtime_config = _variant_config((config or {}).get("invoice_pay_to"), variant)
    display_lines = _configured_display_lines(runtime_config, local_config)
    return any(line.strip() for line in display_lines)


def cost_summary_invoice_columns(headers: list[str]) -> dict[str, str]:
    from box_optimizer.io.excel_writer import _column_letter

    return {
        key: _column_letter(headers.index(key))
        for key in ["Final cost", "Final weight kg", "Scan note"]
        if key in headers
    }
