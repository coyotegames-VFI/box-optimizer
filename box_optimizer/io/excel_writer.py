"""Excel output helpers."""

from pathlib import Path
from dataclasses import dataclass
import re
from xml.sax.saxutils import escape
import zipfile

from box_optimizer.io.qr import qr_png
from box_optimizer.invoice import InvoicePayload, normalize_invoice_variant


REQUIRED_SHEETS = [
    "Summary",
    "Cost Summary",
    "Stock Count",
    "Actual Dimensions",
    "Labels",
    "VFI Intake Form",
    "Optimized to Pack",
    "Label generator",
    "Order Volume Weights",
    "Box Size Summary",
]

OPTIONAL_SHEETS = [
    "_ActualLookupTable",
    "_ActualRateTable",
    "Unmatched SKUs",
    "Multi Box Detail",
    "Pledge Combination Summary",
    "Debug Summary",
    "Packing Detail",
    "Input Column Mapping",
    "Errors and Warnings",
]

FAST_PRODUCTION_SKIPPED_SHEETS = {
    "Label generator",
    "Order Volume Weights",
    "Packing Detail",
    "Multi Box Detail",
    "Pledge Combination Summary",
    "Debug Summary",
    "Input Column Mapping",
}

MAX_LABELS_PER_SHEET = 1000
MAX_MANUAL_ROW_BREAKS_PER_SHEET = 1026

ORDER_VOLUME_WEIGHTS_COLUMNS = [
    "Region",
    "Order ID",
    "VFI #",
    "Country",
    "State/Province",
    "US State Abbreviation",
    "Packed Actual Weight kg",
    "Dimensional Weight kg (/5000)",
    "Chargeable Weight kg",
    "Chargeable Weight g",
    "Customer Cost",
    "Estimated VFI Cost",
    "Margin",
    "Total Units",
    "Box Qty",
    "Box Plan",
    "Per-Box Chargeable Weight",
    "SKU Breakdown",
]

ACTUAL_DIMENSIONS_COLUMNS = [
    "Scan barcode",
    "Weight in grams",
    "Length",
    "Width",
    "Height",
    "Actual DIM weight kg",
    "Estimated weight in grams",
    "Weight warning",
    "Actual total shipping cost",
    "Quoted shipping cost",
    "Actual vs quoted difference",
    "Expected scan barcode",
    "Scan status",
    "Helper/debug separator",
    "Cost Summary VFI #",
    "Group VFI key",
    "Is charge row",
    "Country",
    "Pick count / Total units",
    "Add-on adjusters",
    "Actual weight kg",
    "Actual DIM weight kg helper",
    "Actual chargeable weight kg",
    "Carton chargeable weight kg",
    "Order/group chargeable weight kg",
    "Hub zone / rate zone",
    "Matched rate weight band",
    "Actual hub shipping fee",
    "Pick / add-on fee",
    "Lookup status",
    "Expected scan group VFI key",
]

ACTUAL_DIMENSIONS_MAIN_COLUMNS = set(ACTUAL_DIMENSIONS_COLUMNS[:14])
ACTUAL_DIMENSIONS_FIRST_HELPER_COLUMN = "Cost Summary VFI #"
ACTUAL_DIMENSIONS_HELPER_COLUMNS = set(ACTUAL_DIMENSIONS_COLUMNS[14:])

ACTUAL_DIMENSIONS_COLUMN_WIDTHS = {
    "Scan barcode": 24,
    "Weight in grams": 13,
    "Length": 12,
    "Width": 12,
    "Height": 12,
    "Actual DIM weight kg": 16,
    "Estimated weight in grams": 20,
    "Weight warning": 20,
    "Actual total shipping cost": 18,
    "Quoted shipping cost": 17,
    "Actual vs quoted difference": 18,
    "Expected scan barcode": 26,
    "Scan status": 18,
    "Helper/debug separator": 4,
    "Cost Summary VFI #": 18,
    "Group VFI key": 18,
    "Is charge row": 14,
    "Country": 16,
    "Pick count / Total units": 18,
    "Add-on adjusters": 16,
    "Actual weight kg": 16,
    "Actual DIM weight kg helper": 16,
    "Actual chargeable weight kg": 18,
    "Carton chargeable weight kg": 18,
    "Order/group chargeable weight kg": 18,
    "Hub zone / rate zone": 18,
    "Matched rate weight band": 18,
    "Actual hub shipping fee": 16,
    "Pick / add-on fee": 16,
    "Lookup status": 20,
    "Expected scan group VFI key": 26,
}


@dataclass(frozen=True)
class ExcelFormula:
    formula: str

_UNIT_HINTS = {
    "length": "cm",
    "width": "cm",
    "height": "cm",
    "weight": "kg",
    "volume": "cm3",
}

LABEL_ADDRESS_SPLIT_MIN = 35
LABEL_ADDRESS_SPLIT_MAX = 50


def _column_letter(index: int) -> str:
    result = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def _safe_sheet_name(name: str) -> str:
    invalid = set("[]:*?/\\")
    cleaned = "".join("_" if char in invalid else char for char in name).strip()
    return cleaned[:31] or "Sheet"


def _header_with_units(header: str) -> str:
    normalized = header.lower()
    if header in ORDER_VOLUME_WEIGHTS_COLUMNS:
        return header
    if any(unit in normalized for unit in [" cm", " kg", " lb"]) or normalized.endswith(" g"):
        return header
    if "(" in header and ")" in header:
        return header
    for token, unit in _UNIT_HINTS.items():
        if token in normalized:
            return f"{header} ({unit})"
    return header


def _display_header_for_sheet(sheet_name: str, header: str) -> str:
    if sheet_name == "Actual Dimensions":
        if header == "Helper/debug separator":
            return ""
        return header
    if sheet_name.startswith("Cost Summary") and header in {"Hub Shipping Fee", "Express"}:
        return f"{header} (USD)"
    if header == "Country Detail":
        return "Country"
    if header in {"_Country Scan Future Cost Blank"}:
        return ""
    if header.startswith("_Country Scan Metadata Blank "):
        return ""
    return _header_with_units(header)


def _string_cell(reference: str, value: object, style: int | None = None) -> str:
    style_attr = f' s="{style}"' if style is not None else ""
    text = escape("" if value is None else str(value))
    return (
        f'<c r="{reference}" t="inlineStr"{style_attr}>'
        f"<is><t>{text}</t></is>"
        "</c>"
    )


def _number_cell(reference: str, value: int | float, style: int | None = None) -> str:
    style_attr = f' s="{style}"' if style is not None else ""
    return f'<c r="{reference}"{style_attr}><v>{value}</v></c>'


def _formula_cell(reference: str, formula: str, style: int | None = None) -> str:
    style_attr = f' s="{style}"' if style is not None else ""
    return f'<c r="{reference}"{style_attr}><f>{escape(formula)}</f></c>'


def _cell_xml(row_index: int, column_index: int, value: object, style: int | None = None) -> str:
    reference = f"{_column_letter(column_index)}{row_index}"
    if isinstance(value, ExcelFormula):
        return _formula_cell(reference, value.formula, style)
    if isinstance(value, bool):
        return f'<c r="{reference}" t="b"><v>{1 if value else 0}</v></c>'
    if isinstance(value, int | float) and not isinstance(value, bool):
        return _number_cell(reference, value, style)
    return _string_cell(reference, value, style)


def _cell_style_for_sheet_value(
    sheet_name: str,
    row_index: int,
    column_index: int,
    headers: list[str],
    row: list[object],
) -> int | None:
    header = headers[column_index] if column_index < len(headers) else ""
    if row_index == 1 and sheet_name == "Actual Dimensions":
        if header == ACTUAL_DIMENSIONS_FIRST_HELPER_COLUMN:
            return 29
        if header in ACTUAL_DIMENSIONS_HELPER_COLUMNS:
            return 28
        return 27
    if row_index == 1:
        return 1
    if sheet_name == "Summary" and header == "Value":
        metric = str(row[1] if len(row) > 1 else "")
        if metric == "Total Chargeable Cost":
            return 9
    if sheet_name == "Stock Count" and header == "Remaining":
        try:
            if float(row[column_index] or 0) < 0:
                return 20
        except (TypeError, ValueError):
            pass
    if sheet_name.startswith("Cost Summary") and header in {"Hub Shipping Fee (USD)", "Express (USD)"}:
        return 19
    if sheet_name.startswith("Cost Summary") and header == "Final cost":
        return 19
    if sheet_name.startswith("Cost Summary") and header == "Final weight kg":
        return 31
    if sheet_name == "Actual Dimensions" and header in {
        "Actual hub shipping fee",
        "Pick / add-on fee",
        "Actual total shipping cost",
        "Quoted shipping cost",
        "Actual vs quoted difference",
    }:
        return 19
    if sheet_name == "Actual Dimensions" and header == "Scan status":
        return 26
    if sheet_name == "Actual Dimensions" and header == ACTUAL_DIMENSIONS_FIRST_HELPER_COLUMN:
        return 30
    return None


def _label_cell_style(row_offset: int, column_index: int, value: object) -> int | None:
    text = str(value or "")
    if row_offset == 0:
        return 7
    if (
        text.startswith("Detailed description of contents:")
        or text.startswith("On Arrival:")
        or text.startswith("If product damage is found")
    ):
        return 2
    if text == "Qty":
        return 5
    if text == "SKU":
        return 6
    if text and ((row_offset == 6 and column_index == 5) or (row_offset == 11 and column_index == 3)):
        return 8
    if row_offset >= 13 and column_index in {0, 3} and text:
        return 3
    if row_offset >= 13 and column_index in {1, 4} and text:
        return 4
    return None


def _split_label_from_line(value: object) -> tuple[str, str]:
    text = str(value or "").strip()
    if text == "No.23 Baosheng Rd.,Bld2, 3rd Floor, Longhai, Fujian, China 363107":
        return "No.23 Baosheng Rd.,Bld2, 3rd Floor,", "Longhai, Fujian, China 363107"
    marker = "Longhai,"
    if marker in text:
        first, second = text.split(marker, 1)
        return first.strip(" ,"), f"{marker} {second.strip(' ,')}".strip()
    return text, ""


def _label_to_name_with_backer(label: dict) -> str:
    name = str(label.get("To Name", "") or "").strip()
    backer_id = str(label.get("Backer ID", "") or "").strip()
    if name and backer_id:
        return f"{name}  Backer ID {backer_id}"
    return name


def _label_campaign_pledge_footer(label: dict) -> str:
    campaign = str(label.get("Campaign Name", "") or "").strip()
    pledge = str(label.get("Pledge Configuration", "") or "").strip()
    if campaign and pledge:
        return f"{campaign}     Config: {pledge}"
    if pledge:
        return f"Config: {pledge}"
    return campaign


def _label_total_items_footer(label: dict) -> str:
    total = str(label.get("Total Units", "") or "").strip()
    return f"Item Count: {total}" if total else "Item Count:"


def _label_phone_text(label: dict) -> str:
    phone = str(label.get("Phone", "") or "").strip()
    return f"phone: {phone}" if phone else "phone:"


def _continuation_original_label_number(label: dict) -> str:
    original = str(label.get("Original Label Number", "") or "").strip()
    if original:
        return original
    return str(label.get("Label Number", "") or "").replace("CONTINUED", "").strip()


def _label_config_text(label: dict) -> str:
    config = str(label.get("Pledge Configuration", "") or "").strip()
    return f"Config: {config}" if config else ""


def _label_visible_id(label: dict) -> str:
    value = str(
        label.get("Barcode/QR Value")
        or label.get("Label Value")
        or label.get("Label Number")
        or ""
    ).strip()
    match = re.match(r"^(?P<project>.+?)\s+(?P<number>\d+(?:-\d+)?)$", value)
    if match:
        return f"{match.group('number')} {match.group('project')}"
    return value


def _label_header_country_code(label: dict) -> str:
    for key in ["Country Code", "Country Package Code"]:
        match = re.match(r"\s*([A-Za-z]{2})", str(label.get(key, "") or ""))
        if match:
            return match.group(1).upper()
    return ""


def _label_continuation_header_text(label: dict) -> str:
    config = _label_config_text(label)
    return f"CONTINUED  {config}" if config else "CONTINUED"


def _clean_label_state(value: object) -> str:
    text = str(value or "").strip()
    return "" if text.casefold() == "choose a state" else text


def _label_city_state_zip(label: dict) -> str:
    city = str(label.get("City", "") or "").strip()
    state = _clean_label_state(label.get("State/Province", ""))
    postal = str(label.get("Postal/Zip", "") or "").strip()
    location = ", ".join(part for part in [city, state] if part)
    if location and postal:
        return f"{location} {postal}" if state else f"{location}, {postal}"
    return location or postal


def _label_address_split_index(address: str) -> int:
    window_end = min(len(address), LABEL_ADDRESS_SPLIT_MAX)
    candidates = [
        index
        for index, character in enumerate(address[:window_end])
        if index >= LABEL_ADDRESS_SPLIT_MIN and character in {",", " "}
    ]
    if candidates:
        return max(candidates) + 1
    return min(len(address), LABEL_ADDRESS_SPLIT_MAX)


def _label_display_addresses(label: dict) -> tuple[str, str, bool]:
    address_1 = str(label.get("Address Line 1", "") or "").strip()
    address_2 = str(label.get("Address Line 2", "") or "").strip()
    if len(address_1) <= LABEL_ADDRESS_SPLIT_MAX:
        return address_1, address_2, False

    split_index = _label_address_split_index(address_1)
    first_line = address_1[:split_index].strip()
    overflow = address_1[split_index:].strip(" ,")
    if overflow and address_2:
        second_line = f"{overflow}, {address_2}"
    else:
        second_line = overflow or address_2
    return first_line, second_line, len(second_line) > LABEL_ADDRESS_SPLIT_MAX


def _label_block_rows(label: dict) -> list[list[object]]:
    items_column_1 = str(label.get("Items to Pack Column 1", "") or "").splitlines()
    items_column_2 = str(label.get("Items to Pack Column 2", "") or "").splitlines()
    country_code = _label_header_country_code(label)
    if label.get("Label Continuation"):
        original_label = _continuation_original_label_number(label)
        continuation_items = [item for item in [*items_column_1, *items_column_2] if str(item or "").strip()]
        rows = [
            [_label_visible_id({"Barcode/QR Value": original_label}), "", "", "", country_code, ""],
            ["Continuation for", f"{original_label}  {_label_continuation_header_text(label)}", "", "", "", ""],
            ["Qty", "SKU", "", "Qty", "SKU", ""],
        ]
        for index in range(0, len(continuation_items), 2):
            left_item = continuation_items[index]
            right_item = continuation_items[index + 1] if index + 1 < len(continuation_items) else ""
            rows.append(
                [
                    _label_item_qty(left_item),
                    _label_item_sku(left_item),
                    "",
                    _label_item_qty(right_item),
                    _label_item_sku(right_item),
                    "",
                ]
            )
        rows.append(["", "", "", "", "", ""])
        return rows
    item_line_count = max(len(items_column_1), len(items_column_2), 1)
    from_line_1, from_line_2 = _split_label_from_line(label.get("From", ""))
    notes = str(label.get("Notes", "") or "").strip()
    address_1, address_2, address_needs_notes_space = _label_display_addresses(label)
    notes_text = f"Notes: {notes}" if notes and not address_needs_notes_space else ""
    rows = [
        [_label_visible_id(label), "", "", "", country_code, ""],
        ["Origin:", label.get("Origin", ""), "", "", "", ""],
        ["", "", "", "", "", ""],
        ["From", from_line_1, "", "", "", ""],
        ["", from_line_2, "", "", "", ""],
        ["", "", "", "", "", ""],
        ["To", _label_to_name_with_backer(label), "", "", notes_text, ""],
        ["Address 1", address_1, "", "", "", ""],
        ["Address 2", address_2, "", "", "", ""],
        [
            "City/State/Zip",
            _label_city_state_zip(label),
            "",
            "",
            _label_phone_text(label),
            "",
        ],
        ["Country", label.get("Country", ""), "", "", "", ""],
        ["Detailed description of contents: Board Games-of paper and plastic,non-electrical", "", "", "", "", ""],
        ["Qty", "SKU", "", "Qty", "SKU", ""],
    ]
    for index in range(item_line_count):
        rows.append(
            [
                _label_item_qty(items_column_1[index]) if index < len(items_column_1) else "",
                _label_item_sku(items_column_1[index]) if index < len(items_column_1) else "",
                "",
                _label_item_qty(items_column_2[index]) if index < len(items_column_2) else "",
                _label_item_sku(items_column_2[index]) if index < len(items_column_2) else "",
                "",
            ]
        )
    rows.append(["On Arrival: If shipping box is damaged please take photos.", "", "", "", "", ""])
    rows.append([label.get("On Arrival Note", ""), "", "", "", "", ""])
    rows.append(["", "", "", "", "", ""])
    rows.append([_label_campaign_pledge_footer(label), "", "", "", "", ""])
    rows.append([_label_total_items_footer(label), "", "", "", label.get("Carton Box Designation", ""), ""])
    rows.append([label.get("Factory Name", ""), "", "", "", label.get("Country Name Chinese", ""), ""])
    return rows


def _label_item_qty(item: str) -> str:
    text = str(item or "").strip()
    if " x" not in text:
        return ""
    qty = text.rsplit(" x", 1)[1].strip()
    return qty if qty.replace(".", "", 1).isdigit() else ""


def _label_item_sku(item: str) -> str:
    text = str(item or "").strip()
    if " x" not in text:
        return text
    sku, qty = text.rsplit(" x", 1)
    return sku.strip() if qty.strip().replace(".", "", 1).isdigit() else text


def _label_item_block_end_offset(block_rows: list[list[object]]) -> int:
    for index, row in enumerate(block_rows):
        if row and str(row[0] or "").startswith("On Arrival:"):
            return max(index - 1, 0)
    return len(block_rows) - 1


def _label_detail_offset(block_rows: list[list[object]]) -> int:
    for index, row in enumerate(block_rows):
        if row and str(row[0] or "").startswith("Detailed description of contents:"):
            return index
    return -1


def _label_spacer_offsets(block_rows: list[list[object]]) -> set[int]:
    offsets = set()
    for index, row in enumerate(block_rows):
        if any(str(value or "").strip() for value in row):
            continue
        previous_text = " ".join(str(value or "") for value in block_rows[index - 1]) if index > 0 else ""
        next_text = " ".join(str(value or "") for value in block_rows[index + 1]) if index + 1 < len(block_rows) else ""
        if "Fujian, China" in previous_text and "To" in next_text:
            offsets.add(index)
        if "If product damage is found" in previous_text and ("Config:" in next_text or "Pledge Config:" in next_text):
            offsets.add(index)
    return offsets


def _label_block_cell_style(
    label: dict,
    block_rows: list[list[object]],
    row_offset: int,
    column_index: int,
    value: object,
) -> int | None:
    if label.get("Label Continuation"):
        if row_offset == 0:
            style = 24
        elif row_offset == 1:
            style = 2
        elif row_offset == 2:
            if column_index in {0, 3}:
                style = 11
            elif column_index in {1, 4}:
                style = 12
            else:
                style = 14
        elif row_offset >= 3 and row_offset < len(block_rows) - 1:
            if column_index in {0, 3}:
                style = 13
            elif column_index in {1, 4}:
                style = 14
            else:
                style = 14
        else:
            style = _label_cell_style(row_offset, column_index, value)
        if row_offset == 0 and column_index in {4, 5}:
            return _label_right_aligned_style(style)
        return _label_right_aligned_style(style) if column_index == 5 and str(value or "").strip() else style

    footer_offsets = {len(block_rows) - 3, len(block_rows) - 2, len(block_rows) - 1}
    if row_offset == 0:
        style = 24
        if column_index in {4, 5}:
            return _label_right_aligned_style(style)
        return style
    if row_offset in footer_offsets:
        style = 7
        if row_offset in {len(block_rows) - 2, len(block_rows) - 1} and column_index in {4, 5}:
            return _label_right_aligned_style(style)
        return _label_right_aligned_style(style) if column_index == 5 and str(value or "").strip() else style
    detail_offset = _label_detail_offset(block_rows)
    has_notes = len(block_rows) > 8 and any(str(block_rows[offset][4] or "").startswith("Notes:") for offset in range(6, 9))
    if has_notes and 6 <= row_offset <= 8 and column_index in {4, 5}:
        return 22
    item_block_end = _label_item_block_end_offset(block_rows)
    qty_header_offset = detail_offset + 1 if detail_offset >= 0 else -1
    if detail_offset >= 0 and detail_offset <= row_offset <= item_block_end:
        if row_offset == detail_offset:
            return 10
        if row_offset == qty_header_offset:
            if column_index in {0, 3}:
                return 11
            if column_index in {1, 4}:
                return 12
            return 14
        if row_offset > qty_header_offset:
            if column_index in {0, 3}:
                return 13
            if column_index in {1, 4}:
                return 14
            return 14
        return 14
    style = _label_cell_style(row_offset, column_index, value)
    if (
        style is None
        and str(value or "").strip()
        and row_offset in {1, 3, 4, 6, 7, 8, 9, 10}
        and column_index in {0, 1, 4, 5}
    ):
        style = 21
    return _label_right_aligned_style(style) if column_index == 5 and str(value or "").strip() else style


def _label_right_aligned_style(style: int | None) -> int:
    if style == 7:
        return 17
    if style == 21:
        return 23
    if style == 24:
        return 25
    return 16


def _merge_range(row_index: int, start_column: int, end_column: int) -> str:
    return f"{_column_letter(start_column)}{row_index}:{_column_letter(end_column)}{row_index}"


def _label_merge_ranges(block_rows: list[list[object]], current_row: int) -> list[str]:
    merges = [_merge_range(current_row, 0, 3), _merge_range(current_row, 4, 5)]
    if len(block_rows) > 1 and str(block_rows[1][0] or "") == "Continuation for":
        for offset in range(2, max(len(block_rows) - 1, 2)):
            merges.append(_merge_range(current_row + offset, 1, 2))
            merges.append(_merge_range(current_row + offset, 4, 5))
        return merges
    detail_offset = _label_detail_offset(block_rows)
    item_block_end = _label_item_block_end_offset(block_rows)
    if detail_offset >= 0:
        merges.append(_merge_range(current_row + detail_offset, 0, 5))
        qty_header_offset = detail_offset + 1
        for offset in range(qty_header_offset, item_block_end + 1):
            merges.append(_merge_range(current_row + offset, 1, 2))
            merges.append(_merge_range(current_row + offset, 4, 5))
    if len(block_rows) > 8 and any(str(block_rows[offset][4] or "").startswith("Notes:") for offset in range(6, 9)):
        merges.append(f"E{current_row + 6}:F{current_row + 8}")
    for offset, row in enumerate(block_rows):
        row_index = current_row + offset
        if len(row) > 4 and str(row[4] or "").startswith("phone:"):
            merges.append(_merge_range(row_index, 4, 5))
        if offset == len(block_rows) - 3:
            merges.append(_merge_range(row_index, 0, 5))
        if offset == len(block_rows) - 2:
            merges.append(_merge_range(row_index, 0, 1))
            merges.append(_merge_range(row_index, 4, 5))
        if offset == len(block_rows) - 1:
            if len(row) > 3 and not str(row[2] or "").strip() and not str(row[3] or "").strip():
                merges.append(_merge_range(row_index, 0, 3))
            else:
                merges.append(_merge_range(row_index, 0, 1))
            merges.append(_merge_range(row_index, 4, 5))
    return merges


def _labels_sheet_xml(rows: list[dict], include_drawing: bool = False) -> str:
    labels = rows or [{"Note": "No package labels generated."}]
    row_xml = []
    page_breaks = []
    merge_ranges = []
    current_row = 1
    for label in labels:
        block_rows = _label_block_rows(label)
        merge_ranges.extend(_label_merge_ranges(block_rows, current_row))
        spacer_offsets = _label_spacer_offsets(block_rows)
        for offset, block_row in enumerate(block_rows):
            row_index = current_row + offset
            cells = []
            for column_index, value in enumerate(block_row):
                style = _label_block_cell_style(label, block_rows, offset, column_index, value)
                if value == "" and style is None:
                    continue
                cells.append(
                    _cell_xml(
                        row_index,
                        column_index,
                        value,
                        style=style,
                    )
                )
            footer_offsets = set() if label.get("Label Continuation") else {len(block_rows) - 3, len(block_rows) - 2, len(block_rows) - 1}
            detail_offset = _label_detail_offset(block_rows)
            item_block_end = len(block_rows) - 2 if label.get("Label Continuation") else _label_item_block_end_offset(block_rows)
            qty_header_offset = detail_offset + 1 if detail_offset >= 0 else 2 if label.get("Label Continuation") else -1
            if offset == 0:
                height = ' ht="36" customHeight="1"'
            elif offset in footer_offsets:
                height = ' ht="28" customHeight="1"'
            elif offset in spacer_offsets:
                height = ' ht="10" customHeight="1"'
            elif qty_header_offset >= 0 and qty_header_offset < offset <= item_block_end:
                height = ' ht="22" customHeight="1"'
            elif offset in {2, 3, 4, 6, 7, 8, 9, 10, 12}:
                height = ' ht="22" customHeight="1"'
            else:
                height = ' ht="18" customHeight="1"'
            row_xml.append(f'<row r="{row_index}"{height}>{"".join(cells)}</row>')
        current_row += len(block_rows)
        page_breaks.append(current_row - 1)

    widths = [12, 22, 14, 14, 22, 14]
    cols_xml = "".join(
        f'<col min="{index + 1}" max="{index + 1}" width="{width}" customWidth="1"/>'
        for index, width in enumerate(widths)
    )
    manual_break_count = len(page_breaks) - 1
    if 0 < manual_break_count <= MAX_MANUAL_ROW_BREAKS_PER_SHEET:
        breaks_xml = "".join(
            f'<brk id="{row}" max="16383" man="1"/>'
            for row in page_breaks[:-1]
        )
        row_breaks_xml = f'<rowBreaks count="{manual_break_count}" manualBreakCount="{manual_break_count}">{breaks_xml}</rowBreaks>'
    else:
        row_breaks_xml = ""

    drawing_xml = '<drawing r:id="rId1"/>' if include_drawing else ""
    merge_xml = (
        f'<mergeCells count="{len(merge_ranges)}">'
        + "".join(f'<mergeCell ref="{merge_range}"/>' for merge_range in merge_ranges)
        + "</mergeCells>"
        if merge_ranges
        else ""
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="18"/>'
        f"<cols>{cols_xml}</cols>"
        f'<sheetData>{"".join(row_xml)}</sheetData>'
        f"{merge_xml}"
        '<printOptions horizontalCentered="1"/>'
        '<pageMargins left="0.12" right="0.12" top="0.12" bottom="0.12" header="0.0" footer="0.0"/>'
        '<pageSetup orientation="portrait" fitToWidth="1" fitToHeight="0"/>'
        f"{row_breaks_xml}"
        f"{drawing_xml}"
        "</worksheet>"
    )


def _label_block_lengths(rows: list[dict]) -> list[int]:
    return [len(_label_block_rows(label)) for label in (rows or [{"Note": "No package labels generated."}])]


def _label_qr_images(rows: list[dict]) -> list[tuple[int, bytes]]:
    images = []
    current_row = 1
    for label, block_length in zip(rows, _label_block_lengths(rows), strict=False):
        if label.get("Label Continuation"):
            current_row += block_length
            continue
        value = str(label.get("Barcode/QR Value") or label.get("Label Value") or "").strip()
        if value:
            images.append((current_row, qr_png(value, scale=8, border=4)))
        current_row += block_length
    return images


def _label_qr_image_count(rows: list[dict]) -> int:
    return sum(
        1
        for label in rows
        if not label.get("Label Continuation")
        and str(label.get("Barcode/QR Value") or label.get("Label Value") or "").strip()
    )


def _worksheet_rels_xml(drawing_id: int) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing{drawing_id}.xml"/>'
        "</Relationships>"
    )


def _drawing_xml(image_count: int, start_rows: list[int]) -> str:
    anchors = []
    for index, start_row in enumerate(start_rows, start=1):
        anchors.append(
            '<xdr:oneCellAnchor>'
            '<xdr:from><xdr:col>3</xdr:col><xdr:colOff>457200</xdr:colOff>'
            f'<xdr:row>{max(start_row - 1, 0)}</xdr:row><xdr:rowOff>57150</xdr:rowOff></xdr:from>'
            '<xdr:ext cx="1524000" cy="1524000"/>'
            '<xdr:pic>'
            f'<xdr:nvPicPr><xdr:cNvPr id="{index}" name="QR Code {index}"/><xdr:cNvPicPr/></xdr:nvPicPr>'
            '<xdr:blipFill>'
            f'<a:blip r:embed="rId{index}"/>'
            '<a:stretch><a:fillRect/></a:stretch>'
            '</xdr:blipFill>'
            '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>'
            '</xdr:pic>'
            '<xdr:clientData/>'
            '</xdr:oneCellAnchor>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'{"".join(anchors)}'
        '</xdr:wsDr>'
    )


def _drawing_rels_xml(image_filenames: list[str]) -> str:
    relationships = "".join(
        f'<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/{filename}"/>'
        for index, filename in enumerate(image_filenames, start=1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f"{relationships}"
        "</Relationships>"
    )


def _headers_for_sheet(sheet_name: str, rows: list[dict]) -> list[str]:
    headers = []
    seen = set()

    if sheet_name == "Actual Dimensions":
        headers.extend(ACTUAL_DIMENSIONS_COLUMNS)
        seen.update(ACTUAL_DIMENSIONS_COLUMNS)

    if sheet_name == "Order Volume Weights":
        headers.extend(ORDER_VOLUME_WEIGHTS_COLUMNS)
        seen.update(ORDER_VOLUME_WEIGHTS_COLUMNS)

    for row in rows:
        for key in row:
            if key == "_Country Scan Metadata":
                continue
            if key not in seen:
                seen.add(key)
                headers.append(key)

    return headers




def _is_one_decimal_measure_column(header: str) -> bool:
    normalized = header.lower()
    if any(token in normalized for token in ["fee", "cost", "margin", "price", "rate"]):
        return False
    if normalized.endswith(" g") or normalized.endswith("g") and "kg" not in normalized:
        return False
    measure_tokens = [
        " cm",
        "kg",
        " lb",
        "weight",
        "dimension",
        "placement",
        "length",
        "width",
        "height",
    ]
    return any(token in normalized for token in measure_tokens)


def _format_measure_value(header: str, value: object) -> object:
    if isinstance(value, ExcelFormula):
        return value
    if not _is_one_decimal_measure_column(header):
        return value
    if isinstance(value, bool):
        return value
    if isinstance(value, int | float):
        rounded = round(float(value), 1)
        return int(rounded) if rounded.is_integer() else rounded
    text = str(value).strip()
    if not text:
        return value
    try:
        number = float(text)
    except ValueError:
        return value
    rounded = round(number, 1)
    return int(rounded) if rounded.is_integer() else rounded

def _rows_to_table(sheet_name: str, rows: list[dict]) -> tuple[list[str], list[list[object]]]:
    if not rows:
        if sheet_name == "Actual Dimensions":
            return ACTUAL_DIMENSIONS_COLUMNS, []
        if sheet_name == "Order Volume Weights":
            return ORDER_VOLUME_WEIGHTS_COLUMNS, []
        return ["Note"], [["No records"]]

    headers = _headers_for_sheet(sheet_name, rows)
    display_headers = [_display_header_for_sheet(sheet_name, header) for header in headers]
    values = [
        [_format_measure_value(header, row.get(header, "")) for header in headers]
        for row in rows
    ]
    return display_headers, values


def _column_widths(sheet_name: str, headers: list[str], rows: list[list[object]]) -> list[float]:
    if sheet_name == "Actual Dimensions":
        return [
            4
            if index == 13
            else ACTUAL_DIMENSIONS_COLUMN_WIDTHS.get(header, min(max(len(str(header)) + 2, 14), 18))
            for index, header in enumerate(headers)
        ]

    widths = []
    for index, header in enumerate(headers):
        if sheet_name.startswith("Cost Summary") and header == "Final weight kg":
            widths.append(15)
            continue
        if sheet_name.startswith("Cost Summary") and header == "Final cost":
            widths.append(12)
            continue
        if sheet_name.startswith("Cost Summary") and header == "Scan note":
            widths.append(16)
            continue
        if headers[3:7] == ["Actual weight g", "Actual DIM L", "Actual DIM W", "Actual DIM H"] and index == 3:
            widths.append(13)
            continue
        if headers[3:7] == ["Actual weight g", "Actual DIM L", "Actual DIM W", "Actual DIM H"] and 4 <= index <= 6:
            widths.append(13)
            continue
        longest = len(str(header))
        for row in rows:
            if index < len(row):
                longest = max(longest, len(str(row[index])))
        widths.append(min(max(longest + 2, 12), 40))
    return widths


def _worksheet_column_xml(sheet_name: str, widths: list[float]) -> str:
    column_xml = []
    for index, width in enumerate(widths):
        column_number = index + 1
        hidden = (
            ' hidden="1" collapsed="1"'
            if sheet_name == "VFI Intake Form" and 12 <= column_number <= 27
            else ""
        )
        if sheet_name in {"_ActualLookupTable", "_ActualRateTable"}:
            hidden = ' hidden="1" collapsed="1"'
        column_xml.append(
            f'<col min="{column_number}" max="{column_number}" width="{width}" customWidth="1"{hidden}/>'
        )
    return "".join(column_xml)


def _worksheet_xml(sheet_name: str, rows: list[dict]) -> str:
    headers, values = _rows_to_table(sheet_name, rows)
    table = [headers, *values]
    row_xml = []
    for row_index, row in enumerate(table, start=1):
        cells = [
            _cell_xml(
                row_index,
                column_index,
                value,
                style=_cell_style_for_sheet_value(sheet_name, row_index, column_index, headers, row),
            )
            for column_index, value in enumerate(row)
        ]
        row_xml.append(f'<row r="{row_index}">{"".join(cells)}</row>')

    widths = _column_widths(sheet_name, headers, values)
    cols_xml = _worksheet_column_xml(sheet_name, widths)
    last_column = _column_letter(len(headers) - 1)
    last_row = max(len(table), 1)
    filter_ref = f"A1:{last_column}{last_row}"

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheetViews><sheetView workbookViewId="0">'
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>'
        "</sheetView></sheetViews>"
        f"<cols>{cols_xml}</cols>"
        f'<sheetData>{"".join(row_xml)}</sheetData>'
        f'<autoFilter ref="{filter_ref}"/>'
        "</worksheet>"
    )


def _sheet_range(sheet_name: str, column: str, start_row: int, end_row: int) -> str:
    quoted = str(sheet_name).replace("'", "''")
    return f"'{quoted}'!${column}${start_row}:${column}${end_row}"


def _invoice_formula_ranges(payload: InvoicePayload) -> tuple[str, str, str]:
    if payload.cost_summary_data_end_row < payload.cost_summary_data_start_row:
        return "", "", ""
    final_cost = (
        _sheet_range(
            payload.cost_summary_sheet_name,
            payload.final_cost_column,
            payload.cost_summary_data_start_row,
            payload.cost_summary_data_end_row,
        )
        if payload.final_cost_column
        else ""
    )
    final_weight = (
        _sheet_range(
            payload.cost_summary_sheet_name,
            payload.final_weight_column,
            payload.cost_summary_data_start_row,
            payload.cost_summary_data_end_row,
        )
        if payload.final_weight_column
        else ""
    )
    scan_note = (
        _sheet_range(
            payload.cost_summary_sheet_name,
            payload.scan_note_column,
            payload.cost_summary_data_start_row,
            payload.cost_summary_data_end_row,
        )
        if payload.scan_note_column
        else ""
    )
    return final_cost, final_weight, scan_note


def _truncate_money_value(value: object) -> object:
    if isinstance(value, bool):
        return value
    if isinstance(value, int | float):
        truncated = int(float(value) * 100) / 100
        return round(truncated, 2)
    return value


def _invoice_display_line_cells(line: object) -> tuple[dict[int, object], bool]:
    text = str(line)
    if not text:
        return {0: ""}, False
    lowered = text.lower()
    if not lowered.startswith(("http://", "https://")) and ":" in text:
        label, value = text.split(":", 1)
        return {0: f"{label.strip()}:", 1: value.strip()}, False
    return {0: text}, True


def _invoice_sheet_xml(payload: InvoicePayload) -> str:
    from box_optimizer.invoice import INVOICE_WARNING_TEXT, PAY_TO_INCOMPLETE_WARNING_TEXT

    final_cost_range, _, scan_note_range = _invoice_formula_ranges(payload)
    shipping_fee_formula = f"ROUNDDOWN(SUM({final_cost_range}),2)" if final_cost_range else '""'
    if final_cost_range:
        warning_conditions = [f'SUMPRODUCT(--(LEN(TRIM({final_cost_range}&""))=0))>0']
        if scan_note_range:
            warning_conditions.append(f'COUNTIF({scan_note_range},"<>")>0')
        warning_formula = f'IF(OR({",".join(warning_conditions)}),"{INVOICE_WARNING_TEXT}","")'
    else:
        warning_formula = f'"{INVOICE_WARNING_TEXT}"'

    address_lines = list(payload.address_lines[:2])
    while len(address_lines) < 2:
        address_lines.append("")

    charge_rows: list[tuple[str, object]] = [
        ("shipping fee:", ExcelFormula(shipping_fee_formula)),
        ("inbound fee:", _truncate_money_value(payload.inbound_fee)),
    ]
    if payload.include_canada_ocean_tax:
        charge_rows.append(("Canada Ocean and Tax", ""))
    if payload.include_mx_import_tax:
        charge_rows.append(("MX import tax", ""))
    first_charge_row = 14
    total_row = first_charge_row + len(charge_rows)
    warning_row = total_row + 1
    payment_header_row = warning_row + 1
    pay_to_start = payment_header_row + 1
    total_formula = f"ROUNDDOWN(SUM({','.join(f'C{row}' for row in range(first_charge_row, total_row))}),2)"

    rows: dict[int, dict[int, object]] = {
        1: {0: "VFI Asia"},
        2: {0: "3513 Kensett Way, Raleigh, NC 27612 USA"},
        4: {1: "COMMERCIAL INVOICE"},
        6: {0: "Invoice No:", 1: payload.invoice_number},
        7: {0: "Campaign:", 1: payload.campaign_name},
        8: {0: "invoice to:", 1: payload.bill_to},
        9: {1: address_lines[0]},
        10: {1: address_lines[1]},
        11: {0: "Date:", 1: f"{payload.invoice_date:%m/%d/%Y}"},
        12: {0: "email", 1: payload.email},
        total_row: {1: "Total Due:", 2: ExcelFormula(total_formula), 3: "USD"},
        warning_row: {0: ExcelFormula(warning_formula)},
        payment_header_row: {0: "Payments can be made to Any of the below options:"},
    }
    for offset, (label, value) in enumerate(charge_rows):
        row_index = first_charge_row + offset
        rows[row_index] = {
            1: label,
            2: value,
            3: "USD",
        }
        if row_index == first_charge_row:
            rows[row_index][0] = "ship order:"
        elif row_index == first_charge_row + 1:
            rows[row_index][0] = payload.ship_order_count
    for offset, address_line in enumerate(payload.address_lines[2:], start=1):
        rows.setdefault(10 + offset, {})[2] = address_line
    display_line_merge_rows = []
    for offset, (label, value) in enumerate(payload.pay_to_lines):
        row_index = pay_to_start + offset
        if payload.pay_to_display_lines:
            cells, merge_full_line = _invoice_display_line_cells(label)
            rows.setdefault(row_index, {}).update(cells)
            if merge_full_line:
                display_line_merge_rows.append(row_index)
        else:
            rows.setdefault(row_index, {})[0] = label
            rows[row_index][1] = value
    if payload.pay_to_incomplete:
        row_index = pay_to_start + len(payload.pay_to_lines)
        rows.setdefault(row_index, {})[0] = PAY_TO_INCOMPLETE_WARNING_TEXT

    row_xml = []
    for row_index in range(1, max(rows) + 1):
        cells = []
        for column_index in range(7):
            if column_index not in rows.get(row_index, {}):
                continue
            value = rows[row_index][column_index]
            style = None
            if row_index in {4, total_row, warning_row, payment_header_row}:
                style = 2
            if first_charge_row <= row_index <= total_row and column_index == 2:
                style = 19
            if row_index == 1:
                style = 32
            if row_index == 2:
                style = 33
            if row_index == first_charge_row + 1 and column_index == 0:
                style = 8
            if row_index == warning_row:
                style = 26
            cells.append(_cell_xml(row_index, column_index, value, style=style))
        row_xml.append(f'<row r="{row_index}">{"".join(cells)}</row>')

    cols_xml = "".join(
        f'<col min="{index}" max="{index}" width="{width}" customWidth="1"/>'
        for index, width in enumerate([18, 22, 18, 12, 16, 16, 16], start=1)
    )
    merge_ranges = [
        "A1:D1",
        "A2:D2",
        "B4:D4",
        "B8:D8",
        "B9:D9",
        "B10:D10",
        "B12:D12",
        f"A{warning_row}:G{warning_row}",
        f"A{payment_header_row}:G{payment_header_row}",
        *[f"A{row_index}:D{row_index}" for row_index in display_line_merge_rows],
    ]
    merge_xml = (
        f'<mergeCells count="{len(merge_ranges)}">'
        + "".join(f'<mergeCell ref="{ref}"/>' for ref in merge_ranges)
        + "</mergeCells>"
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<cols>{cols_xml}</cols>"
        f'<sheetData>{"".join(row_xml)}</sheetData>'
        f"{merge_xml}"
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
        '<pageSetup orientation="portrait" fitToWidth="1" fitToHeight="0"/>'
        "</worksheet>"
    )


def _workbook_xml(sheet_names: list[str]) -> str:
    sheets = "".join(
        f'<sheet name="{escape(name)}" sheetId="{index}" r:id="rId{index}"'
        f'{" state=" + chr(34) + "hidden" + chr(34) if name in {"_ActualLookupTable", "_ActualRateTable"} else ""}/>'
        for index, name in enumerate(sheet_names, start=1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<sheets>{sheets}</sheets>"
        '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>'
        "</workbook>"
    )


def _workbook_rels(sheet_count: int) -> str:
    worksheet_rels = "".join(
        f'<Relationship Id="rId{index}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        f'Target="worksheets/sheet{index}.xml"/>'
        for index in range(1, sheet_count + 1)
    )
    styles_id = sheet_count + 1
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f"{worksheet_rels}"
        f'<Relationship Id="rId{styles_id}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
        "</Relationships>"
    )


def _root_rels() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        "</Relationships>"
    )


def _content_types(sheet_count: int, drawing_count: int = 0) -> str:
    overrides = "".join(
        f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for index in range(1, sheet_count + 1)
    )
    drawing_overrides = "".join(
        f'<Override PartName="/xl/drawings/drawing{index}.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
        for index in range(1, drawing_count + 1)
    )
    png_default = '<Default Extension="png" ContentType="image/png"/>' if drawing_count else ""
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f"{png_default}"
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        f"{overrides}"
        f"{drawing_overrides}"
        "</Types>"
    )


def _styles_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<numFmts count="3">'
        '<numFmt numFmtId="164" formatCode="$#,##0.00&quot; (USD)&quot;"/>'
        '<numFmt numFmtId="165" formatCode="$#,##0.00"/>'
        '<numFmt numFmtId="166" formatCode="0.0"/>'
        '</numFmts>'
        '<fonts count="12">'
        '<font><sz val="11"/><name val="Calibri"/></font>'
        '<font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font>'
        '<font><b/><sz val="12"/><name val="Calibri"/></font>'
        '<font><sz val="18"/><name val="Calibri"/></font>'
        '<font><b/><sz val="16"/><name val="Calibri"/></font>'
        '<font><b/><sz val="16"/><name val="Calibri"/></font>'
        '<font><b/><sz val="25"/><name val="Calibri"/></font>'
        '<font><b/><sz val="16"/><name val="Calibri"/></font>'
        '<font><color rgb="FFFF0000"/><sz val="11"/><name val="Calibri"/></font>'
        '<font><sz val="15"/><name val="Calibri"/></font>'
        '<font><b/><sz val="31"/><name val="Calibri"/></font>'
        '<font><b/><color rgb="FFFF0000"/><sz val="11"/><name val="Calibri"/></font>'
        "</fonts>"
        '<fills count="4">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFD9E2F3"/><bgColor indexed="64"/></patternFill></fill>'
        "</fills>"
        '<borders count="3">'
        '<border><left/><right/><top/><bottom/><diagonal/></border>'
        '<border><left/><right/><top style="thick"><color auto="1"/></top><bottom style="thick"><color auto="1"/></bottom><diagonal/></border>'
        '<border><left style="thick"><color auto="1"/></left><right style="thick"><color auto="1"/></right><top style="thick"><color auto="1"/></top><bottom style="thick"><color auto="1"/></bottom><diagonal/></border>'
        '</borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="34">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
        '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
        '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center"/></xf>'
        '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
        '<xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center"/></xf>'
        '<xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
        '<xf numFmtId="0" fontId="6" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="left"/></xf>'
        '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
        '<xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>'
        '<xf numFmtId="0" fontId="4" fillId="0" borderId="2" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>'
        '<xf numFmtId="0" fontId="5" fillId="0" borderId="2" xfId="0" applyFont="1" applyBorder="1"/>'
        '<xf numFmtId="0" fontId="3" fillId="0" borderId="2" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>'
        '<xf numFmtId="0" fontId="3" fillId="0" borderId="2" xfId="0" applyFont="1" applyBorder="1"/>'
        '<xf numFmtId="0" fontId="6" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="right"/></xf>'
        '<xf numFmtId="0" fontId="6" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="right"/></xf>'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>'
        '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
        '<xf numFmtId="0" fontId="8" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
        '<xf numFmtId="0" fontId="9" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
        '<xf numFmtId="0" fontId="9" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="9" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="right"/></xf>'
        '<xf numFmtId="0" fontId="10" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
        '<xf numFmtId="0" fontId="10" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="right"/></xf>'
        '<xf numFmtId="0" fontId="11" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="top" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="top" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="2" fillId="3" borderId="2" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top" wrapText="1"/></xf>'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1"/>'
        '<xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
        '<xf numFmtId="0" fontId="10" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center"/></xf>'
        '<xf numFmtId="0" fontId="9" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center"/></xf>'
        "</cellXfs>"
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        "</styleSheet>"
    )


def _build_sheet_payloads(
    rows: list[dict] | None,
    sheets: dict[str, list[dict]] | None,
    **named_rows: list[dict],
) -> list[tuple[str, object]]:
    country_scan_sheets = named_rows.pop("country_scan_sheets", None) or {}
    invoice_payload = named_rows.pop("invoice_payload", None)
    output_mode = _normalized_workbook_output_mode(named_rows.pop("workbook_output_mode", "full"))
    payloads = {name: [] for name in REQUIRED_SHEETS}
    if rows is not None:
        payloads["Summary"] = rows
    if sheets:
        payloads.update(sheets)

    aliases = {
        "summary_rows": "Summary",
        "cost_summary_rows": "Cost Summary",
        "stock_count_rows": "Stock Count",
        "actual_dimensions_rows": "Actual Dimensions",
        "actual_lookup_rows": "_ActualLookupTable",
        "actual_rate_rows": "_ActualRateTable",
        "vfi_intake_form_rows": "VFI Intake Form",
        "optimized_to_pack_rows": "Optimized to Pack",
        "label_generator_rows": "Label generator",
        "labels_rows": "Labels",
        "order_volume_weights_rows": "Order Volume Weights",
        "box_size_summary_rows": "Box Size Summary",
        "unmatched_skus_rows": "Unmatched SKUs",
        "packing_detail_rows": "Packing Detail",
        "multi_box_detail_rows": "Multi Box Detail",
        "pledge_combination_summary_rows": "Pledge Combination Summary",
        "debug_summary_rows": "Debug Summary",
        "input_column_mapping_rows": "Input Column Mapping",
        "errors_and_warnings_rows": "Errors and Warnings",
    }
    for key, sheet_name in aliases.items():
        if key in named_rows and named_rows[key] is not None:
            payloads[sheet_name] = named_rows[key]

    ordered = [(name, payloads.get(name, [])) for name in REQUIRED_SHEETS]
    ordered = _split_label_sheet_payloads(ordered)
    required_and_optional = {*REQUIRED_SHEETS, *OPTIONAL_SHEETS}
    for name, sheet_rows in payloads.items():
        if name.startswith("Cost Summary -") and sheet_rows:
            ordered[1] = (name, sheet_rows)
            break
    for name in OPTIONAL_SHEETS:
        if payloads.get(name):
            ordered.append((name, payloads[name]))

    if country_scan_sheets:
        labels_indexes = [
            index
            for index, (name, _rows) in enumerate(ordered)
            if _is_labels_sheet_name(name)
        ]
        labels_index = labels_indexes[-1] if labels_indexes else None
        insert_index = len(ordered) if labels_index is None else labels_index + 1
        for offset, (name, sheet_rows) in enumerate(country_scan_sheets.items()):
            if sheet_rows:
                ordered.insert(insert_index + offset, (name, _country_scan_rows_with_pallet_id(sheet_rows)))
                required_and_optional.add(name)

    for name, sheet_rows in payloads.items():
        if name not in required_and_optional and not name.startswith("Cost Summary -") and sheet_rows:
            ordered.append((name, sheet_rows))

    if isinstance(invoice_payload, InvoicePayload) and normalize_invoice_variant(invoice_payload.variant):
        actual_dimensions_index = next(
            (index for index, (name, _rows) in enumerate(ordered) if name == "Actual Dimensions"),
            None,
        )
        insert_index = len(ordered) if actual_dimensions_index is None else actual_dimensions_index
        ordered.insert(insert_index, ("Invoice", invoice_payload))

    if output_mode == "fast_production":
        operational_sheet_names = {
            "Summary",
            "Cost Summary",
            "Stock Count",
            "Actual Dimensions",
            "_ActualLookupTable",
            "_ActualRateTable",
            "VFI Intake Form",
            "Optimized to Pack",
            "Box Size Summary",
            "Errors and Warnings",
            "Invoice",
            *country_scan_sheets.keys(),
        }
        ordered = [
            (name, sheet_rows)
            for name, sheet_rows in ordered
            if name in operational_sheet_names or _is_labels_sheet_name(name) or name.startswith("Cost Summary -")
        ]

    return ordered


def _country_scan_rows_with_pallet_id(rows: list[dict]) -> list[dict]:
    output = []
    for row in rows:
        if "Pallet ID" in row:
            output.append(row)
        else:
            output.append({"Pallet ID": "", **row})
    return output


def _is_labels_sheet_name(name: str) -> bool:
    return name == "Labels" or (name.startswith("Labels ") and name.removeprefix("Labels ").isdigit())


def _split_label_sheet_payloads(
    ordered: list[tuple[str, list[dict]]],
) -> list[tuple[str, list[dict]]]:
    output = []
    for name, sheet_rows in ordered:
        if name != "Labels" or len(sheet_rows) <= MAX_LABELS_PER_SHEET:
            output.append((name, sheet_rows))
            continue
        for index in range(0, len(sheet_rows), MAX_LABELS_PER_SHEET):
            chunk = sheet_rows[index : index + MAX_LABELS_PER_SHEET]
            sheet_number = index // MAX_LABELS_PER_SHEET + 1
            sheet_name = "Labels" if sheet_number == 1 else f"Labels {sheet_number}"
            output.append((sheet_name, chunk))
    return output


def _normalized_workbook_output_mode(value: object) -> str:
    mode = str(value or "full").strip().lower()
    if mode in {"admin", "full"}:
        return "full"
    if mode in {"worker", "fast_production"}:
        return "fast_production"
    return "full"


def workbook_sheet_stats(
    rows: list[dict] | None = None,
    sheets: dict[str, list[dict]] | None = None,
    **named_rows: list[dict],
) -> dict:
    """Return sheet-writing stats for a workbook payload without rendering XML."""
    output_mode = _normalized_workbook_output_mode(named_rows.get("workbook_output_mode", "full"))
    full_named_rows = dict(named_rows)
    full_named_rows["workbook_output_mode"] = "full"
    full_payloads = _build_sheet_payloads(rows, sheets, **full_named_rows)
    filtered_payloads = _build_sheet_payloads(rows, sheets, **named_rows)
    full_names = [name for name, _sheet_rows in full_payloads]
    filtered_names = [name for name, _sheet_rows in filtered_payloads]
    skipped = [name for name in full_names if name not in filtered_names]
    labels_rows = [
        label
        for name, sheet_rows in filtered_payloads
        if _is_labels_sheet_name(name) and isinstance(sheet_rows, list)
        for label in sheet_rows
    ]
    country_scan_sheets = named_rows.get("country_scan_sheets") or {}
    return {
        "workbook_output_mode": output_mode,
        "sheets_written": filtered_names,
        "sheets_skipped": skipped,
        "sheets_written_count": len(filtered_names),
        "sheets_skipped_count": len(skipped),
        "country_sheet_count": sum(1 for _name, sheet_rows in country_scan_sheets.items() if sheet_rows),
        "qr_images_written": _label_qr_image_count(labels_rows),
    }


def _safe_sheet_names(names: list[str]) -> list[str]:
    safe_names = []
    used = set()
    for name in names:
        base = _safe_sheet_name(name)
        candidate = base
        suffix = 2
        while candidate in used:
            suffix_text = f" {suffix}"
            candidate = _safe_sheet_name(f"{base[:31 - len(suffix_text)]}{suffix_text}")
            suffix += 1
        used.add(candidate)
        safe_names.append(candidate)
    return safe_names


def write_workbook(
    path: str,
    rows: list[dict] | None = None,
    sheets: dict[str, list[dict]] | None = None,
    **named_rows: list[dict],
) -> str:
    """Write an XLSX workbook with required box optimizer report tabs."""
    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    sheet_payloads = _build_sheet_payloads(rows, sheets, **named_rows)
    safe_names = _safe_sheet_names([name for name, _ in sheet_payloads])
    label_drawing_entries = []
    next_media_index = 1
    for sheet_index, (sheet_name, sheet_rows) in enumerate(sheet_payloads, start=1):
        if not _is_labels_sheet_name(sheet_name) or not isinstance(sheet_rows, list):
            continue
        qr_images = _label_qr_images(sheet_rows)
        if not qr_images:
            continue
        drawing_id = len(label_drawing_entries) + 1
        image_filenames = []
        for _start_row, _image in qr_images:
            image_filenames.append(f"label_qr_{next_media_index}.png")
            next_media_index += 1
        label_drawing_entries.append(
            {
                "sheet_index": sheet_index,
                "drawing_id": drawing_id,
                "images": qr_images,
                "image_filenames": image_filenames,
            }
        )
    drawing_id_by_sheet_index = {
        entry["sheet_index"]: entry["drawing_id"]
        for entry in label_drawing_entries
    }

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", _content_types(len(sheet_payloads), drawing_count=len(label_drawing_entries)))
        archive.writestr("_rels/.rels", _root_rels())
        archive.writestr("xl/workbook.xml", _workbook_xml(safe_names))
        archive.writestr("xl/_rels/workbook.xml.rels", _workbook_rels(len(sheet_payloads)))
        archive.writestr("xl/styles.xml", _styles_xml())
        for index, (sheet_name, sheet_rows) in enumerate(sheet_payloads, start=1):
            sheet_xml = (
                _labels_sheet_xml(sheet_rows, include_drawing=index in drawing_id_by_sheet_index)
                if _is_labels_sheet_name(sheet_name) and isinstance(sheet_rows, list)
                else _invoice_sheet_xml(sheet_rows)
                if sheet_name == "Invoice" and isinstance(sheet_rows, InvoicePayload)
                else _worksheet_xml(sheet_name, sheet_rows)
            )
            archive.writestr(
                f"xl/worksheets/sheet{index}.xml",
                sheet_xml,
            )
        for entry in label_drawing_entries:
            sheet_index = entry["sheet_index"]
            drawing_id = entry["drawing_id"]
            images = entry["images"]
            image_filenames = entry["image_filenames"]
            archive.writestr(f"xl/worksheets/_rels/sheet{sheet_index}.xml.rels", _worksheet_rels_xml(drawing_id))
            archive.writestr(
                f"xl/drawings/drawing{drawing_id}.xml",
                _drawing_xml(len(images), [start_row for start_row, _image in images]),
            )
            archive.writestr(f"xl/drawings/_rels/drawing{drawing_id}.xml.rels", _drawing_rels_xml(image_filenames))
            for filename, (_start_row, image) in zip(image_filenames, images, strict=True):
                archive.writestr(f"xl/media/{filename}", image)

    return str(output_path)
