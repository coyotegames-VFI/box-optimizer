"""Excel and CSV input helpers."""

import csv
import re
import zipfile
from dataclasses import dataclass, field, replace
from pathlib import Path
from xml.etree import ElementTree

from box_optimizer.io.column_mapper import (
    infer_columns,
    infer_dimension_unit,
    infer_weight_unit,
    is_metadata_column,
)
from box_optimizer.models import OrderLine, SKUItem, UnmatchedSKURecord
from box_optimizer.normalize import normalize_dimensions, normalize_sku
from box_optimizer.weights import normalize_weight


@dataclass(frozen=True)
class SourceRows:
    """Rows read from a source sheet or CSV."""

    sheet_name: str
    rows: list[dict]


@dataclass(frozen=True)
class IntakeResult:
    """Matched intake data with unmatched SKU records preserved."""

    sku_items: list[SKUItem]
    order_lines: list[OrderLine]
    matched_order_lines: list[OrderLine]
    unmatched_skus: list[UnmatchedSKURecord]
    column_mappings: list[dict] = field(default_factory=list)
    debug: dict = field(default_factory=dict)


_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def _parse_number(value: object, default: float = 0.0) -> float:
    if value is None or value == "":
        return default
    match = re.search(
        r"-?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?",
        str(value).replace(",", ""),
    )
    if not match:
        return default
    return float(match.group(0))


def _quantity_value(value: object) -> int | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    lowered = text.lower()
    if lowered in {"0", "0.0", "no", "n", "false"}:
        return None
    if lowered in {"yes", "y", "true"}:
        return 1
    match = re.fullmatch(r"(?:x\s*)?(\d+)(?:\s*x)?", lowered)
    if match:
        quantity = int(match.group(1))
        return quantity if quantity > 0 else None
    if re.fullmatch(r"\d+(?:\.0+)?", lowered):
        quantity = int(float(lowered))
        return quantity if quantity > 0 else None
    return None


def _parse_quantity(value: object) -> int:
    quantity = _quantity_value(value)
    return 1 if quantity is None else max(quantity, 1)


def _looks_like_quantity(value: object) -> bool:
    if value is None:
        return True
    text = str(value).strip()
    if not text:
        return True
    return _quantity_value(text) is not None or text.lower() in {
        "0",
        "0.0",
        "no",
        "n",
        "false",
    }


def _parse_dimensions(value: object, default_unit: str = "cm"):
    text = str(value or "").strip()
    if not text:
        return None

    unit_match = re.search(r"(mm|cm|in|ft)\b", text, flags=re.IGNORECASE)
    unit = unit_match.group(1).lower() if unit_match else default_unit
    parts = re.split(r"\s*(?:x|X|×|\*|/)\s*", text)
    numbers = []
    for part in parts:
        match = re.search(r"-?\d+(?:\.\d+)?", part.replace(",", ""))
        if match:
            numbers.append(float(match.group(0)))
    if len(numbers) < 2:
        return None
    if len(numbers) == 2:
        return normalize_dimensions(numbers[0], numbers[1], None, unit=unit)
    return normalize_dimensions(numbers[0], numbers[1], numbers[2], unit=unit)


def _metadata(row: dict, source_path: str, sheet_name: str) -> dict:
    metadata = dict(row)
    metadata["_source_file"] = str(source_path)
    metadata["_source_sheet"] = sheet_name
    return metadata


def _mapping_record(
    workbook: str,
    sheet: str,
    detected_format: str,
    metadata_columns: list[str] | None = None,
    product_quantity_columns: list[str] | None = None,
    dimension_column: str | None = None,
    weight_column: str | None = None,
    warnings: list[str] | None = None,
) -> dict:
    return {
        "workbook": workbook,
        "sheet": sheet,
        "detected format": detected_format,
        "detected metadata columns": " | ".join(metadata_columns or []),
        "detected product quantity columns": " | ".join(product_quantity_columns or []),
        "detected dimension column": dimension_column or "",
        "detected weight column": weight_column or "",
        "warnings": " | ".join(warnings or []),
    }


def _read_csv(path: str) -> list[SourceRows]:
    with open(path, newline="", encoding="utf-8-sig") as file:
        reader = csv.DictReader(file)
        rows = [dict(row) for row in reader]
    return [SourceRows(sheet_name=Path(path).stem, rows=rows)] if rows else []


def _column_index(cell_reference: str) -> int:
    letters = re.sub(r"[^A-Z]", "", cell_reference.upper())
    index = 0
    for letter in letters:
        index = index * 26 + ord(letter) - ord("A") + 1
    return index - 1


def _xlsx_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ElementTree.fromstring(archive.read("xl/sharedStrings.xml"))
    values = []
    for item in root.findall("main:si", _NS):
        text_parts = [node.text or "" for node in item.findall(".//main:t", _NS)]
        values.append("".join(text_parts))
    return values


def _xlsx_sheet_paths(archive: zipfile.ZipFile) -> list[tuple[str, str]]:
    workbook = ElementTree.fromstring(archive.read("xl/workbook.xml"))
    rels = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    rel_targets = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels.findall("pkgrel:Relationship", _NS)
    }
    sheets = []
    for sheet in workbook.findall("main:sheets/main:sheet", _NS):
        name = sheet.attrib["name"]
        rel_id = sheet.attrib[f"{{{_NS['rel']}}}id"]
        target = rel_targets[rel_id].lstrip("/")
        path = target if target.startswith("xl/") else f"xl/{target}"
        sheets.append((name, path))
    return sheets


def _xlsx_cell_value(cell, shared_strings: list[str]) -> str:
    value_node = cell.find("main:v", _NS)
    inline_node = cell.find("main:is/main:t", _NS)
    if inline_node is not None:
        return inline_node.text or ""
    if value_node is None:
        return ""
    value = value_node.text or ""
    if cell.attrib.get("t") == "s":
        return shared_strings[int(value)]
    return value


def _header_score(values: list[object]) -> int:
    headers = [str(value).strip() for value in values]
    nonblank = [header for header in headers if header]
    mapping = infer_columns(headers)
    metadata_count = sum(1 for header in headers if is_metadata_column(header))
    return len(mapping) * 4 + metadata_count * 3 + min(len(nonblank), 3)




def _dimension_header_indices(headers: list[str]) -> tuple[int, int, int] | None:
    for index, header in enumerate(headers):
        normalized = re.sub(r"[^a-z0-9]+", "", str(header or "").lower())
        if "dimension" not in normalized and "dims" not in normalized:
            continue
        if index + 2 >= len(headers):
            continue
        next_headers = [str(headers[index + offset] or "").strip() for offset in range(3)]
        if all(next_headers):
            continue
        return index, index + 1, index + 2
    return None


def _expand_merged_dimension_headers(headers: list[str], table: list[list[object]]) -> list[str]:
    expanded = list(headers)
    indices = _dimension_header_indices(expanded)
    if indices is None:
        return expanded
    length_index, width_index, height_index = indices
    base = expanded[length_index]
    if not base:
        return expanded
    sample_rows = table[1:6]
    has_three_numeric_columns = all(
        any(str(row[column] if column < len(row) else "").strip() for row in sample_rows)
        for column in indices
    )
    if not has_three_numeric_columns:
        return expanded
    expanded[length_index] = f"Length {base}"
    expanded[width_index] = f"Width {base}"
    expanded[height_index] = f"Height {base}"
    return expanded

def _headers_and_data_start(table: list[list[object]]) -> tuple[list[str], int]:
    first_headers = _expand_merged_dimension_headers(
        [str(value).strip() for value in table[0]],
        table,
    )
    first_score = _header_score(first_headers)
    if len(table) < 2:
        return first_headers, 1

    second_headers = _expand_merged_dimension_headers(
        [str(value).strip() for value in table[1]],
        table[1:],
    )
    second_score = _header_score(second_headers)
    second_mapping = infer_columns(second_headers)
    second_has_header_signal = (
        any(is_metadata_column(header) for header in second_headers)
        or bool({"sku", "order_id"} & set(second_mapping))
    )
    if second_has_header_signal and second_score >= 4 and second_score > first_score:
        return [
            second if second else first
            for first, second in zip(first_headers, second_headers, strict=True)
        ], 2
    return first_headers, 1


def _read_xlsx(path: str) -> list[SourceRows]:
    source_rows = []
    with zipfile.ZipFile(path) as archive:
        shared_strings = _xlsx_shared_strings(archive)
        for sheet_name, sheet_path in _xlsx_sheet_paths(archive):
            root = ElementTree.fromstring(archive.read(sheet_path))
            raw_rows = []
            for row in root.findall(".//main:row", _NS):
                values = {}
                for cell in row.findall("main:c", _NS):
                    values[_column_index(cell.attrib.get("r", ""))] = _xlsx_cell_value(
                        cell,
                        shared_strings,
                    )
                if values:
                    raw_rows.append(values)
            if not raw_rows:
                continue
            max_index = max(max(row) for row in raw_rows)
            table = [
                [row.get(index, "") for index in range(max_index + 1)]
                for row in raw_rows
            ]
            headers, data_start = _headers_and_data_start(table)
            rows = [
                {headers[index]: value for index, value in enumerate(values) if headers[index]}
                for values in table[data_start:]
                if any(str(value).strip() for value in values)
            ]
            if rows:
                source_rows.append(SourceRows(sheet_name=sheet_name, rows=rows))
    return source_rows


def read_workbook(path: str) -> list[SourceRows]:
    """Read all useful rows from a CSV or XLSX workbook."""
    suffix = Path(path).suffix.lower()
    if suffix == ".csv":
        return _read_csv(path)
    if suffix == ".xlsx":
        return _read_xlsx(path)
    raise ValueError(f"Unsupported intake file type: {path}")


def _read_sku_master_with_mappings(path: str) -> tuple[list[SKUItem], list[dict], int]:
    sku_items = []
    mappings = []
    rows_read = 0
    for source in read_workbook(path):
        if not source.rows:
            continue
        rows_read += len(source.rows)
        headers = list(source.rows[0].keys())
        mapping = infer_columns(headers)
        warnings = []
        has_separate_dimensions = "length" in mapping and "width" in mapping
        has_combined_dimensions = "dimensions" in mapping
        if "sku" not in mapping:
            warnings.append("No SKU column detected")
        if not has_separate_dimensions and not has_combined_dimensions:
            warnings.append("No separate or combined dimension columns detected")
        mappings.append(
            _mapping_record(
                path,
                source.sheet_name,
                "sku-master",
                dimension_column=mapping.get("dimensions")
                or " / ".join(
                    column
                    for column in [
                        mapping.get("length"),
                        mapping.get("width"),
                        mapping.get("height"),
                    ]
                    if column
                ),
                weight_column=mapping.get("weight"),
                warnings=warnings,
            )
        )
        if "sku" not in mapping or not (has_separate_dimensions or has_combined_dimensions):
            continue

        for row in source.rows:
            raw_sku = str(row.get(mapping["sku"], "")).strip()
            if not raw_sku:
                continue

            if has_combined_dimensions:
                dimension_unit = infer_dimension_unit(mapping["dimensions"])
                dimensions = _parse_dimensions(
                    row.get(mapping["dimensions"]),
                    default_unit=dimension_unit,
                )
                if dimensions is None:
                    continue
                is_flat = len(dimensions.original_dimensions) == 2
            else:
                height_header = mapping.get("height")
                length = _parse_number(row.get(mapping["length"]))
                width = _parse_number(row.get(mapping["width"]))
                height = _parse_number(row.get(height_header)) if height_header else None
                if length <= 0 or width <= 0 or (height_header and (height is None or height <= 0)):
                    continue
                dimension_unit = infer_dimension_unit(mapping["length"])
                dimensions = normalize_dimensions(length, width, height, unit=dimension_unit)
                is_flat = height is None

            weight_header = mapping.get("weight")
            weight_value = _parse_number(row.get(weight_header)) if weight_header else 0.0
            weight_unit = infer_weight_unit(weight_header)
            weight = normalize_weight(weight_value, weight_unit)
            product_name = str(row.get(mapping.get("product_name", ""), "")).strip() or raw_sku

            sku_items.append(
                SKUItem(
                    raw_sku=raw_sku,
                    canonical_sku=normalize_sku(raw_sku),
                    product_name=product_name,
                    length_cm=dimensions.dimensions.length,
                    width_cm=dimensions.dimensions.width,
                    height_cm=dimensions.dimensions.height,
                    weight_kg=weight.weight_kg,
                    is_flat=is_flat,
                    aliases=(),
                    metadata=_metadata(row, path, source.sheet_name),
                )
            )
    return sku_items, mappings, rows_read


def read_sku_master(path: str) -> list[SKUItem]:
    """Read SKU master rows from CSV or XLSX and preserve all metadata columns."""
    sku_items, _, _ = _read_sku_master_with_mappings(path)
    return sku_items


def _product_header_keys(value: object) -> set[str]:
    text = str(value or "").strip()
    return {
        key
        for key in {
            normalize_sku(text),
            re.sub(r"[\\W_]+", "", text.lower()),
        }
        if key
    }


def _wide_product_columns(
    rows: list[dict],
    headers: list[str],
    explicit_mapping: dict,
    allowed_product_keys: set[str] | None = None,
) -> list[str]:
    excluded = set(explicit_mapping.values())
    product_columns = []
    for header in headers:
        if header in excluded or is_metadata_column(header):
            continue
        if allowed_product_keys is not None and not (_product_header_keys(header) & allowed_product_keys):
            continue
        values = [row.get(header) for row in rows]
        nonblank = [value for value in values if str(value or "").strip()]
        if not nonblank:
            continue
        quantity_like = sum(1 for value in values if _looks_like_quantity(value))
        if quantity_like / len(values) >= 0.8:
            product_columns.append(header)
    return product_columns


def _has_order_identity(row: dict, mapping: dict, metadata_columns: list[str]) -> bool:
    order_header = mapping.get("order_id")
    if order_header:
        value = str(row.get(order_header, "") or "").strip()
        if value and value.lower() not in {"0", "0.0", "none", "nan", "total", "totals"}:
            return True
    identity_headers = [
        header
        for header in metadata_columns
        if is_metadata_column(header)
        and re.search(r"backer|order|vfi|name|email|phone|address|add1|shipping|country|city|postal|zip", header, re.I)
    ]
    meaningful = 0
    for header in identity_headers:
        value = str(row.get(header, "") or "").strip()
        if value and value.lower() not in {"0", "0.0", "none", "nan", "#ref!", "#value!"}:
            meaningful += 1
    return meaningful >= 2


def _read_orders_with_mappings(
    path: str,
    allowed_product_keys: set[str] | None = None,
) -> tuple[list[OrderLine], list[dict], int, int]:
    order_lines = []
    mappings = []
    rows_read = 0
    wide_product_column_count = 0
    for source in read_workbook(path):
        if not source.rows:
            continue
        rows_read += len(source.rows)
        headers = list(source.rows[0].keys())
        mapping = infer_columns(headers)
        warnings = []
        if "sku" in mapping:
            detected_format = "long-format"
            product_columns = []
            metadata_columns = [header for header in headers if header not in set(mapping.values())]
        else:
            detected_format = "wide-format"
            product_columns = _wide_product_columns(source.rows, headers, mapping, allowed_product_keys)
            wide_product_column_count += len(product_columns)
            metadata_columns = [header for header in headers if header not in product_columns]
            if not product_columns:
                warnings.append("No product quantity columns detected")

        mappings.append(
            _mapping_record(
                path,
                source.sheet_name,
                detected_format,
                metadata_columns=metadata_columns,
                product_quantity_columns=product_columns,
                warnings=warnings,
            )
        )

        for index, row in enumerate(source.rows, start=1):
            if "sku" in mapping:
                raw_sku = str(row.get(mapping["sku"], "")).strip()
                if not raw_sku:
                    continue
                order_lines.append(
                    OrderLine(
                        order_id=str(row.get(mapping.get("order_id", ""), f"{source.sheet_name}-{index}")).strip(),
                        raw_sku=raw_sku,
                        canonical_sku=normalize_sku(raw_sku),
                        quantity=_parse_quantity(row.get(mapping.get("quantity", ""), 1)),
                        region=str(row.get(mapping.get("region", ""), "") or "") or source.sheet_name,
                        country=str(row.get(mapping.get("country", ""), "") or "") or source.sheet_name,
                        state_province=str(row.get(mapping.get("state_province", ""), "") or "") or None,
                        metadata=_metadata(row, path, source.sheet_name),
                    )
                )
                continue

            if not _has_order_identity(row, mapping, metadata_columns):
                continue
            metadata = {
                header: row.get(header)
                for header in metadata_columns
                if str(row.get(header, "")).strip()
            }
            metadata["_source_file"] = str(path)
            metadata["_source_sheet"] = source.sheet_name
            order_id = str(
                row.get(mapping.get("order_id", ""), "") or f"{source.sheet_name}-{index}"
            ).strip()
            region = str(row.get(mapping.get("region", ""), "") or source.sheet_name) or None
            country = str(row.get(mapping.get("country", ""), "") or source.sheet_name) or None
            state_province = str(row.get(mapping.get("state_province", ""), "") or "") or None
            for product_column in product_columns:
                quantity = _quantity_value(row.get(product_column))
                if quantity is None or quantity <= 0:
                    continue
                order_lines.append(
                    OrderLine(
                        order_id=order_id,
                        raw_sku=product_column,
                        canonical_sku=normalize_sku(product_column),
                        quantity=quantity,
                        region=region,
                        country=country,
                        state_province=state_province,
                        metadata=dict(metadata),
                    )
                )
    return order_lines, mappings, rows_read, wide_product_column_count


def read_orders(path: str) -> list[OrderLine]:
    """Read order rows from CSV or XLSX and preserve all metadata columns."""
    order_lines, _, _, _ = _read_orders_with_mappings(path)
    return order_lines


def _match_key(value: object) -> str:
    text = str(value or "").strip().lower()
    text = text.replace("“", "\"").replace("”", "\"").replace("’", "'").replace("×", "x")
    return re.sub(r"[\W_]+", "", text)


def _soft_match_key(value: object, *, drop_pack: bool = False) -> str:
    text = str(value or "").strip().lower()
    text = text.replace("â€œ", "\"").replace("â€", "\"").replace("â€™", "'").replace("Ã—", "x")
    text = re.sub(r"\[[^\]]+\]", " ", text)
    text = re.sub(
        r"\((english|german|french|spanish|italian|polish|portuguese|russian|chinese|japanese|korean)\)",
        " ",
        text,
    )
    if drop_pack:
        text = re.sub(r"\bpack\b", " ", text)
    return re.sub(r"[\W_]+", "", text)


def _match_keys(value: object) -> list[str]:
    keys = [
        normalize_sku(value),
        _match_key(value),
        _soft_match_key(value),
        _soft_match_key(value, drop_pack=True),
    ]
    return [key for key in dict.fromkeys(keys) if key]


def match_orders_to_sku_master(
    order_lines: list[OrderLine],
    sku_items: list[SKUItem],
) -> tuple[list[OrderLine], list[UnmatchedSKURecord]]:
    """Match order lines to SKU master records while preserving unmatched SKUs."""
    known_skus = {}
    collisions = set()
    for item in sku_items:
        candidates = [item.canonical_sku, item.raw_sku, item.product_name, *item.aliases]
        for candidate in candidates:
            if candidate:
                for key in _match_keys(candidate):
                    if key in known_skus and known_skus[key] != item.canonical_sku:
                        collisions.add(key)
                        continue
                    known_skus[key] = item.canonical_sku
    for key in collisions:
        known_skus.pop(key, None)

    matched = []
    unmatched = []
    for line in order_lines:
        canonical = next((known_skus[key] for key in _match_keys(line.raw_sku) if key in known_skus), None)
        if canonical:
            matched.append(replace(line, canonical_sku=canonical))
        else:
            unmatched.append(
                UnmatchedSKURecord(
                    order_line=line,
                    reason="SKU not found in master data",
                    metadata=dict(line.metadata),
                )
            )
    return matched, unmatched


def read_intake(sku_master_path: str, orders_path: str) -> IntakeResult:
    """Read SKU and order files and return matched plus unmatched records."""
    sku_items, sku_mappings, sku_rows_read = _read_sku_master_with_mappings(sku_master_path)
    allowed_product_keys = {
        key
        for item in sku_items
        for candidate in [item.canonical_sku, item.raw_sku, item.product_name, *item.aliases]
        for key in _match_keys(candidate)
    }
    order_lines, order_mappings, order_rows_read, wide_product_column_count = _read_orders_with_mappings(
        orders_path,
        allowed_product_keys=allowed_product_keys,
    )
    matched, unmatched = match_orders_to_sku_master(order_lines, sku_items)
    product_columns = []
    for mapping in order_mappings:
        columns = mapping["detected product quantity columns"]
        if columns:
            product_columns.extend(columns.split(" | "))
    return IntakeResult(
        sku_items=sku_items,
        order_lines=order_lines,
        matched_order_lines=matched,
        unmatched_skus=unmatched,
        column_mappings=[*sku_mappings, *order_mappings],
        debug={
            "sku_rows_read": sku_rows_read,
            "sku_items_parsed": len(sku_items),
            "order_rows_read": order_rows_read,
            "wide_product_columns_detected": wide_product_column_count,
            "detected_product_quantity_columns": product_columns,
            "order_lines_created": len(order_lines),
            "matched": len(matched),
            "unmatched": len(unmatched),
        },
    )

