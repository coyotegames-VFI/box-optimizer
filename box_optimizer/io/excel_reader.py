"""Excel and CSV input helpers."""

import csv
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree

from box_optimizer.io.column_mapper import (
    infer_columns,
    infer_dimension_unit,
    infer_weight_unit,
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


_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def _parse_number(value: object, default: float = 0.0) -> float:
    if value is None or value == "":
        return default
    match = re.search(r"-?\d+(?:\.\d+)?", str(value).replace(",", ""))
    if not match:
        return default
    return float(match.group(0))


def _parse_quantity(value: object) -> int:
    quantity = int(_parse_number(value, default=1))
    return max(quantity, 1)


def _metadata(row: dict, source_path: str, sheet_name: str) -> dict:
    metadata = dict(row)
    metadata["_source_file"] = str(source_path)
    metadata["_source_sheet"] = sheet_name
    return metadata


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
            headers = [str(value).strip() for value in table[0]]
            rows = [
                {headers[index]: value for index, value in enumerate(values) if headers[index]}
                for values in table[1:]
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


def read_sku_master(path: str) -> list[SKUItem]:
    """Read SKU master rows from CSV or XLSX and preserve all metadata columns."""
    sku_items = []
    for source in read_workbook(path):
        if not source.rows:
            continue
        mapping = infer_columns(list(source.rows[0].keys()))
        if "sku" not in mapping or "length" not in mapping or "width" not in mapping:
            continue

        for row in source.rows:
            raw_sku = str(row.get(mapping["sku"], "")).strip()
            if not raw_sku:
                continue

            height_header = mapping.get("height")
            length = _parse_number(row.get(mapping["length"]))
            width = _parse_number(row.get(mapping["width"]))
            height = _parse_number(row.get(height_header)) if height_header else None
            dimension_unit = infer_dimension_unit(mapping["length"])
            dimensions = normalize_dimensions(length, width, height, unit=dimension_unit)

            weight_header = mapping.get("weight")
            weight_value = _parse_number(row.get(weight_header)) if weight_header else 0.0
            weight_unit = infer_weight_unit(weight_header)
            weight = normalize_weight(weight_value, weight_unit)
            product_name = str(row.get(mapping.get("product_name", ""), "")).strip()

            sku_items.append(
                SKUItem(
                    raw_sku=raw_sku,
                    canonical_sku=normalize_sku(raw_sku),
                    product_name=product_name,
                    length_cm=dimensions.dimensions.length,
                    width_cm=dimensions.dimensions.width,
                    height_cm=dimensions.dimensions.height,
                    weight_kg=weight.weight_kg,
                    is_flat=height is None,
                    aliases=(),
                    metadata=_metadata(row, path, source.sheet_name),
                )
            )
    return sku_items


def read_orders(path: str) -> list[OrderLine]:
    """Read order rows from CSV or XLSX and preserve all metadata columns."""
    order_lines = []
    for source in read_workbook(path):
        if not source.rows:
            continue
        mapping = infer_columns(list(source.rows[0].keys()))
        if "sku" not in mapping:
            continue

        for index, row in enumerate(source.rows, start=1):
            raw_sku = str(row.get(mapping["sku"], "")).strip()
            if not raw_sku:
                continue
            order_lines.append(
                OrderLine(
                    order_id=str(row.get(mapping.get("order_id", ""), f"row-{index}")).strip(),
                    raw_sku=raw_sku,
                    canonical_sku=normalize_sku(raw_sku),
                    quantity=_parse_quantity(row.get(mapping.get("quantity", ""), 1)),
                    region=str(row.get(mapping.get("region", ""), "") or "") or None,
                    country=str(row.get(mapping.get("country", ""), "") or "") or None,
                    state_province=str(row.get(mapping.get("state_province", ""), "") or "") or None,
                    metadata=_metadata(row, path, source.sheet_name),
                )
            )
    return order_lines


def match_orders_to_sku_master(
    order_lines: list[OrderLine],
    sku_items: list[SKUItem],
) -> tuple[list[OrderLine], list[UnmatchedSKURecord]]:
    """Match order lines to SKU master records while preserving unmatched SKUs."""
    known_skus = {item.canonical_sku for item in sku_items}
    for item in sku_items:
        known_skus.update(normalize_sku(alias) for alias in item.aliases)

    matched = []
    unmatched = []
    for line in order_lines:
        if line.canonical_sku in known_skus:
            matched.append(line)
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
    sku_items = read_sku_master(sku_master_path)
    order_lines = read_orders(orders_path)
    matched, unmatched = match_orders_to_sku_master(order_lines, sku_items)
    return IntakeResult(
        sku_items=sku_items,
        order_lines=order_lines,
        matched_order_lines=matched,
        unmatched_skus=unmatched,
    )
