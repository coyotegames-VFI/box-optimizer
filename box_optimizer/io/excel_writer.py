"""Excel output helpers."""

from pathlib import Path
from xml.sax.saxutils import escape
import zipfile


REQUIRED_SHEETS = [
    "Summary",
    "Order Volume Weights",
    "Optimized to Pack",
    "Box Size Summary",
]

OPTIONAL_SHEETS = [
    "Unmatched SKUs",
    "Multi Box Detail",
    "Pledge Combination Summary",
    "Packing Detail",
    "Input Column Mapping",
    "Errors and Warnings",
]

ORDER_VOLUME_WEIGHTS_COLUMNS = [
    "Region",
    "Order ID",
    "Country",
    "State/Province",
    "US State Abbreviation",
    "Packed Actual Weight kg",
    "Dimensional Weight kg (/5000)",
    "Chargeable Weight kg",
    "Chargeable Weight g",
    "Total Units",
    "Box Qty",
    "Box Plan",
    "Per-Box Chargeable Weight",
    "SKU Breakdown",
]

_UNIT_HINTS = {
    "length": "cm",
    "width": "cm",
    "height": "cm",
    "weight": "kg",
    "volume": "cm3",
}


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


def _cell_xml(row_index: int, column_index: int, value: object, style: int | None = None) -> str:
    reference = f"{_column_letter(column_index)}{row_index}"
    if isinstance(value, bool):
        return f'<c r="{reference}" t="b"><v>{1 if value else 0}</v></c>'
    if isinstance(value, int | float) and not isinstance(value, bool):
        return _number_cell(reference, value, style)
    return _string_cell(reference, value, style)


def _headers_for_sheet(sheet_name: str, rows: list[dict]) -> list[str]:
    headers = []
    seen = set()

    if sheet_name == "Order Volume Weights":
        headers.extend(ORDER_VOLUME_WEIGHTS_COLUMNS)
        seen.update(ORDER_VOLUME_WEIGHTS_COLUMNS)

    for row in rows:
        for key in row:
            if key not in seen:
                seen.add(key)
                headers.append(key)

    return headers


def _rows_to_table(sheet_name: str, rows: list[dict]) -> tuple[list[str], list[list[object]]]:
    if not rows:
        if sheet_name == "Order Volume Weights":
            return ORDER_VOLUME_WEIGHTS_COLUMNS, []
        return ["Note"], [["No records"]]

    headers = _headers_for_sheet(sheet_name, rows)
    display_headers = [_header_with_units(header) for header in headers]
    values = [[row.get(header, "") for header in headers] for row in rows]
    return display_headers, values


def _column_widths(headers: list[str], rows: list[list[object]]) -> list[float]:
    widths = []
    for index, header in enumerate(headers):
        longest = len(str(header))
        for row in rows:
            if index < len(row):
                longest = max(longest, len(str(row[index])))
        widths.append(min(max(longest + 2, 12), 40))
    return widths


def _worksheet_xml(sheet_name: str, rows: list[dict]) -> str:
    headers, values = _rows_to_table(sheet_name, rows)
    table = [headers, *values]
    row_xml = []
    for row_index, row in enumerate(table, start=1):
        cells = [
            _cell_xml(row_index, column_index, value, style=1 if row_index == 1 else None)
            for column_index, value in enumerate(row)
        ]
        row_xml.append(f'<row r="{row_index}">{"".join(cells)}</row>')

    widths = _column_widths(headers, values)
    cols_xml = "".join(
        f'<col min="{index + 1}" max="{index + 1}" width="{width}" customWidth="1"/>'
        for index, width in enumerate(widths)
    )
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


def _workbook_xml(sheet_names: list[str]) -> str:
    sheets = "".join(
        f'<sheet name="{escape(name)}" sheetId="{index}" r:id="rId{index}"/>'
        for index, name in enumerate(sheet_names, start=1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<sheets>{sheets}</sheets>"
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


def _content_types(sheet_count: int) -> str:
    overrides = "".join(
        f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for index in range(1, sheet_count + 1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        f"{overrides}"
        "</Types>"
    )


def _styles_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2">'
        '<font><sz val="11"/><name val="Calibri"/></font>'
        '<font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font>'
        "</fonts>"
        '<fills count="3">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>'
        "</fills>"
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="2">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
        "</cellXfs>"
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        "</styleSheet>"
    )


def _build_sheet_payloads(
    rows: list[dict] | None,
    sheets: dict[str, list[dict]] | None,
    **named_rows: list[dict],
) -> list[tuple[str, list[dict]]]:
    payloads = {name: [] for name in REQUIRED_SHEETS}
    if rows is not None:
        payloads["Summary"] = rows
    if sheets:
        payloads.update(sheets)

    aliases = {
        "summary_rows": "Summary",
        "order_volume_weights_rows": "Order Volume Weights",
        "optimized_to_pack_rows": "Optimized to Pack",
        "box_size_summary_rows": "Box Size Summary",
        "unmatched_skus_rows": "Unmatched SKUs",
        "packing_detail_rows": "Packing Detail",
        "multi_box_detail_rows": "Multi Box Detail",
        "pledge_combination_summary_rows": "Pledge Combination Summary",
        "input_column_mapping_rows": "Input Column Mapping",
        "errors_and_warnings_rows": "Errors and Warnings",
    }
    for key, sheet_name in aliases.items():
        if key in named_rows and named_rows[key] is not None:
            payloads[sheet_name] = named_rows[key]

    ordered = [(name, payloads.get(name, [])) for name in REQUIRED_SHEETS]
    for name in OPTIONAL_SHEETS:
        if payloads.get(name):
            ordered.append((name, payloads[name]))

    for name, sheet_rows in payloads.items():
        if name not in REQUIRED_SHEETS and name not in OPTIONAL_SHEETS and sheet_rows:
            ordered.append((name, sheet_rows))

    return ordered


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
    safe_names = [_safe_sheet_name(name) for name, _ in sheet_payloads]

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", _content_types(len(sheet_payloads)))
        archive.writestr("_rels/.rels", _root_rels())
        archive.writestr("xl/workbook.xml", _workbook_xml(safe_names))
        archive.writestr("xl/_rels/workbook.xml.rels", _workbook_rels(len(sheet_payloads)))
        archive.writestr("xl/styles.xml", _styles_xml())
        for index, (sheet_name, sheet_rows) in enumerate(sheet_payloads, start=1):
            archive.writestr(
                f"xl/worksheets/sheet{index}.xml",
                _worksheet_xml(sheet_name, sheet_rows),
            )

    return str(output_path)
