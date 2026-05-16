import csv
import zipfile
from pathlib import Path

from box_optimizer import optimize_workbook
from box_optimizer.io.excel_reader import read_workbook


def _write_csv(path: Path, rows: list[dict]) -> None:
    with open(path, "w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def _inline_cell(column, row, value):
    return (
        f'<c r="{column}{row}" t="inlineStr">'
        f"<is><t>{value}</t></is>"
        "</c>"
    )


def _write_xlsx(path: Path, sheet_name: str, rows: list[dict]) -> None:
    headers = list(rows[0].keys())
    table = [headers, *[[row.get(header, "") for header in headers] for row in rows]]
    xml_rows = []
    for row_number, row_values in enumerate(table, start=1):
        cells = [
            _inline_cell(chr(ord("A") + column), row_number, value)
            for column, value in enumerate(row_values)
        ]
        xml_rows.append(f'<row r="{row_number}">{"".join(cells)}</row>')
    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<sheetData>{"".join(xml_rows)}</sheetData>'
        "</worksheet>"
    )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheets><sheet name="{sheet_name}" sheetId="1" r:id="rId1"/></sheets>'
        "</workbook>"
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"/>'
        "</Relationships>"
    )
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("xl/workbook.xml", workbook)
        archive.writestr("xl/_rels/workbook.xml.rels", rels)
        archive.writestr("xl/worksheets/sheet1.xml", sheet_xml)


def test_optimize_workbook_public_api_writes_output_and_returns_summary(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized_shipping_plan.xlsx"
    _write_csv(
        sku_master_path,
        [
            {
                "SKU": "Core Game",
                "Product Name": "Core Game",
                "Length": "5",
                "Width": "5",
                "Height": "5",
                "Weight kg": "1",
            }
        ],
    )
    _write_csv(
        orders_path,
        [
            {
                "Order ID": "1001",
                "SKU": "Core Game",
                "Quantity": "1",
                "Region": "NA",
                "Country": "US",
                "State": "California",
                "Pledge Level": "Deluxe",
            },
            {
                "Order ID": "1002",
                "SKU": "Missing SKU",
                "Quantity": "1",
                "Region": "EU",
                "Country": "DE",
                "State": "",
                "Pledge Level": "Standard",
            },
        ],
    )

    result = optimize_workbook(
        sku_master_path=str(sku_master_path),
        orders_path=str(orders_path),
        output_path=str(output_path),
        config={
            "max_carton_cm": [74, 37, 44],
            "dimensional_divisor": 5000,
            "packing_weight_uplift": 1.15,
            "standardization_tolerance_cm": 2,
            "preserve_region_sheets": True,
        },
    )

    assert result["output_path"] == str(output_path)
    assert result["orders_processed"] == 1
    assert result["boxes_created"] == 1
    assert result["box_types"] == 1
    assert result["unmatched_skus"] == 1
    assert result["warnings"] == ["1 unmatched SKU rows were preserved."]
    assert output_path.exists()

    workbook_rows = read_workbook(str(output_path))
    assert [sheet.sheet_name for sheet in workbook_rows[:3]] == [
        "Summary",
        "Order Volume Weights",
        "Box Size Summary",
    ]
    order_volume_rows = workbook_rows[1].rows
    assert order_volume_rows[0]["Order ID"] == "1001"
    assert order_volume_rows[0]["US State Abbreviation"] == "CA"
    assert order_volume_rows[0]["Pledge Level"] == "Deluxe"
    assert order_volume_rows[0]["SKU Breakdown"] == "CORE GAME x1"


def test_optimize_workbook_returns_config_warnings_for_unsupported_overrides(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized_shipping_plan.xlsx"
    _write_csv(
        sku_master_path,
        [
            {
                "SKU": "A",
                "Product Name": "A",
                "Length": "5",
                "Width": "5",
                "Height": "5",
                "Weight kg": "1",
            }
        ],
    )
    _write_csv(
        orders_path,
        [{"Order ID": "1", "SKU": "A", "Quantity": "1"}],
    )

    result = optimize_workbook(
        sku_master_path=str(sku_master_path),
        orders_path=str(orders_path),
        output_path=str(output_path),
        config={"max_carton_cm": [10, 10, 10]},
    )

    assert "Custom max_carton_cm is not yet supported; using 74 x 37 x 44 cm." in result["warnings"]


def test_end_to_end_sample_xlsx_wide_orders_produces_output_rows(tmp_path):
    sku_master_path = tmp_path / "sample_sku_master.xlsx"
    orders_path = tmp_path / "sample_orders.xlsx"
    output_path = tmp_path / "optimized_shipping_plan.xlsx"
    _write_xlsx(
        sku_master_path,
        "SKU Master",
        [
            {
                "SKU": "CG-001",
                "Product Name": "Core Game",
                "Dimensions": "10 x 5 x 3 cm",
                "Weight kg": "1",
            },
            {
                "SKU": "EXP-A",
                "Product Name": "Expansion A",
                "Dimensions": "8 x 4 x 2 cm",
                "Weight kg": "0.5",
            },
        ],
    )
    _write_xlsx(
        orders_path,
        "US",
        [
            {
                "Order ID": "1001",
                "Country": "US",
                "Core Game": "1",
                "Expansion A": "2",
            }
        ],
    )

    result = optimize_workbook(
        sku_master_path=str(sku_master_path),
        orders_path=str(orders_path),
        output_path=str(output_path),
        config={"preserve_region_sheets": False},
    )

    assert result["orders_processed"] > 0
    workbook_rows = read_workbook(str(output_path))
    order_volume_rows = next(
        sheet.rows for sheet in workbook_rows if sheet.sheet_name == "Order Volume Weights"
    )
    assert len(order_volume_rows) > 0
