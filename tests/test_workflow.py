import csv
import zipfile
from pathlib import Path

from box_optimizer import optimize_workbook
from box_optimizer.io.excel_reader import read_workbook
from box_optimizer.weights import packed_actual_weight_kg
from box_optimizer.workflow import format_kg_display, inspect_workbook


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


def _write_xlsx_table(path: Path, sheet_name: str, table: list[list[object]]) -> None:
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


def _sheet_rows(path: Path, sheet_name: str) -> list[dict]:
    return next(sheet.rows for sheet in read_workbook(str(path)) if sheet.sheet_name == sheet_name)


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


def test_order_exceeding_single_carton_cap_is_split_not_fatal(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized_shipping_plan.xlsx"
    _write_csv(
        sku_master_path,
        [
            {
                "SKU": "Big Box",
                "Product Name": "Big Box",
                "Length": "42",
                "Width": "31",
                "Height": "27",
                "Weight kg": "2",
            }
        ],
    )
    _write_csv(
        orders_path,
        [{"Order ID": "SPLIT-1", "SKU": "Big Box", "Quantity": "3"}],
    )

    result = optimize_workbook(
        sku_master_path=str(sku_master_path),
        orders_path=str(orders_path),
        output_path=str(output_path),
        config={"packing_mode": "fast", "preserve_region_sheets": False},
    )

    assert result["orders_processed"] == 1
    assert result["boxes_created"] > 1
    workbook_rows = read_workbook(str(output_path))
    order_volume_rows = next(
        sheet.rows for sheet in workbook_rows if sheet.sheet_name == "Order Volume Weights"
    )
    assert int(float(order_volume_rows[0]["Box Qty"])) > 1
    multi_box_rows = next(
        sheet.rows for sheet in workbook_rows if sheet.sheet_name == "Multi Box Detail"
    )
    assert int(float(multi_box_rows[0]["Box Qty"])) > 1


def test_single_item_larger_than_cap_is_warned_not_fatal(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized_shipping_plan.xlsx"
    _write_csv(
        sku_master_path,
        [
            {
                "SKU": "Good",
                "Product Name": "Good",
                "Length": "5",
                "Width": "5",
                "Height": "5",
                "Weight kg": "1",
            },
            {
                "SKU": "Huge",
                "Product Name": "Huge",
                "Length": "80",
                "Width": "10",
                "Height": "10",
                "Weight kg": "3",
            },
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "OK-1", "SKU": "Good", "Quantity": "1"},
            {"Order ID": "BAD-1", "SKU": "Huge", "Quantity": "1"},
        ],
    )

    result = optimize_workbook(
        sku_master_path=str(sku_master_path),
        orders_path=str(orders_path),
        output_path=str(output_path),
        config={"packing_mode": "fast", "preserve_region_sheets": False},
    )

    assert result["orders_processed"] == 1
    assert result["warning_count"] > 0
    workbook_rows = read_workbook(str(output_path))
    order_volume_rows = next(
        sheet.rows for sheet in workbook_rows if sheet.sheet_name == "Order Volume Weights"
    )
    assert [row["Order ID"] for row in order_volume_rows] == ["OK-1"]
    warning_rows = next(
        sheet.rows for sheet in workbook_rows if sheet.sheet_name == "Errors and Warnings"
    )
    assert any(row["Order ID"] == "BAD-1" and row["Error Type"] == "OversizedItem" for row in warning_rows)


def test_optimized_carton_dimensions_remain_capped_but_vendor_assignment_can_exceed_old_cap(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized_shipping_plan.xlsx"
    _write_csv(
        sku_master_path,
        [
            {
                "SKU": "Near Cap",
                "Product Name": "Near Cap",
                "Length": "70",
                "Width": "33",
                "Height": "40",
                "Weight kg": "2",
            }
        ],
    )
    _write_csv(
        orders_path,
        [{"Order ID": "CAP-1", "SKU": "Near Cap", "Quantity": "1"}],
    )

    optimize_workbook(
        sku_master_path=str(sku_master_path),
        orders_path=str(orders_path),
        output_path=str(output_path),
        config={"packing_mode": "fast", "preserve_region_sheets": False},
    )

    workbook_rows = read_workbook(str(output_path))
    order_volume_rows = next(
        sheet.rows for sheet in workbook_rows if sheet.sheet_name == "Order Volume Weights"
    )
    for row in order_volume_rows:
        assert float(row["Assigned Box Length cm"]) >= 74
        assert row["Vendor Box ID"]
    warning_rows = next(
        sheet.rows for sheet in workbook_rows if sheet.sheet_name == "Errors and Warnings"
    )
    warning_keys = [
        (
            row["Order ID"],
            row["SKU"],
            row["Stage"],
            row["Error Type"],
            row["Message"],
        )
        for row in warning_rows
    ]
    assert len(warning_keys) == len(set(warning_keys))


def test_inspect_reports_matched_and_unmatched_rule_keys(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    _write_csv(
        sku_master_path,
        [
            {
                "SKU": "Core Game",
                "Product Name": "Earth Under Siege Core Box [72104]",
                "Length": "31",
                "Width": "22",
                "Height": "8",
                "Weight kg": "1",
            }
        ],
    )
    _write_csv(
        orders_path,
        [{"Order ID": "1", "SKU": "Core Game", "Quantity": "1"}],
    )

    result = inspect_workbook(
        str(sku_master_path),
        str(orders_path),
        config={
            "sku_rules": {
                "Earth Under Siege Core Box [72104]": {"prepacked": True},
                "Missing Rule": {"no_padding": True},
            }
        },
    )

    assert result["matched_rule_keys"] == ["Earth Under Siege Core Box [72104]"]
    assert result["unmatched_rule_keys"] == ["Missing Rule"]


def test_no_padding_sku_skips_normal_padding(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [{"SKU": "Flat", "Product Name": "Flat", "Length": "10", "Width": "5", "Height": "3", "Weight kg": "1"}],
    )
    _write_csv(orders_path, [{"Order ID": "1", "SKU": "Flat", "Quantity": "1"}])

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={
            "sku_rules": {"Flat": {"no_padding": True}},
            "preserve_region_sheets": False,
            "use_vendor_box_menu": False,
        },
    )

    row = _sheet_rows(output_path, "Multi Box Detail")[0]
    assert float(row["Length cm"]) == 12
    assert float(row["Width cm"]) == 7
    assert float(row["Height cm"]) == 5


def test_prepacked_sku_uses_existing_dimensions(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [{"SKU": "Core", "Product Name": "Core", "Length": "31", "Width": "22", "Height": "8", "Weight kg": "1"}],
    )
    _write_csv(orders_path, [{"Order ID": "1", "SKU": "Core", "Quantity": "1"}])

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={
            "sku_rules": {"Core": {"prepacked": True}},
            "preserve_region_sheets": False,
            "use_vendor_box_menu": False,
        },
    )

    row = _sheet_rows(output_path, "Multi Box Detail")[0]
    assert float(row["Length cm"]) == 31
    assert float(row["Width cm"]) == 22
    assert float(row["Height cm"]) == 8


def test_forced_box_cm_sets_assigned_box_dimensions(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [{"SKU": "Core", "Product Name": "Core", "Length": "30", "Width": "20", "Height": "7", "Weight kg": "1"}],
    )
    _write_csv(orders_path, [{"Order ID": "1", "SKU": "Core", "Quantity": "1"}])

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={
            "sku_rules": {
                "Core": {
                    "prepacked": True,
                    "forced_box_cm": [31, 22, 8],
                    "box_type": "PREPACK-CORE-BOX",
                }
            },
            "preserve_region_sheets": False,
        },
    )

    row = _sheet_rows(output_path, "Order Volume Weights")[0]
    assert row["Box Type"] == "PREPACK-CORE-BOX"
    assert float(row["Assigned Box Length cm"]) == 31
    assert float(row["Assigned Box Width cm"]) == 22
    assert float(row["Assigned Box Height cm"]) == 8


def test_ships_alone_sku_is_separated_from_addons(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {"SKU": "Core", "Product Name": "Core", "Length": "31", "Width": "22", "Height": "8", "Weight kg": "1"},
            {"SKU": "Addon", "Product Name": "Addon", "Length": "8", "Width": "5", "Height": "2", "Weight kg": "0.2"},
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "1", "SKU": "Core", "Quantity": "1"},
            {"Order ID": "1", "SKU": "Addon", "Quantity": "1"},
        ],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={
            "packing_mode": "fast",
            "sku_rules": {
                "Core": {
                    "prepacked": True,
                    "ships_alone": True,
                    "box_type": "PREPACK-CORE-BOX",
                }
            },
            "preserve_region_sheets": False,
        },
    )

    order_rows = _sheet_rows(output_path, "Order Volume Weights")
    assert int(float(order_rows[0]["Box Qty"])) == 2
    multi_rows = _sheet_rows(output_path, "Multi Box Detail")
    assert len(multi_rows) == 2
    assert any(row["Box Type"] == "PREPACK-CORE-BOX" for row in multi_rows)


def test_rule_assigned_box_type_does_not_show_vendor_assignment(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {"SKU": "Ready", "Product Name": "Ready", "Length": "64", "Width": "32", "Height": "32", "Weight kg": "1"},
            {"SKU": "Addon", "Product Name": "Addon", "Length": "8", "Width": "5", "Height": "2", "Weight kg": "0.2"},
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "1", "SKU": "Ready", "Quantity": "1"},
            {"Order ID": "1", "SKU": "Addon", "Quantity": "1"},
        ],
    )

    result = optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={
            "packing_mode": "fast",
            "use_vendor_box_menu": True,
            "sku_rules": {
                "Ready": {
                    "prepacked": True,
                    "no_padding": True,
                    "ships_alone": True,
                    "box_type": "READY-SKU-CARTON",
                }
            },
            "preserve_region_sheets": False,
        },
    )

    assert result["box_types"] == 2
    ready_row = next(row for row in _sheet_rows(output_path, "Multi Box Detail") if row["Box Type"] == "READY-SKU-CARTON")
    assert ready_row["Vendor Box ID"] == ""
    assert ready_row["Box Selection Decision"] == "rule_assigned_box"
    assert float(ready_row["Length cm"]) == 64
    assert float(ready_row["Width cm"]) == 32
    assert float(ready_row["Height cm"]) == 32


def test_forced_box_validation_warnings_do_not_crash(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {"SKU": "TooBigBox", "Product Name": "TooBigBox", "Length": "5", "Width": "5", "Height": "5", "Weight kg": "1"},
            {"SKU": "TooSmallBox", "Product Name": "TooSmallBox", "Length": "30", "Width": "20", "Height": "7", "Weight kg": "1"},
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "1", "SKU": "TooBigBox", "Quantity": "1"},
            {"Order ID": "2", "SKU": "TooSmallBox", "Quantity": "1"},
        ],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={
            "sku_rules": {
                "TooBigBox": {"forced_box_cm": [80, 20, 20]},
                "TooSmallBox": {"prepacked": True, "forced_box_cm": [10, 10, 5]},
            },
            "preserve_region_sheets": False,
        },
    )

    warning_rows = _sheet_rows(output_path, "Errors and Warnings")
    assert any(row["Error Type"] == "ForcedBoxExceedsCap" for row in warning_rows)
    assert any(row["Error Type"] == "ForcedBoxTooSmall" for row in warning_rows)


def test_default_order_summary_has_one_order_row_and_box_detail_rows(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {
                "SKU": "Big Box",
                "Product Name": "Big Box",
                "Length": "42",
                "Width": "31",
                "Height": "27",
                "Weight kg": "2",
            }
        ],
    )
    _write_csv(
        orders_path,
        [{"Order ID": f"ORDER-{index}", "SKU": "Big Box", "Quantity": "3"} for index in range(1, 6)],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={"packing_mode": "fast", "max_orders": 5, "preserve_region_sheets": False},
    )

    order_rows = _sheet_rows(output_path, "Order Volume Weights")
    multi_rows = _sheet_rows(output_path, "Multi Box Detail")
    assert len(order_rows) == 5
    assert len(multi_rows) == sum(int(float(row["Box Qty"])) for row in order_rows)
    for row in order_rows:
        matching_detail = [detail for detail in multi_rows if detail["Order ID"] == row["Order ID"]]
        assert int(float(row["Box Qty"])) == len(matching_detail)
        assert row["Box Type"] == "MULTI-BOX"
        assert row["Box Standardization Note"] == "Multi-box order; see Multi Box Detail"
        assert "Box 1:" in row["Box Plan"]


def test_box_size_summary_includes_usage_counts(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [{"SKU": "A", "Product Name": "A", "Length": "5", "Width": "5", "Height": "5", "Weight kg": "1"}],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "1", "SKU": "A", "Quantity": "1"},
            {"Order ID": "2", "SKU": "A", "Quantity": "1"},
        ],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={"preserve_region_sheets": False},
    )

    rows = _sheet_rows(output_path, "Box Size Summary")
    assert "Box Count" in rows[0]
    assert "Order Count" in rows[0]
    assert int(float(rows[0]["Box Count"])) == 2
    assert int(float(rows[0]["Order Count"])) == 2


def test_pledge_combination_summary_groups_identical_breakdowns(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {"SKU": "Core", "Product Name": "Core", "Length": "5", "Width": "5", "Height": "5", "Weight kg": "1"},
            {"SKU": "Addon", "Product Name": "Addon", "Length": "4", "Width": "4", "Height": "2", "Weight kg": "0.2"},
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "1", "SKU": "Core", "Quantity": "1"},
            {"Order ID": "1", "SKU": "Addon", "Quantity": "1"},
            {"Order ID": "2", "SKU": "Core", "Quantity": "1"},
            {"Order ID": "2", "SKU": "Addon", "Quantity": "1"},
        ],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={"preserve_region_sheets": False},
    )

    rows = _sheet_rows(output_path, "Pledge Combination Summary")
    assert len(rows) == 1
    assert int(float(rows[0]["Order Count"])) == 2
    assert rows[0]["SKU Breakdown"] == "ADDON x1 | CORE x1"


def test_box_detail_output_granularity_repeats_order_per_box_for_debug(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {
                "SKU": "Big Box",
                "Product Name": "Big Box",
                "Length": "42",
                "Width": "31",
                "Height": "27",
                "Weight kg": "2",
            }
        ],
    )
    _write_csv(orders_path, [{"Order ID": "SPLIT", "SKU": "Big Box", "Quantity": "3"}])

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={
            "packing_mode": "fast",
            "output_granularity": "box_detail",
            "preserve_region_sheets": False,
        },
    )

    order_rows = _sheet_rows(output_path, "Order Volume Weights")
    assert len(order_rows) == 3
    assert {row["Order ID"] for row in order_rows} == {"SPLIT"}


def test_display_carton_dimensions_are_whole_centimeters_rounded_up(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {"SKU": "A", "Product Name": "A", "Length": "30", "Width": "19", "Height": "9", "Weight kg": "1"},
            {"SKU": "B", "Product Name": "B", "Length": "34", "Width": "19", "Height": "9", "Weight kg": "1"},
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "1", "SKU": "A", "Quantity": "1"},
            {"Order ID": "2", "SKU": "B", "Quantity": "1"},
        ],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={
            "sku_rules": {
                "A": {"prepacked": True, "forced_box_cm": [35.0, 20, 10]},
                "B": {"prepacked": True, "forced_box_cm": [35.6, 20, 10]},
            },
            "preserve_region_sheets": False,
        },
    )

    rows = _sheet_rows(output_path, "Order Volume Weights")
    by_order = {row["Order ID"]: row for row in rows}
    assert float(by_order["1"]["Assigned Box Length cm"]) == 35
    assert float(by_order["2"]["Assigned Box Length cm"]) == 36
    for sheet_name in ["Order Volume Weights", "Multi Box Detail", "Box Size Summary", "Pledge Combination Summary"]:
        for row in _sheet_rows(output_path, sheet_name):
            for column in [key for key in row if key.endswith("Length cm") or key.endswith("Width cm") or key.endswith("Height cm")]:
                assert float(row[column]).is_integer()
                assert "." not in row[column]


def test_weight_display_truncates_to_one_decimal_without_changing_internal_math(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [{"SKU": "Heavy", "Product Name": "Heavy", "Length": "5", "Width": "5", "Height": "5", "Weight kg": "5.21634"}],
    )
    _write_csv(orders_path, [{"Order ID": "1", "SKU": "Heavy", "Quantity": "1"}])

    internal = packed_actual_weight_kg(5.21634)
    assert internal > 5.99
    assert format_kg_display(internal) == 5.9

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={"preserve_region_sheets": False},
    )

    row = _sheet_rows(output_path, "Order Volume Weights")[0]
    assert float(row["Packed Actual Weight kg"]) == 5.9
    assert len(str(row["Packed Actual Weight kg"]).split(".")[-1]) == 1


def test_order_metadata_columns_survive_at_end_of_order_volume_weights(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [{"SKU": "Core", "Product Name": "Core", "Length": "5", "Width": "5", "Height": "5", "Weight kg": "1"}],
    )
    _write_csv(
        orders_path,
        [
            {
                "Backer Number": "BK-7",
                "Order ID": "1",
                "SKU": "Core",
                "Quantity": "1",
                "Name": "Ada Lovelace",
                "Email": "ada@example.com",
                "Address 1": "1 Algorithm Ave",
                "Shipping Method": "Courier",
                "Notes": "Leave with desk",
            }
        ],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={"preserve_region_sheets": False},
    )

    row = _sheet_rows(output_path, "Order Volume Weights")[0]
    headers = list(row)
    assert row["Backer Number"] == "BK-7"
    assert row["Name"] == "Ada Lovelace"
    assert row["Email"] == "ada@example.com"
    assert row["Address 1"] == "1 Algorithm Ave"
    assert row["Shipping Method"] == "Courier"
    assert row["Notes"] == "Leave with desk"
    assert headers.index("Backer Number") > headers.index("Warning Summary")


def test_clean_default_order_volume_weights_hides_audit_columns(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [{"SKU": "Core", "Product Name": "Core", "Length": "5", "Width": "5", "Height": "5", "Weight kg": "1"}],
    )
    _write_csv(orders_path, [{"Order ID": "1", "SKU": "Core", "Quantity": "1"}])

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={"preserve_region_sheets": False},
    )

    row = _sheet_rows(output_path, "Order Volume Weights")[0]
    for hidden_column in [
        "Actual Item Weight lb",
        "Packed Actual Weight lb (+15%)",
        "Bundled/Padded Volume cmÂ³",
        "Dimensional Weight lb",
        "Chargeable Weight lb",
        "Length cm",
        "Width cm",
        "Height cm",
        "Optimized Length cm",
        "Optimized Width cm",
        "Optimized Height cm",
    ]:
        assert hidden_column not in row
    assert "Assigned Box Length cm" in row
    assert "Packed Actual Weight kg" in row
    assert "Chargeable Weight g" in row


def test_wide_format_metadata_columns_detected_and_preserved(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [{"SKU": "Core", "Product Name": "Core", "Length": "5", "Width": "5", "Height": "5", "Weight kg": "1"}],
    )
    _write_csv(
        orders_path,
        [
            {
                "Backer Number": "B-100",
                "Name": "Grace Hopper",
                "Email": "grace@example.com",
                "Address Line 1": "1 Compiler Way",
                "Address Type": "Residential",
                "Customer ID": "42",
                "Shipping Service": "Express",
                "Fulfillment Status": "Ready",
                "Core": "1",
            }
        ],
    )

    inspected = inspect_workbook(str(sku_master_path), str(orders_path))
    assert "Core" in inspected["detected_product_quantity_columns"]
    for metadata_column in [
        "Backer Number",
        "Name",
        "Email",
        "Address Line 1",
        "Address Type",
        "Customer ID",
        "Shipping Service",
        "Fulfillment Status",
    ]:
        assert metadata_column in inspected["detected_order_columns"]
        assert metadata_column not in inspected["detected_product_quantity_columns"]

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={"preserve_region_sheets": False},
    )

    row = _sheet_rows(output_path, "Order Volume Weights")[0]
    assert row["Backer Number"] == "B-100"
    assert row["Name"] == "Grace Hopper"
    assert row["Address Type"] == "Residential"
    assert row["Customer ID"] == "42"
    assert row["Address Line 1"] == "1 Compiler Way"
    assert list(row).index("Backer Number") > list(row).index("Warning Summary")


def test_second_row_xlsx_headers_preserve_metadata_in_order_volume_weights(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.xlsx"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [{"SKU": "Core", "Product Name": "Core", "Length": "5", "Width": "5", "Height": "5", "Weight kg": "1"}],
    )
    _write_xlsx_table(
        orders_path,
        "Orders",
        [
            ["Campaign export", "", "", "", ""],
            ["Backer Number", "Name", "Email", "Address Type", "Core"],
            ["B-200", "Ada Lovelace", "ada@example.com", "Residential", "1"],
        ],
    )

    inspected = inspect_workbook(str(sku_master_path), str(orders_path))
    assert "Core" in inspected["detected_product_quantity_columns"]
    for metadata_column in ["Backer Number", "Name", "Email", "Address Type"]:
        assert metadata_column in inspected["detected_order_columns"]

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={"preserve_region_sheets": False},
    )

    row = _sheet_rows(output_path, "Order Volume Weights")[0]
    headers = list(row)
    assert row["Backer Number"] == "B-200"
    assert row["Name"] == "Ada Lovelace"
    assert row["Email"] == "ada@example.com"
    assert row["Address Type"] == "Residential"
    assert headers.index("Backer Number") > headers.index("Warning Summary")


def test_packing_detail_uses_clear_coordinate_columns(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [{"SKU": "A", "Product Name": "A", "Length": "5", "Width": "5", "Height": "5", "Weight kg": "1"}],
    )
    _write_csv(orders_path, [{"Order ID": "1", "SKU": "A", "Quantity": "1"}])

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={"preserve_region_sheets": False},
    )

    row = _sheet_rows(output_path, "Packing Detail")[0]
    assert "Placement X cm" in row
    assert "Placement Y cm" in row
    assert "Placement Z cm" in row
    assert "X cm" not in row
    assert row["Placement Note"] == "Placed item coordinate"


def test_multi_box_detail_is_one_row_per_order_box_with_combined_skus(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {"SKU": "Core", "Product Name": "Core", "Length": "31", "Width": "22", "Height": "8", "Weight kg": "1"},
            {"SKU": "Addon", "Product Name": "Addon", "Length": "8", "Width": "5", "Height": "2", "Weight kg": "0.2"},
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "Asia-2", "SKU": "Core", "Quantity": "1"},
            {"Order ID": "Asia-2", "SKU": "Addon", "Quantity": "2"},
        ],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={
            "sku_rules": {"Core": {"prepacked": True, "ships_alone": True}},
            "packing_mode": "fast",
            "preserve_region_sheets": False,
        },
    )

    rows = _sheet_rows(output_path, "Multi Box Detail")
    assert len(rows) == 2
    assert {row["Order Box ID"] for row in rows} == {"Asia-2-1", "Asia-2-2"}
    assert any("ADDON x2" in row["SKUs in Box"] for row in rows)


def test_fast_mode_combines_large_item_and_small_addons_when_they_fit(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {"SKU": "Large", "Product Name": "Large", "Length": "40", "Width": "25", "Height": "15", "Weight kg": "2"},
            {"SKU": "Small A", "Product Name": "Small A", "Length": "6", "Width": "5", "Height": "2", "Weight kg": "0.2"},
            {"SKU": "Small B", "Product Name": "Small B", "Length": "5", "Width": "5", "Height": "2", "Weight kg": "0.2"},
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "1", "SKU": "Large", "Quantity": "1"},
            {"Order ID": "1", "SKU": "Small A", "Quantity": "1"},
            {"Order ID": "1", "SKU": "Small B", "Quantity": "1"},
        ],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={"packing_mode": "fast", "preserve_region_sheets": False},
    )

    row = _sheet_rows(output_path, "Order Volume Weights")[0]
    assert int(float(row["Box Qty"])) == 1
    detail = _sheet_rows(output_path, "Multi Box Detail")
    assert len(detail) == 1
    assert "LARGE x1" in detail[0]["SKUs in Box"]
    assert "SMALL A x1" in detail[0]["SKUs in Box"]


def test_asia_like_fast_mode_fixture_does_not_split_small_items_into_six_boxes(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {"SKU": "3D Terrain Pack", "Product Name": "3D Terrain Pack", "Length": "34", "Width": "24", "Height": "8", "Weight kg": "1.2"},
            {"SKU": "Dice Tray", "Product Name": "Dice Tray", "Length": "20", "Width": "15", "Height": "3", "Weight kg": "0.3"},
            {"SKU": "Earth Under Siege Core Box", "Product Name": "Earth Under Siege Core Box", "Length": "31", "Width": "22", "Height": "8", "Weight kg": "1.5"},
            {"SKU": "Gametrayz Campaign Trays", "Product Name": "Gametrayz Campaign Trays", "Length": "34", "Width": "22", "Height": "5", "Weight kg": "0.8"},
            {"SKU": "Plastic Token Pack", "Product Name": "Plastic Token Pack", "Length": "12", "Width": "8", "Height": "2", "Weight kg": "0.2"},
            {"SKU": "Token Upgrade Pack", "Product Name": "Token Upgrade Pack", "Length": "13", "Width": "9", "Height": "2", "Weight kg": "0.2"},
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "Asia-2", "SKU": "3D Terrain Pack", "Quantity": "1"},
            {"Order ID": "Asia-2", "SKU": "Dice Tray", "Quantity": "1"},
            {"Order ID": "Asia-2", "SKU": "Earth Under Siege Core Box", "Quantity": "1"},
            {"Order ID": "Asia-2", "SKU": "Gametrayz Campaign Trays", "Quantity": "1"},
            {"Order ID": "Asia-2", "SKU": "Plastic Token Pack", "Quantity": "1"},
            {"Order ID": "Asia-2", "SKU": "Token Upgrade Pack", "Quantity": "1"},
        ],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={"packing_mode": "fast", "preserve_region_sheets": False},
    )

    row = _sheet_rows(output_path, "Order Volume Weights")[0]
    assert int(float(row["Box Qty"])) < 6


def test_pledge_combination_summary_expands_multi_box_patterns(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {"SKU": "Core", "Product Name": "Core", "Length": "31", "Width": "22", "Height": "8", "Weight kg": "1"},
            {"SKU": "Addon", "Product Name": "Addon", "Length": "8", "Width": "5", "Height": "2", "Weight kg": "0.2"},
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "1", "SKU": "Core", "Quantity": "1"},
            {"Order ID": "1", "SKU": "Addon", "Quantity": "1"},
            {"Order ID": "2", "SKU": "Core", "Quantity": "1"},
            {"Order ID": "2", "SKU": "Addon", "Quantity": "1"},
        ],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={
            "sku_rules": {"Core": {"prepacked": True, "ships_alone": True}},
            "packing_mode": "fast",
            "preserve_region_sheets": False,
        },
    )

    rows = _sheet_rows(output_path, "Pledge Combination Summary")
    assert len(rows) == 2
    assert {int(float(row["Box Number"])) for row in rows} == {1, 2}
    assert {int(float(row["Box Qty"])) for row in rows} == {2}
    assert all(int(float(row["Order Count"])) == 2 for row in rows)
    assert all(float(row["Length cm"]).is_integer() for row in rows)
    assert all("Chargeable Weight kg" in row for row in rows)


def test_prepacked_sku_without_ships_alone_can_mix_with_addons(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {"SKU": "Core", "Product Name": "Core", "Length": "31", "Width": "22", "Height": "8", "Weight kg": "1"},
            {"SKU": "Addon", "Product Name": "Addon", "Length": "8", "Width": "5", "Height": "2", "Weight kg": "0.2"},
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "1", "SKU": "Core", "Quantity": "1"},
            {"Order ID": "1", "SKU": "Addon", "Quantity": "1"},
        ],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={"packing_mode": "fast", "sku_rules": {"Core": {"prepacked": True}}, "preserve_region_sheets": False},
    )

    row = _sheet_rows(output_path, "Order Volume Weights")[0]
    assert int(float(row["Box Qty"])) == 1
    detail = _sheet_rows(output_path, "Multi Box Detail")[0]
    assert "CORE x1" in detail["SKUs in Box"]
    assert "ADDON x1" in detail["SKUs in Box"]
    assert float(detail["Length cm"]) > 31
    assert float(detail["Width cm"]) > 22
    assert float(detail["Height cm"]) > 8


def test_forced_box_sku_without_ships_alone_can_mix_when_it_fits(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {"SKU": "Core", "Product Name": "Core", "Length": "30", "Width": "20", "Height": "7", "Weight kg": "1"},
            {"SKU": "Addon", "Product Name": "Addon", "Length": "8", "Width": "5", "Height": "2", "Weight kg": "0.2"},
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "1", "SKU": "Core", "Quantity": "1"},
            {"Order ID": "1", "SKU": "Addon", "Quantity": "1"},
        ],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={
            "packing_mode": "fast",
            "sku_rules": {"Core": {"prepacked": True, "forced_box_cm": [31, 22, 8]}},
            "preserve_region_sheets": False,
        },
    )

    row = _sheet_rows(output_path, "Order Volume Weights")[0]
    assert int(float(row["Box Qty"])) == 1
    assert "CORE x1" in _sheet_rows(output_path, "Multi Box Detail")[0]["SKUs in Box"]
    assert "ADDON x1" in _sheet_rows(output_path, "Multi Box Detail")[0]["SKUs in Box"]


def test_prepacked_ships_alone_and_forced_no_mix_rules_split(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {"SKU": "Solo", "Product Name": "Solo", "Length": "30", "Width": "20", "Height": "7", "Weight kg": "1"},
            {"SKU": "Forced", "Product Name": "Forced", "Length": "10", "Width": "8", "Height": "3", "Weight kg": "0.5"},
            {"SKU": "Addon", "Product Name": "Addon", "Length": "8", "Width": "5", "Height": "2", "Weight kg": "0.2"},
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "1", "SKU": "Solo", "Quantity": "1"},
            {"Order ID": "1", "SKU": "Addon", "Quantity": "1"},
            {"Order ID": "2", "SKU": "Forced", "Quantity": "1"},
            {"Order ID": "2", "SKU": "Addon", "Quantity": "1"},
        ],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={
            "packing_mode": "fast",
            "sku_rules": {
                "Solo": {"prepacked": True, "ships_alone": True},
                "Forced": {"forced_box_cm": [15, 10, 5], "can_mix_with_other_items": False},
            },
            "preserve_region_sheets": False,
        },
    )

    rows = {row["Order ID"]: row for row in _sheet_rows(output_path, "Order Volume Weights")}
    assert int(float(rows["1"]["Box Qty"])) == 2
    assert int(float(rows["2"]["Box Qty"])) == 2


def test_prepacked_sku_splits_when_addons_do_not_fit(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {"SKU": "Core", "Product Name": "Core", "Length": "70", "Width": "33", "Height": "40", "Weight kg": "1"},
            {"SKU": "Addon", "Product Name": "Addon", "Length": "20", "Width": "15", "Height": "3", "Weight kg": "0.2"},
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "1", "SKU": "Core", "Quantity": "1"},
            {"Order ID": "1", "SKU": "Addon", "Quantity": "1"},
        ],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={"packing_mode": "fast", "sku_rules": {"Core": {"prepacked": True}}, "preserve_region_sheets": False},
    )

    row = _sheet_rows(output_path, "Order Volume Weights")[0]
    assert int(float(row["Box Qty"])) > 1


def test_default_order_volume_excludes_source_fields_but_keeps_real_metadata(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [{"SKU": "Core", "Product Name": "Core", "Length": "5", "Width": "5", "Height": "5", "Weight kg": "1"}],
    )
    _write_csv(
        orders_path,
        [{"Order ID": "1", "SKU": "Core", "Quantity": "1", "Name": "Jane Backer", "Email": "jane@example.com"}],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={"preserve_region_sheets": False},
    )

    row = _sheet_rows(output_path, "Order Volume Weights")[0]
    assert "_source_file" not in row
    assert "_source_sheet" not in row
    assert row["Name"] == "Jane Backer"
    assert row["Email"] == "jane@example.com"

    mapping_rows = _sheet_rows(output_path, "Input Column Mapping")
    assert mapping_rows[0]["workbook"]
