import csv
import zipfile
from pathlib import Path

from box_optimizer import optimize_workbook
from box_optimizer.io.excel_reader import read_workbook
from box_optimizer.workflow import inspect_workbook


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


def test_output_carton_dimensions_never_exceed_cap(tmp_path):
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
        assert float(row["Length cm"]) <= 74
        assert float(row["Width cm"]) <= 37
        assert float(row["Height cm"]) <= 44
        assert float(row["Assigned Box Length cm"]) <= 74
        assert float(row["Assigned Box Width cm"]) <= 37
        assert float(row["Assigned Box Height cm"]) <= 44


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
        config={"sku_rules": {"Flat": {"no_padding": True}}, "preserve_region_sheets": False},
    )

    row = _sheet_rows(output_path, "Order Volume Weights")[0]
    assert float(row["Optimized Length cm"]) == 11
    assert float(row["Optimized Width cm"]) == 6
    assert float(row["Optimized Height cm"]) == 4


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
        config={"sku_rules": {"Core": {"prepacked": True}}, "preserve_region_sheets": False},
    )

    row = _sheet_rows(output_path, "Order Volume Weights")[0]
    assert float(row["Optimized Length cm"]) == 32
    assert float(row["Optimized Width cm"]) == 23
    assert float(row["Optimized Height cm"]) == 9


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
