import csv
import json
import logging
import zipfile
from collections import Counter
from xml.etree import ElementTree
from pathlib import Path

from box_optimizer import optimize_workbook
from box_optimizer.io import excel_writer as excel_writer_module
from box_optimizer.io.excel_reader import read_workbook
from box_optimizer.models import Dimensions, OrderLine, SKUItem
from box_optimizer.packing.packer import OptimizedCartonResult, Placement
from box_optimizer.packing.splitter import SplitCarton, SplitResult
from box_optimizer import rate_sources as rate_source_module
from box_optimizer.rate_sources import active_rate_sheet_path, rate_sheet_metadata_path, sha256_file
from box_optimizer.weights import packed_actual_weight_kg
import box_optimizer.workflow as workflow_module
from box_optimizer.workflow import SKUCampaignRule, _packed_items_for_order, format_kg_display, inspect_workbook


_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}




def _sku_item(sku: str, dimensions: Dimensions, weight_kg: float = 1) -> SKUItem:
    return SKUItem(
        raw_sku=sku,
        canonical_sku=sku,
        product_name=sku,
        length_cm=dimensions.length,
        width_cm=dimensions.width,
        height_cm=dimensions.height,
        weight_kg=weight_kg,
    )


def _order_line(order_id: str, sku: str, quantity: int = 1) -> OrderLine:
    return OrderLine(order_id=order_id, raw_sku=sku, canonical_sku=sku, quantity=quantity)


def test_vfi_intake_received_column_maps_quantity_aliases_for_reporting():
    rows = workflow_module._vfi_intake_form_rows_with_received(
        [
            {"SKU": "CORE", "quantity": "100"},
            {"SKU": "EXP", "Quantity": "25"},
            {"SKU": "TOKEN", "数量": "12"},
            {"SKU": "ADDON", "Stock Received": "7"},
        ]
    )

    assert [row["Received"] for row in rows] == ["100", "25", "12", "7"]


def test_sku_intake_summary_rows_include_intake_and_order_skus_with_remaining_quantities():
    sku_items = [
        _sku_item("CORE", Dimensions(10, 10, 2)),
        _sku_item("INTAKEONLY", Dimensions(5, 5, 1)),
    ]
    order_lines = [
        _order_line("1", "CORE", 2),
        _order_line("2", "CORE", 3),
        _order_line("3", "ORDERONLY", 4),
    ]
    intake_rows = workflow_module._vfi_intake_form_rows_with_received(
        [
            {"SKU": "CORE", "quantity": "10"},
            {"SKU": "INTAKEONLY", "quantity": "6"},
        ]
    )

    rows = workflow_module._sku_intake_summary_rows(sku_items, order_lines, intake_rows)
    by_sku = {row["SKU"]: row for row in rows}

    assert list(by_sku) == ["CORE", "INTAKEONLY", "ORDERONLY"]
    assert by_sku["CORE"]["Received Quantity"] == 10
    assert by_sku["CORE"]["Required Quantity"] == 5
    assert by_sku["CORE"]["Remaining"] == 5
    assert by_sku["INTAKEONLY"]["Required Quantity"] == 0
    assert by_sku["INTAKEONLY"]["Remaining"] == 6
    assert by_sku["ORDERONLY"]["Received Quantity"] == 0
    assert by_sku["ORDERONLY"]["Required Quantity"] == 4
    assert by_sku["ORDERONLY"]["Remaining"] == -4


def test_single_order_configurations_sort_by_country_after_repeated_configurations():
    box_rows = [
        {"Order ID": "R1", "Box Number": 1, "SKU Breakdown": "REPEAT A x1", "Country": "Japan", "Box Type": "VB 1", "Chargeable Weight kg": 2},
        {"Order ID": "R2", "Box Number": 1, "SKU Breakdown": "REPEAT A x1", "Country": "China", "Box Type": "VB 1", "Chargeable Weight kg": 2},
        {"Order ID": "R3", "Box Number": 1, "SKU Breakdown": "REPEAT A x1", "Country": "Hong Kong", "Box Type": "VB 1", "Chargeable Weight kg": 2},
        {"Order ID": "R4", "Box Number": 1, "SKU Breakdown": "REPEAT B x1", "Country": "Singapore", "Box Type": "VB 2", "Chargeable Weight kg": 3},
        {"Order ID": "R5", "Box Number": 1, "SKU Breakdown": "REPEAT B x1", "Country": "Malaysia", "Box Type": "VB 2", "Chargeable Weight kg": 3},
        {"Order ID": "S1", "Box Number": 1, "SKU Breakdown": "ALPHA JAPAN x1", "Country": "Japan", "Box Type": "VB 3", "Chargeable Weight kg": 4},
        {"Order ID": "S2", "Box Number": 1, "SKU Breakdown": "OMEGA CHINA B x1", "Country": "China", "Box Type": "VB 4", "Chargeable Weight kg": 5},
        {"Order ID": "S3", "Box Number": 1, "SKU Breakdown": "BETA HONG KONG x1", "Country": "Hong Kong", "Box Type": "VB 5", "Chargeable Weight kg": 6},
        {"Order ID": "S4", "Box Number": 1, "SKU Breakdown": "DELTA CHINA A x1", "Country": "China", "Box Type": "VB 6", "Chargeable Weight kg": 7},
        {"Order ID": "S5", "Box Number": 1, "SKU Breakdown": "AAA MISSING COUNTRY x1", "Country": "", "Box Type": "VB 7", "Chargeable Weight kg": 8},
    ]

    ordered_combos = [combo for combo, _entry in workflow_module._combo_entries_for_optimized_to_pack(box_rows)]
    pledge_config_by_combo = workflow_module._pledge_config_by_combo(box_rows)

    assert ordered_combos == [
        "REPEAT A x1",
        "REPEAT B x1",
        "DELTA CHINA A x1",
        "OMEGA CHINA B x1",
        "BETA HONG KONG x1",
        "ALPHA JAPAN x1",
        "AAA MISSING COUNTRY x1",
    ]
    assert pledge_config_by_combo == {
        "REPEAT A x1": 1,
        "REPEAT B x1": 2,
        "DELTA CHINA A x1": 3,
        "OMEGA CHINA B x1": 4,
        "BETA HONG KONG x1": 5,
        "ALPHA JAPAN x1": 6,
        "AAA MISSING COUNTRY x1": 7,
    }
    label_rows = workflow_module._label_generator_rows(box_rows, pledge_config_by_combo)
    config_by_order = {row["Order ID"]: row["Pledge Configuration"] for row in label_rows}
    assert config_by_order["S4"] == 3
    assert config_by_order["S2"] == 4
    assert config_by_order["S3"] == 5
    assert config_by_order["S1"] == 6
    assert config_by_order["S5"] == 7

    ordered_entries = workflow_module._combo_entries_for_optimized_to_pack(box_rows)
    all_order_ids = sorted(order_id for _combo, entry in ordered_entries for order_id in entry["order_ids"])
    all_box_types = sorted(row["Box Type"] for _combo, entry in ordered_entries for rows in entry["boxes"].values() for row in rows)
    all_weights = sorted(row["Chargeable Weight kg"] for _combo, entry in ordered_entries for rows in entry["boxes"].values() for row in rows)

    assert all_order_ids == [row["Order ID"] for row in sorted(box_rows, key=lambda row: row["Order ID"])]
    assert all_box_types == [row["Box Type"] for row in sorted(box_rows, key=lambda row: row["Box Type"])]
    assert all_weights == [row["Chargeable Weight kg"] for row in sorted(box_rows, key=lambda row: row["Chargeable Weight kg"])]


def test_similar_footprint_items_bundle_before_padding_once():
    sku_lookup = {
        "A": _sku_item("A", Dimensions(30, 25, 6.5), 0.4),
        "B": _sku_item("B", Dimensions(25, 25, 9.5), 0.3),
    }

    items = _packed_items_for_order(
        [_order_line("1", "A"), _order_line("1", "B")],
        sku_lookup,
        {},
        bundle_footprint_tolerance_cm=5,
    )

    assert len(items) == 1
    assert items[0].canonical_sku == "BUNDLE[A x1 | B x1]"
    assert items[0].unpadded_dimensions == Dimensions(30, 25, 16)
    assert items[0].padded_dimensions == Dimensions(32, 27, 18)
    assert round(items[0].weight_kg, 1) == 0.7


def test_too_tall_similar_footprint_bundle_falls_back_to_individual_items():
    sku_lookup = {"BIG": _sku_item("BIG", Dimensions(42, 31, 27), 2)}

    items = _packed_items_for_order(
        [_order_line("1", "BIG", quantity=3)],
        sku_lookup,
        {},
        bundle_footprint_tolerance_cm=5,
    )

    assert len(items) == 1
    assert items[0].canonical_sku == "BIG"
    assert items[0].quantity == 3
    assert not items[0].canonical_sku.startswith("BUNDLE[")


def test_no_padding_items_do_not_bundle_without_explicit_bundle_rule():
    sku_lookup = {
        "A": _sku_item("A", Dimensions(30, 25, 6.5)),
        "B": _sku_item("B", Dimensions(25, 25, 9.5)),
    }

    items = _packed_items_for_order(
        [_order_line("1", "A"), _order_line("1", "B")],
        sku_lookup,
        {"A": SKUCampaignRule(key="A", no_padding=True)},
        bundle_footprint_tolerance_cm=5,
    )

    assert len(items) == 2
    assert {item.canonical_sku for item in items} == {"A", "B"}
    assert next(item for item in items if item.canonical_sku == "A").padded_dimensions == Dimensions(30, 25, 6.5)


def test_compressible_items_shrink_before_packing_and_skip_normal_padding():
    sku_lookup = {"PLUSH": _sku_item("PLUSH", Dimensions(40, 30, 10), 0.5)}

    items = _packed_items_for_order(
        [_order_line("1", "PLUSH")],
        sku_lookup,
        {
            "PLUSH": SKUCampaignRule(
                key="PLUSH",
                compressible=True,
                compressed_height_ratio=0.5,
                compressed_volume_ratio=0.75,
            )
        },
    )

    assert len(items) == 1
    assert items[0].unpadded_dimensions == Dimensions(40, 30, 10)
    assert items[0].padded_dimensions == Dimensions(40, 30, 5)
    assert "compressible" in items[0].rule_applied


def test_compressible_items_do_not_bundle_as_rigid_stacks():
    sku_lookup = {
        "PLUSH": _sku_item("PLUSH", Dimensions(40, 30, 10), 0.5),
        "BOOK": _sku_item("BOOK", Dimensions(39, 29, 2), 0.8),
    }

    items = _packed_items_for_order(
        [_order_line("1", "PLUSH"), _order_line("1", "BOOK")],
        sku_lookup,
        {"PLUSH": SKUCampaignRule(key="PLUSH", compressible=True)},
    )

    assert len(items) == 2
    assert not any(item.canonical_sku.startswith("BUNDLE[") for item in items)


def test_wrap_around_largest_item_uses_largest_other_footprint_without_padding():
    sku_lookup = {
        "Playmat": _sku_item("Playmat", Dimensions(70, 35, 1), 0.5),
        "Core": _sku_item("Core", Dimensions(30, 20, 7), 1.5),
    }

    items = _packed_items_for_order(
        [_order_line("1", "Playmat"), _order_line("1", "Core")],
        sku_lookup,
        {"Playmat": SKUCampaignRule(key="Playmat", wrap_around_largest_item=True, wrapped_height_cm=4)},
        bundle_footprint_tolerance_cm=5,
    )

    playmat = next(item for item in items if item.canonical_sku == "Playmat")
    assert playmat.unpadded_dimensions == Dimensions(30, 20, 4)
    assert playmat.padded_dimensions == Dimensions(30, 20, 4)
    assert playmat.rule_applied == "Playmat wrap around largest item"


def test_wrap_around_largest_item_keeps_solo_footprint_with_wrapped_height():
    sku_lookup = {"Playmat": _sku_item("Playmat", Dimensions(70, 35, 1), 0.5)}

    items = _packed_items_for_order(
        [_order_line("1", "Playmat")],
        sku_lookup,
        {
            "Playmat": SKUCampaignRule(
                key="Playmat",
                no_padding=True,
                wrap_around_largest_item=True,
                wrapped_height_cm=4,
            )
        },
        bundle_footprint_tolerance_cm=5,
    )

    assert len(items) == 1
    assert items[0].unpadded_dimensions == Dimensions(70, 35, 4)
    assert items[0].padded_dimensions == Dimensions(70, 35, 4)


def test_no_mix_items_do_not_bundle():
    sku_lookup = {
        "A": _sku_item("A", Dimensions(30, 25, 6.5)),
        "B": _sku_item("B", Dimensions(25, 25, 9.5)),
    }

    items = _packed_items_for_order(
        [_order_line("1", "A"), _order_line("1", "B")],
        sku_lookup,
        {"A": SKUCampaignRule(key="A", can_mix_with_other_items=False)},
        bundle_footprint_tolerance_cm=5,
    )

    assert len(items) == 2
    assert all(not item.canonical_sku.startswith("BUNDLE[") for item in items)

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
            _inline_cell(excel_writer_module._column_letter(column), row_number, value)
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
            _inline_cell(excel_writer_module._column_letter(column), row_number, value)
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


def _write_xlsx_tables(path: Path, sheets: dict[str, list[list[object]]]) -> None:
    sheet_entries = []
    rel_entries = []
    worksheet_payloads = []
    for index, (sheet_name, table) in enumerate(sheets.items(), start=1):
        xml_rows = []
        for row_number, row_values in enumerate(table, start=1):
            cells = [
                _inline_cell(excel_writer_module._column_letter(column), row_number, value)
                for column, value in enumerate(row_values)
            ]
            xml_rows.append(f'<row r="{row_number}">{"".join(cells)}</row>')
        worksheet_payloads.append(
            (
                f"xl/worksheets/sheet{index}.xml",
                '<?xml version="1.0" encoding="UTF-8"?>'
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                f'<sheetData>{"".join(xml_rows)}</sheetData>'
                "</worksheet>",
            )
        )
        sheet_entries.append(f'<sheet name="{sheet_name}" sheetId="{index}" r:id="rId{index}"/>')
        rel_entries.append(
            f'<Relationship Id="rId{index}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
        )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheets>{"".join(sheet_entries)}</sheets>'
        "</workbook>"
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'{"".join(rel_entries)}'
        "</Relationships>"
    )
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("xl/workbook.xml", workbook)
        archive.writestr("xl/_rels/workbook.xml.rels", rels)
        for worksheet_path, worksheet_xml in worksheet_payloads:
            archive.writestr(worksheet_path, worksheet_xml)


def _wide_row(width: int, values: dict[int, object]) -> list[object]:
    row = [""] * width
    for index, value in values.items():
        row[index] = value
    return row


def _sheet_rows(path: Path, sheet_name: str) -> list[dict]:
    return next(sheet.rows for sheet in read_workbook(str(path)) if sheet.sheet_name == sheet_name)


def _workbook_sheet_names(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as archive:
        workbook_root = ElementTree.fromstring(archive.read("xl/workbook.xml"))
    return [
        sheet.attrib["name"]
        for sheet in workbook_root.findall("main:sheets/main:sheet", _NS)
    ]


def _sheet_xml(path: Path, sheet_name: str) -> str:
    with zipfile.ZipFile(path) as archive:
        workbook_root = ElementTree.fromstring(archive.read("xl/workbook.xml"))
        rels_root = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rel_targets = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in rels_root.findall("pkgrel:Relationship", _NS)
        }
        for sheet in workbook_root.findall("main:sheets/main:sheet", _NS):
            if sheet.attrib["name"] == sheet_name:
                rel_id = sheet.attrib[f"{{{_NS['rel']}}}id"]
                target = rel_targets[rel_id].lstrip("/")
                sheet_path = target if target.startswith("xl/") else f"xl/{target}"
                return archive.read(sheet_path).decode("utf-8")
    raise AssertionError(f"Sheet not found: {sheet_name}")


def _inline_cell_text(sheet_xml: str, reference: str) -> str:
    root = ElementTree.fromstring(sheet_xml)
    cell = root.find(f".//main:c[@r='{reference}']", _NS)
    if cell is None:
        return ""
    text = cell.find(".//main:t", _NS)
    return "" if text is None else text.text or ""


def _cell_formula(sheet_xml: str, reference: str) -> str:
    root = ElementTree.fromstring(sheet_xml)
    formula = root.find(f".//main:c[@r='{reference}']/main:f", _NS)
    return "" if formula is None else formula.text or ""


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
    sheet_names = _workbook_sheet_names(output_path)
    assert sheet_names[:4] == ["Summary", "Cost Summary", "Actual Dimensions", "Labels"]
    for required_sheet in [
        "VFI Intake Form",
        "Optimized to Pack",
        "Label generator",
        "Order Volume Weights",
        "Box Size Summary",
    ]:
        assert required_sheet in sheet_names
    order_volume_rows = next(sheet.rows for sheet in workbook_rows if sheet.sheet_name == "Order Volume Weights")
    assert order_volume_rows[0]["Order ID"] == "1001"
    assert order_volume_rows[0]["US State Abbreviation"] == "CA"
    assert order_volume_rows[0]["Pledge Level"] == "Deluxe"
    assert order_volume_rows[0]["SKU Breakdown"] == "CORE GAME x1"


def test_optimize_workbook_adds_actual_dimensions_after_cost_summary_with_barcode_formulas(tmp_path, monkeypatch):
    monkeypatch.setenv("BOX_OPTIMIZER_RATE_SHEET_DIR", str(tmp_path / "managed_rates"))
    monkeypatch.delenv("BOX_OPTIMIZER_RATE_SYNC_URL", raising=False)
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized_shipping_plan.xlsx"
    rate_path = tmp_path / "rates.xlsx"
    _write_xlsx_table(
        rate_path,
        "Zone Key",
        [
            ["HUB", "", "", ""],
            ["KG", 0.5, 1, 1.5],
            ["Zone USA", 11.15, 12.25, 13.35],
            ["", "", "", ""],
            ["Zone 0", "Zone USA"],
            ["United States", "USA"],
        ],
    )
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
                "Country": "US",
                "State": "California",
            }
        ],
    )

    optimize_workbook(
        sku_master_path=str(sku_master_path),
        orders_path=str(orders_path),
        output_path=str(output_path),
        config={**workflow_module.DEFAULT_CONFIG, "rate_sheet_path": str(rate_path)},
    )

    sheet_names = _workbook_sheet_names(output_path)
    assert sheet_names[1:3] == ["Cost Summary", "Actual Dimensions"]
    assert "_ActualLookupTable" in sheet_names
    assert "_ActualRateTable" in sheet_names

    cost_xml = _sheet_xml(output_path, "Cost Summary")
    cost_root = ElementTree.fromstring(cost_xml)
    cost_headers = [
        cell.find(".//main:t", _NS).text
        for cell in cost_root.findall("main:sheetData/main:row[@r='1']/main:c", _NS)
    ]
    assert cost_headers[-3:] == ["Final weight kg", "Final cost", "Scan note"]
    cost_formulas = [
        formula.text or ""
        for formula in cost_root.findall("main:sheetData/main:row[@r='2']/main:c/main:f", _NS)
    ]
    cost_formula_text = "\n".join(cost_formulas)
    assert "XLOOKUP" not in cost_xml
    assert "LET(" not in cost_xml
    assert "INDEX('Actual Dimensions'!$Y:$Y,MATCH($B2,'Actual Dimensions'!$O:$O,0))" in cost_formula_text
    assert "INDEX('Actual Dimensions'!$I:$I,MATCH($B2,'Actual Dimensions'!$O:$O,0))" in cost_formula_text
    assert "COUNTA('Actual Dimensions'!$A:$A)<=1" in cost_formula_text
    assert 'COUNTIFS(\'Actual Dimensions\'!$AE:$AE,$B2,\'Actual Dimensions\'!$M:$M,"Not Scanned")>0' in cost_formula_text
    assert 'COUNTIFS(\'Actual Dimensions\'!$P:$P,$B2,\'Actual Dimensions\'!$M:$M,"Not Scanned")>0' not in cost_formula_text
    assert "MATCH($B2,'Actual Dimensions'!$L:$L,0)" not in cost_formula_text
    assert "Item not scanned" in cost_formula_text
    cost_widths = {
        int(column.attrib["min"]): float(column.attrib["width"])
        for column in cost_root.findall("main:cols/main:col", _NS)
    }
    final_weight_column = len(cost_headers) - 2
    final_cost_column = len(cost_headers) - 1
    scan_note_column = len(cost_headers)
    assert cost_widths[final_weight_column] == 15
    assert cost_widths[final_cost_column] == 12
    assert cost_widths[scan_note_column] == 16
    final_weight_ref = f"{excel_writer_module._column_letter(final_weight_column - 1)}2"
    final_cost_ref = f"{excel_writer_module._column_letter(final_cost_column - 1)}2"
    assert cost_root.find(f".//main:c[@r='{final_weight_ref}']", _NS).attrib["s"] == "31"
    assert cost_root.find(f".//main:c[@r='{final_cost_ref}']", _NS).attrib["s"] == "19"
    assert cost_root.find(".//main:c[@r='P2']", _NS).attrib["s"] == "19"
    with zipfile.ZipFile(output_path) as archive:
        styles_root = ElementTree.fromstring(archive.read("xl/styles.xml"))
    cell_formats = styles_root.findall("main:cellXfs/main:xf", _NS)
    assert cell_formats[19].attrib["numFmtId"] == "165"
    assert cell_formats[31].attrib["numFmtId"] == "166"

    actual_xml = _sheet_xml(output_path, "Actual Dimensions")
    actual_root = ElementTree.fromstring(actual_xml)
    actual_headers = [
        cell.find(".//main:t", _NS).text or ""
        for cell in actual_root.findall("main:sheetData/main:row[@r='1']/main:c", _NS)
    ]
    assert actual_headers[:14] == [
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
        "",
    ]
    assert actual_headers[14:30] == [
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
    ]
    assert actual_headers[30] == "Expected scan group VFI key"
    assert _inline_cell_text(actual_xml, "N1") == ""

    assert "XLOOKUP" not in actual_xml
    assert "LET(" not in actual_xml
    assert 'ISNUMBER(SEARCH(" of ",$A2))' in actual_xml
    assert 'LEFT($A2,FIND(" ",$A2)-1)&amp;" "&amp;MID($A2,FIND("@",SUBSTITUTE($A2," ","@",4))+1,999)' in actual_xml
    assert 'IF($A2="","",IF(IF(ISNUMBER(SEARCH(" of ",$A2)),IFERROR(VALUE(MID($A2,FIND(" ",$A2)+1,FIND(" of ",$A2)-FIND(" ",$A2)-1))=1,FALSE),$A2=$P2),$P2,""))' in actual_xml
    assert 'VALUE(RIGHT($A2,LEN($A2)-FIND("@",SUBSTITUTE($A2,"-","@",LEN($A2)-LEN(SUBSTITUTE($A2,"-",""))))))&gt;1' in actual_xml
    assert "MATCH($P2,'_ActualLookupTable'!$A:$A,0)" in actual_xml
    assert "INDEX('_ActualLookupTable'!$H:$H" in actual_xml
    assert 'CEILING($C2,0.5)*CEILING($D2,0.5)*CEILING($E2,0.5)/5000' in actual_xml
    assert "MAX($B2/1000,$F2)" in actual_xml
    assert 'IF($B2&gt;$G2*1.2,"Greater than expected",IF($B2&lt;$G2*0.95,"Less than expected",""))' in actual_xml
    assert 'SUMIF($P:$P,$P2,$X:$X)' in actual_xml
    assert "INDEX('_ActualRateTable'!$D:$D,MATCH($Z2&amp;\"|\"&amp;" in actual_xml
    assert "$B2/1000" in actual_xml
    assert "CEILING($Y2-0.000000001,0.5)" in actual_xml
    assert '$Q2&lt;&gt;"Yes"' in actual_xml
    assert "SUMPRODUCT(--(TRIM($A$2:$A$2)=TRIM($L2)))" in actual_xml
    assert 'IF($L2="","",IFERROR(IF(ISNUMBER(SEARCH(" of ",$L2)),LEFT($L2,FIND(" ",$L2)-1)&amp;" "&amp;MID($L2,FIND("@",SUBSTITUTE($L2," ","@",4))+1,999)' in actual_xml
    assert 'VALUE(RIGHT($L2,LEN($L2)-FIND("@",SUBSTITUTE($L2,"-","@",LEN($L2)-LEN(SUBSTITUTE($L2,"-",""))))))&gt;1' in actual_xml
    assert "Not Scanned" in actual_xml
    assert "No barcode match" in actual_xml
    assert "No hub rate" in actual_xml
    assert "Dimension only" in actual_xml

    lookup_rows = _sheet_rows(output_path, "_ActualLookupTable")
    rate_rows = _sheet_rows(output_path, "_ActualRateTable")
    assert lookup_rows[0]["Country"] == "United States"
    assert lookup_rows[0]["Total Units"] == "1"
    assert lookup_rows[0]["Pick / add-on fee"] == "2.0"
    assert lookup_rows[0]["Quoted shipping cost"] == "15.35"
    assert lookup_rows[0]["Estimated weight g"]
    assert rate_rows[0]["Zone"] == "Zone USA"
    assert rate_rows[0]["Weight Band kg"] == "0.5"
    with zipfile.ZipFile(output_path) as archive:
        workbook_xml = archive.read("xl/workbook.xml").decode("utf-8")
    assert 'name="_ActualLookupTable"' in workbook_xml
    assert 'name="_ActualRateTable"' in workbook_xml
    assert 'state="hidden"' in workbook_xml


def test_actual_dimensions_multi_carton_formulas_group_by_base_barcode_and_gate_costs():
    row = workflow_module._actual_dimensions_rows(1)[0]

    lookup_formula = row["Cost Summary VFI #"].formula
    group_key_formula = row["Group VFI key"].formula
    expected_group_key_formula = row["Expected scan group VFI key"].formula
    charge_flag_formula = row["Is charge row"].formula
    group_weight_formula = row["Order/group chargeable weight kg"].formula
    matched_band_formula = row["Matched rate weight band"].formula
    hub_fee_formula = row["Actual hub shipping fee"].formula
    pick_fee_formula = row["Pick / add-on fee"].formula
    total_formula = row["Actual total shipping cost"].formula
    quoted_formula = row["Quoted shipping cost"].formula
    diff_formula = row["Actual vs quoted difference"].formula
    status_formula = row["Lookup status"].formula

    assert 'CEILING($C2,0.5)*CEILING($D2,0.5)*CEILING($E2,0.5)/5000' == row["Actual DIM weight kg"].formula.split(',"",', 1)[1][:-1]
    assert 'MAX($B2/1000,$F2)' in row["Actual chargeable weight kg"].formula
    assert 'INDEX(\'_ActualLookupTable\'!$H:$H' in row["Estimated weight in grams"].formula
    assert '$B2>$G2*1.2' in row["Weight warning"].formula
    assert '$B2<$G2*0.95' in row["Weight warning"].formula
    assert '"Greater than expected"' in row["Weight warning"].formula
    assert '"Less than expected"' in row["Weight warning"].formula
    assert "SUMPRODUCT(--(TRIM($A$2:$A$2)=TRIM($L2)))" in row["Scan status"].formula
    assert '"Not Scanned"' in row["Scan status"].formula
    assert 'IF($A2="","",IF(IF(ISNUMBER(SEARCH(" of ",$A2)),IFERROR(VALUE(MID($A2,FIND(" ",$A2)+1,FIND(" of ",$A2)-FIND(" ",$A2)-1))=1,FALSE),$A2=$P2),$P2,""))' == lookup_formula
    assert 'LEFT($A2,FIND(" ",$A2)-1)&" "&MID($A2,FIND("@",SUBSTITUTE($A2," ","@",4))+1,999)' in group_key_formula
    assert 'VALUE(RIGHT($A2,LEN($A2)-FIND("@",SUBSTITUTE($A2,"-","@",LEN($A2)-LEN(SUBSTITUTE($A2,"-",""))))))>1' in group_key_formula
    assert 'LEFT($A2,FIND("@",SUBSTITUTE($A2,"-","@",LEN($A2)-LEN(SUBSTITUTE($A2,"-",""))))-1)' in group_key_formula
    assert 'IF($L2="","",' in expected_group_key_formula
    assert 'LEFT($L2,FIND(" ",$L2)-1)&" "&MID($L2,FIND("@",SUBSTITUTE($L2," ","@",4))+1,999)' in expected_group_key_formula
    assert 'VALUE(RIGHT($L2,LEN($L2)-FIND("@",SUBSTITUTE($L2,"-","@",LEN($L2)-LEN(SUBSTITUTE($L2,"-",""))))))>1' in expected_group_key_formula
    assert 'LEFT($L2,FIND("@",SUBSTITUTE($L2,"-","@",LEN($L2)-LEN(SUBSTITUTE($L2,"-",""))))-1)' in expected_group_key_formula
    assert 'IF($A2="","",IF(IF(ISNUMBER(SEARCH(" of ",$A2)),IFERROR(VALUE(MID($A2,FIND(" ",$A2)+1,FIND(" of ",$A2)-FIND(" ",$A2)-1))=1,FALSE),$A2=$P2),"Yes","No"))' == charge_flag_formula
    assert "MATCH($P2,'_ActualLookupTable'!$A:$A,0)" in row["Country"].formula
    assert "SUMIF($P:$P,$P2,$X:$X)" in group_weight_formula
    assert "CEILING($Y2-0.000000001,0.5)" in matched_band_formula

    for cost_formula in [
        matched_band_formula,
        hub_fee_formula,
        pick_fee_formula,
        total_formula,
        quoted_formula,
        diff_formula,
    ]:
        assert '$Q2<>"Yes"' in cost_formula

    assert "Dimension only" in status_formula
    assert "XLOOKUP" not in "".join(value.formula for value in row.values() if hasattr(value, "formula"))
    assert "LET(" not in "".join(value.formula for value in row.values() if hasattr(value, "formula"))


def test_actual_rate_table_keys_match_excel_numeric_band_text_for_whole_and_repeat_weights():
    rates = {step / 2: step for step in range(1, 99)}
    rate_sheet = workflow_module.CustomerRateSheet(
        hub=workflow_module.CustomerRateLane(
            rates_by_zone={"Zone 0": rates},
            zone_by_country={"united states": "Zone 0"},
            max_weight_kg=49.0,
        ),
        express=workflow_module.CustomerRateLane(rates_by_zone={}, zone_by_country={}),
    )

    rate_rows = workflow_module._actual_rate_rows(rate_sheet)
    rates_by_key = {row["Rate Key"]: row["Hub Rate"] for row in rate_rows}

    for expected_key in ["Zone 0|0.5", "Zone 0|1", "Zone 0|6", "Zone 0|9", "Zone 0|10.5", "Zone 0|40", "Zone 0|49"]:
        assert expected_key in rates_by_key
    for old_key in ["Zone 0|1.0", "Zone 0|6.0", "Zone 0|9.0", "Zone 0|11.0", "Zone 0|40.0", "Zone 0|49.0"]:
        assert old_key not in rates_by_key

    assert next(row for row in rate_rows if row["Rate Key"] == "Zone 0|6")["Weight Band kg"] == 6.0
    assert next(row for row in rate_rows if row["Rate Key"] == "Zone 0|9")["Weight Band kg"] == 9.0
    assert next(row for row in rate_rows if row["Rate Key"] == "Zone 0|10.5")["Weight Band kg"] == 10.5

    def excel_numeric_text(value: float) -> str:
        return str(int(value)) if float(value).is_integer() else f"{value:g}"

    def excel_ceiling_half(value: float) -> float:
        return int((value * 2) + 0.999999999) / 2

    max_weight = rate_rows[0]["Max Weight kg"]
    max_rate = rate_rows[0]["Max Weight Rate"]

    def simulated_actual_hub_fee(matched_band: float) -> float:
        repeated_max_weights = int((matched_band - 0.000000001) / max_weight)
        remainder = excel_ceiling_half((matched_band - 0.000000001) % max_weight)
        lookup_band = max_weight if remainder == 0 else remainder
        return repeated_max_weights * max_rate + rates_by_key[f"Zone 0|{excel_numeric_text(lookup_band)}"]

    assert simulated_actual_hub_fee(6) == rates_by_key["Zone 0|6"]
    assert simulated_actual_hub_fee(9) == rates_by_key["Zone 0|9"]
    assert simulated_actual_hub_fee(10.5) == rates_by_key["Zone 0|10.5"]
    assert simulated_actual_hub_fee(60) == rates_by_key["Zone 0|49"] + rates_by_key["Zone 0|11"]
    assert simulated_actual_hub_fee(66.5) == rates_by_key["Zone 0|49"] + rates_by_key["Zone 0|17.5"]

    formula_text = "".join(value.formula for value in workflow_module._actual_dimensions_rows(1)[0].values() if hasattr(value, "formula"))
    assert "XLOOKUP" not in formula_text
    assert "LET(" not in formula_text


def test_actual_dimensions_barcode_parts_separate_scanner_barcode_lookup_and_group_key():
    assert workflow_module._actual_dimensions_barcode_parts("2 Game") == workflow_module.ActualDimensionsBarcodeParts(
        cost_summary_vfi="2 Game",
        group_vfi_key="2 Game",
        is_charge_row=True,
    )
    assert workflow_module._actual_dimensions_barcode_parts(
        "2 1 of 2 Game"
    ) == workflow_module.ActualDimensionsBarcodeParts(
        cost_summary_vfi="2 Game",
        group_vfi_key="2 Game",
        is_charge_row=True,
    )
    assert workflow_module._actual_dimensions_barcode_parts(
        "2 2 of 2 Game"
    ) == workflow_module.ActualDimensionsBarcodeParts(
        cost_summary_vfi="",
        group_vfi_key="2 Game",
        is_charge_row=False,
    )
    assert workflow_module._actual_dimensions_barcode_parts(
        "15 1 of 4 ITFFKS1"
    ) == workflow_module.ActualDimensionsBarcodeParts(
        cost_summary_vfi="15 ITFFKS1",
        group_vfi_key="15 ITFFKS1",
        is_charge_row=True,
    )
    assert workflow_module._actual_dimensions_barcode_parts(
        "15 2 of 4 ITFFKS1"
    ) == workflow_module.ActualDimensionsBarcodeParts(
        cost_summary_vfi="",
        group_vfi_key="15 ITFFKS1",
        is_charge_row=False,
    )
    assert workflow_module._actual_dimensions_barcode_parts(
        "15 4 of 4 ITFFKS1"
    ) == workflow_module.ActualDimensionsBarcodeParts(
        cost_summary_vfi="",
        group_vfi_key="15 ITFFKS1",
        is_charge_row=False,
    )
    assert workflow_module._actual_dimensions_barcode_parts(
        "One Page Rules-1"
    ) == workflow_module.ActualDimensionsBarcodeParts(
        cost_summary_vfi="One Page Rules-1",
        group_vfi_key="One Page Rules-1",
        is_charge_row=True,
    )
    assert workflow_module._actual_dimensions_barcode_parts(
        "One Page Rules-1-2"
    ) == workflow_module.ActualDimensionsBarcodeParts(
        cost_summary_vfi="",
        group_vfi_key="One Page Rules-1",
        is_charge_row=False,
    )


def test_actual_dimensions_rows_match_small_expected_scan_count_without_default_cap():
    expected = [f"PKG-{index}" for index in range(1, 21)]

    rows = workflow_module._actual_dimensions_rows(expected_scan_barcodes=expected)

    assert not hasattr(workflow_module, "ACTUAL_DIMENSIONS_FORMULA_ROWS")
    assert workflow_module._actual_dimensions_rows() == []
    assert len(rows) == 20
    assert [row["Expected scan barcode"] for row in rows] == expected
    assert rows[-1]["Expected scan group VFI key"].formula.startswith('IF($L21="","",')
    assert rows[-1]["Scan status"].formula == (
        'IF($L21="","",IF(SUMPRODUCT(--(TRIM($A$2:$A$21)=TRIM($L21)))>0,"","Not Scanned"))'
    )


def test_actual_dimensions_rows_fill_beyond_one_thousand_expected_packages():
    expected = [f"PKG-{index}" for index in range(1, 1006)]

    rows = workflow_module._actual_dimensions_rows(expected_scan_barcodes=expected)

    assert len(rows) == 1005
    assert rows[-1]["Expected scan barcode"] == "PKG-1005"
    assert rows[-1]["Expected scan group VFI key"].formula.startswith('IF($L1006="","",')
    assert "MAX($B1006/1000,$F1006)" in rows[-1]["Actual chargeable weight kg"].formula


def test_actual_dimensions_rows_fill_through_ten_thousand_expected_packages():
    expected = [f"PKG-{index}" for index in range(1, 10001)]

    rows = workflow_module._actual_dimensions_rows(expected_scan_barcodes=expected)

    assert len(rows) == 10000
    assert rows[-1]["Expected scan barcode"] == "PKG-10000"
    assert rows[-1]["Expected scan group VFI key"].formula.startswith('IF($L10001="","",')
    assert "SUMPRODUCT(--(TRIM($A$2:$A$10001)=TRIM($L10001)))" in rows[-1]["Scan status"].formula


def test_cost_summary_scan_note_uses_full_actual_dimensions_expected_helper_columns():
    rate_sheet = workflow_module.CustomerRateSheet(
        hub=workflow_module._empty_rate_lane(),
        express=workflow_module._empty_rate_lane(),
    )
    rows = workflow_module._cost_summary_rows(
        [{"VFI #": "PKG-10000", "Country": "US"}],
        {},
        workflow_module.CustomerRateSheetSelection(sheet=rate_sheet, path="", filename="", source="test"),
        {},
    )

    formula = rows[0]["Scan note"].formula

    assert "COUNTIFS('Actual Dimensions'!$AE:$AE,$B2" in formula
    assert "'Actual Dimensions'!$M:$M,\"Not Scanned\"" in formula
    assert "$AE$2:$AE$1001" not in formula
    assert "$M$2:$M$1001" not in formula


def test_cost_summary_scan_note_uses_package_set_status_for_multi_package_group():
    rate_sheet = workflow_module.CustomerRateSheet(
        hub=workflow_module._empty_rate_lane(),
        express=workflow_module._empty_rate_lane(),
    )
    rows = workflow_module._cost_summary_rows(
        [{"VFI #": "15 ITFFKS1", "Country": "US"}],
        {},
        workflow_module.CustomerRateSheetSelection(sheet=rate_sheet, path="", filename="", source="test"),
        {},
    )
    actual_rows = workflow_module._actual_dimensions_rows(
        expected_scan_barcodes=["15 1 of 2 ITFFKS1", "15 2 of 2 ITFFKS1"]
    )

    formula = rows[0]["Scan note"].formula

    assert 'COUNTIFS(\'Actual Dimensions\'!$AE:$AE,$B2' in formula
    assert '\'Actual Dimensions\'!$M:$M,"Not Scanned"' in formula
    assert ',"Item not scanned","")' in formula
    assert 'ISNUMBER(SEARCH(" of ",$L2))' in actual_rows[0]["Expected scan group VFI key"].formula
    assert 'ISNUMBER(SEARCH(" of ",$L3))' in actual_rows[1]["Expected scan group VFI key"].formula
    assert "TRIM($L2)" in actual_rows[0]["Scan status"].formula
    assert "TRIM($L3)" in actual_rows[1]["Scan status"].formula


def test_cost_summary_scan_note_keeps_blank_until_scans_are_present():
    rate_sheet = workflow_module.CustomerRateSheet(
        hub=workflow_module._empty_rate_lane(),
        express=workflow_module._empty_rate_lane(),
    )
    rows = workflow_module._cost_summary_rows(
        [{"VFI #": "15 ITFFKS1", "Country": "US"}],
        {},
        workflow_module.CustomerRateSheetSelection(sheet=rate_sheet, path="", filename="", source="test"),
        {},
    )

    formula = rows[0]["Scan note"].formula

    assert formula.startswith('IF(COUNTA(\'Actual Dimensions\'!$A:$A)<=1,"",')
    assert "Item not scanned" in formula


def test_workbook_actual_dimensions_writes_expected_rows_beyond_one_thousand(tmp_path):
    expected = [f"PKG-{index}" for index in range(1, 1006)]
    output_path = tmp_path / "actual_dimensions_extent.xlsx"

    excel_writer_module.write_workbook(
        str(output_path),
        actual_dimensions_rows=workflow_module._actual_dimensions_rows(expected_scan_barcodes=expected),
    )

    actual_xml = _sheet_xml(output_path, "Actual Dimensions")
    actual_root = ElementTree.fromstring(actual_xml)

    assert _inline_cell_text(actual_xml, "L1006") == "PKG-1005"
    assert _cell_formula(actual_xml, "AE1006").startswith('IF($L1006="","",')
    assert _cell_formula(actual_xml, "M1006") == (
        'IF($L1006="","",IF(SUMPRODUCT(--(TRIM($A$2:$A$1006)=TRIM($L1006)))>0,"","Not Scanned"))'
    )
    assert actual_root.find(".//main:row[@r='1007']", _NS) is None


def test_output_workbook_includes_received_intake_column_and_sku_intake_summary(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized_shipping_plan.xlsx"
    _write_csv(
        sku_master_path,
        [
            {
                "SKU": "CORE",
                "Product Name": "Core Game",
                "Length": "5",
                "Width": "5",
                "Height": "5",
                "Weight kg": "1",
                "quantity": "10",
            },
            {
                "SKU": "EXTRA",
                "Product Name": "Extra Item",
                "Length": "2",
                "Width": "2",
                "Height": "2",
                "Weight kg": "0.1",
                "quantity": "6",
            },
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "1001", "SKU": "CORE", "Quantity": "2", "Country": "US"},
            {"Order ID": "1002", "SKU": "CORE", "Quantity": "3", "Country": "US"},
            {"Order ID": "1003", "SKU": "ORDERONLY", "Quantity": "4", "Country": "US"},
        ],
    )

    optimize_workbook(
        sku_master_path=str(sku_master_path),
        orders_path=str(orders_path),
        output_path=str(output_path),
        config={},
    )

    intake_rows = _sheet_rows(output_path, "VFI Intake Form")
    assert intake_rows[0]["Received"] == "10"
    assert intake_rows[1]["Received"] == "6"

    summary_rows = _sheet_rows(output_path, "Summary")
    assert any(row["SKU"] == "SKU Intake Summary" for row in summary_rows)
    summary_xml = _sheet_xml(output_path, "Summary")
    assert _inline_cell_text(summary_xml, "E2") == "SKU Intake Summary"
    assert _inline_cell_text(summary_xml, "E3") == "SKU"
    assert _inline_cell_text(summary_xml, "F3") == "Received Quantity"
    assert _inline_cell_text(summary_xml, "G3") == "Required Quantity"
    assert _inline_cell_text(summary_xml, "H3") == "Remaining"
    summary_by_sku = {row["SKU"]: row for row in summary_rows if row.get("SKU") in {"CORE", "EXTRA", "ORDERONLY"}}

    assert int(summary_by_sku["CORE"]["Received Quantity"]) == 10
    assert int(summary_by_sku["CORE"]["Required Quantity"]) == 5
    assert int(summary_by_sku["CORE"]["Remaining"]) == 5
    assert int(summary_by_sku["EXTRA"]["Received Quantity"]) == 6
    assert int(summary_by_sku["EXTRA"]["Required Quantity"]) == 0
    assert int(summary_by_sku["EXTRA"]["Remaining"]) == 6
    assert int(summary_by_sku["ORDERONLY"]["Received Quantity"]) == 0
    assert int(summary_by_sku["ORDERONLY"]["Required Quantity"]) == 4
    assert int(summary_by_sku["ORDERONLY"]["Remaining"]) == -4


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
            {"Order ID": "OK-1", "SKU": "Good", "Quantity": "1", "Country": "", "Backer ID": ""},
            {"Order ID": "BAD-1", "SKU": "Huge", "Quantity": "1", "Country": "Germany", "Backer ID": "B-BAD"},
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
    oversized = next(row for row in warning_rows if row["Order ID"] == "BAD-1" and row["Error Type"] == "OversizedItem")
    assert oversized["Severity"] == "Error"
    assert oversized["SKU"] == "HUGE"
    assert oversized["Backer ID"] == "B-BAD"
    assert oversized["Country"] == "Germany"
    assert "74 x 37 x 44 cm" in oversized["Message"]
    assert "Carton cap: 74 x 37 x 44 cm" in oversized["Box/Fit Context"]

    summary_rows = next(sheet.rows for sheet in workbook_rows if sheet.sheet_name == "Summary")
    summary_warning = next(
        row for row in summary_rows if row["Section"] == "Packing Errors" and row.get("SKU") == "HUGE"
    )
    assert summary_warning["Severity"] == "Error"
    assert summary_warning["SKU"] == "HUGE"
    assert summary_warning["Order ID"] == "BAD-1"
    assert summary_warning["Backer ID"] == "B-BAD"
    assert summary_warning["Country"] == "Germany"
    assert "74 x 37 x 44 cm" in summary_warning["Reason / Message"]


def test_summary_shows_errors_only_and_keeps_nonblocking_warnings_off_front_sheet():
    rows = workflow_module._clean_summary_rows(
        {
            "orders_processed": 1,
            "boxes_created": 1,
            "box_types": 1,
            "unmatched_skus": 0,
        },
        [],
        [],
        {},
        errors_and_warnings_rows=[
            {
                "Severity": "Warning",
                "Order ID": "WARN-1",
                "Backer ID": "B-WARN",
                "Country": "France",
                "SKU": "CORE",
                "Stage": "report",
                "Error Type": "RetailBulkReview",
                "Message": "Non-blocking review warning.",
                "Box/Fit Context": "VB 1",
            },
            {
                "Severity": "Error",
                "Order ID": "ERR-1",
                "Backer ID": "B-ERR",
                "Country": "Germany",
                "SKU": "HUGE",
                "Stage": "packing",
                "Error Type": "OversizedItem",
                "Message": "SKU HUGE could not be packed: exceeds carton cap.",
                "Box/Fit Context": "Carton cap: 74 x 37 x 44 cm",
            },
        ],
    )

    error_rows = [row for row in rows if row["Section"] == "Packing Errors"]
    assert len(error_rows) == 1
    assert error_rows[0]["Severity"] == "Error"
    assert error_rows[0]["SKU"] == "HUGE"
    assert error_rows[0]["Reason / Message"] == "SKU HUGE could not be packed: exceeds carton cap."
    assert error_rows[0]["Order ID"] == "ERR-1"
    assert error_rows[0]["Backer ID"] == "B-ERR"
    assert error_rows[0]["Country"] == "Germany"
    assert error_rows[0]["Box/Fit Context"] == "Carton cap: 74 x 37 x 44 cm"
    assert all(row.get("SKU") != "CORE" for row in rows if row["Section"] == "Packing Errors")


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
    multi_box_rows = next(
        sheet.rows for sheet in workbook_rows if sheet.sheet_name == "Multi Box Detail"
    )
    for row in multi_box_rows:
        assert float(row["Length cm"]) >= 74
        assert row["Vendor Box ID"]
    warning_sheet = next(
        (sheet for sheet in workbook_rows if sheet.sheet_name == "Errors and Warnings"),
        None,
    )
    if warning_sheet:
        warning_keys = [
            (
                row["Order ID"],
                row["SKU"],
                row["Stage"],
                row["Error Type"],
                row["Message"],
            )
            for row in warning_sheet.rows
        ]
        assert len(warning_keys) == len(set(warning_keys))


def test_vendor_box_height_is_cut_down_to_packed_height_plus_two(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized_shipping_plan.xlsx"
    _write_csv(
        sku_master_path,
        [
            {
                "SKU": "Wide Tall",
                "Product Name": "Wide Tall",
                "Length": "43",
                "Width": "37",
                "Height": "30",
                "Weight kg": "1",
            }
        ],
    )
    _write_csv(
        orders_path,
        [{"Order ID": "CUT-1", "SKU": "Wide Tall", "Quantity": "1"}],
    )

    optimize_workbook(
        sku_master_path=str(sku_master_path),
        orders_path=str(orders_path),
        output_path=str(output_path),
        config={
            "packing_mode": "fast",
            "preserve_region_sheets": False,
            "sku_rules": {"Wide Tall": {"no_padding": True}},
        },
    )

    workbook_rows = read_workbook(str(output_path))
    multi_box_rows = next(
        sheet.rows for sheet in workbook_rows if sheet.sheet_name == "Multi Box Detail"
    )
    detail = multi_box_rows[0]
    assert detail["Vendor Box ID"]
    assert detail["Box Type"].startswith("VB ")
    assert "cutdown" in detail["Box Type"]
    assert float(detail["Height cm"]) == 32
    assert "Vendor box height cut down to 32 cm." in detail["Box Standardization Note"]

    order_rows = next(
        sheet.rows for sheet in workbook_rows if sheet.sheet_name == "Order Volume Weights"
    )
    assert "cutdown" in order_rows[0]["Box Plan"]


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
    detail_row = _sheet_rows(output_path, "Multi Box Detail")[0]
    assert "PREPACK-CORE-BOX" in row["Box Plan"]
    assert detail_row["Box Type"] == "PREPACK-CORE-BOX"
    assert float(detail_row["Length cm"]) == 31
    assert float(detail_row["Width cm"]) == 22
    assert float(detail_row["Height cm"]) == 8


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


def test_prepacked_final_ship_alone_quantity_creates_one_carton_per_unit(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {
                "SKU": "Ready",
                "Product Name": "Ready",
                "Length": "64",
                "Width": "32",
                "Height": "32",
                "Weight lb": "17.5",
            },
        ],
    )
    _write_csv(
        orders_path,
        [{"Order ID": "1", "SKU": "Ready", "Quantity": "2"}],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={
            "packing_mode": "fast",
            "sku_rules": {
                "Ready": {
                    "prepacked": True,
                    "no_padding": True,
                    "ships_alone": True,
                    "can_mix_with_other_items": False,
                    "box_type": "READY-SKU-CARTON",
                }
            },
            "preserve_region_sheets": False,
        },
    )

    detail_rows = [
        row
        for row in _sheet_rows(output_path, "Multi Box Detail")
        if row["Box Type"] == "READY-SKU-CARTON"
    ]
    assert len(detail_rows) == 2
    assert all(row["SKUs in Box"] == "READY x1" for row in detail_rows)
    assert all(float(row["Chargeable Weight kg"]) == 13.1 for row in detail_rows)

    summary_row = next(
        row
        for row in _sheet_rows(output_path, "Box Size Summary")
        if row["Box Type"] == "READY-SKU-CARTON"
    )
    assert int(float(summary_row["Box Count"])) == 2
    assert int(float(summary_row["Unit Count"])) == 2
    assert float(summary_row["Max Chargeable Weight kg"]) == 13.1


def test_separate_playmat_charge_bucket_forces_ship_as_is_behavior(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {"SKU": "PLAYMAT", "Product Name": "Playmat", "Length": "60", "Width": "35", "Height": "2", "Weight kg": "0.5"},
            {"SKU": "ADDON", "Product Name": "Addon", "Length": "8", "Width": "5", "Height": "2", "Weight kg": "0.2"},
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "1", "SKU": "PLAYMAT", "Quantity": "2"},
            {"Order ID": "1", "SKU": "ADDON", "Quantity": "1"},
        ],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={
            "packing_mode": "fast",
            "separate_playmat_charge_skus": ["PLAYMAT"],
            "sku_rules": {"PLAYMAT": {"can_mix_with_other_items": True, "no_padding": False}},
            "preserve_region_sheets": False,
        },
    )

    detail_rows = _sheet_rows(output_path, "Multi Box Detail")
    playmat_rows = [row for row in detail_rows if row["SKUs in Box"] == "PLAYMAT x1"]
    addon_rows = [row for row in detail_rows if row["SKUs in Box"] == "ADDON x1"]
    assert len(playmat_rows) == 2
    assert len(addon_rows) == 1
    assert all(row["Vendor Box ID"] == "" for row in playmat_rows)
    assert all(row["Box Selection Decision"] == "rule_assigned_box" for row in playmat_rows)
    assert all(float(row["Length cm"]) == 60 for row in playmat_rows)
    assert all(float(row["Width cm"]) == 35 for row in playmat_rows)
    assert all(float(row["Height cm"]) == 2 for row in playmat_rows)
    assert all("separate playmat charge" in row["Box Standardization Note"].lower() for row in playmat_rows)


def test_playmat_name_without_config_bucket_does_not_get_special_rule(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [{"SKU": "PLAYMAT", "Product Name": "Playmat", "Length": "20", "Width": "10", "Height": "2", "Weight kg": "0.5"}],
    )
    _write_csv(orders_path, [{"Order ID": "1", "SKU": "PLAYMAT", "Quantity": "1"}])

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={"packing_mode": "fast", "preserve_region_sheets": False, "use_vendor_box_menu": False},
    )

    detail_row = _sheet_rows(output_path, "Multi Box Detail")[0]
    cost_row = _sheet_rows(output_path, "Cost Summary")[0]
    assert detail_row["Rule Applied"] == ""
    assert "Separate Playmat Charge" not in cost_row.get("Shipping Rate Note", "")


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
        assert "Box Type" not in row
        assert "Box Standardization Note" not in row
        assert "Box 1:" in row["Box Plan"]
        assert all(detail["Box Type"] in row["Box Plan"] for detail in matching_detail)


def test_workbook_presentation_tabs_and_compact_columns(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {"SKU": "All-in Storage Solution [72570]", "Product Name": "All-in Storage Solution [72570]", "Length": "20", "Width": "10", "Height": "5", "Weight kg": "1"},
            {"SKU": "Core", "Product Name": "Core", "Length": "10", "Width": "8", "Height": "3", "Weight kg": "0.5"},
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "1", "SKU": "All-in Storage Solution [72570]", "Quantity": "1", "Country": "USA", "State": "Armed Forces AE", "Name": "Alice"},
            {"Order ID": "1", "SKU": "Core", "Quantity": "1", "Country": "USA", "State": "Armed Forces AE", "Name": "Alice"},
            {"Order ID": "2", "SKU": "Core", "Quantity": "1", "Country": "Republic of Korea", "State": "", "Name": "Bob"},
            {"Order ID": "2", "SKU": "All-in Storage Solution [72570]", "Quantity": "1", "Country": "Republic of Korea", "State": "", "Name": "Bob"},
        ],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={
            "packing_mode": "fast",
            "preserve_region_sheets": False,
            "sku_rules": {
                "All-in Storage Solution [72570]": {
                    "prepacked": True,
                    "no_padding": True,
                    "ships_alone": True,
                    "can_mix_with_other_items": False,
                    "box_type": "All-in Storage Solution carton",
                }
            },
        },
    )

    workbook = read_workbook(str(output_path))
    sheet_names = _workbook_sheet_names(output_path)
    assert sheet_names[:4] == ["Summary", "Cost Summary", "Actual Dimensions", "Labels"]
    for required_sheet in [
        "VFI Intake Form",
        "Optimized to Pack",
        "Label generator",
        "Order Volume Weights",
        "Box Size Summary",
    ]:
        assert required_sheet in sheet_names

    summary_rows = _sheet_rows(output_path, "Summary")
    assert any(row["Section"] == "Boxes Needed" and row["Metric"].startswith("VB ") for row in summary_rows)
    assert any(row["Section"] == "Rules Applied Summary" and "ship-alone" in row["Value"] for row in summary_rows)
    assert all(row["Section"] != "Unique Warning Summary" for row in summary_rows)
    debug_summary_rows = _sheet_rows(output_path, "Debug Summary")
    assert any(row["Section"] == "Rules Applied Summary" and "ship-alone" in row["Value"] for row in debug_summary_rows)
    assert any(row["Section"] == "Unique Warning Summary" for row in debug_summary_rows)

    order_rows = {row["Order ID"]: row for row in _sheet_rows(output_path, "Order Volume Weights")}
    headers = list(next(iter(order_rows.values())))
    assert order_rows["1"]["Country"] == "United States"
    assert order_rows["1"]["US State Abbreviation"] == "AE"
    assert order_rows["2"]["Country"] == "South Korea"
    assert "Box Type" not in headers
    assert "All-in Storage Solution carton" in order_rows["1"]["Box Plan"]
    assert "VB " in order_rows["1"]["Box Plan"]
    assert "Box Plan" in headers
    assert headers.index("Box Plan") == headers.index("Box Qty") + 1
    assert headers.index("Per-Box Chargeable Weight") == headers.index("Box Plan") + 1
    assert "All-in Storage Solution carton:" in order_rows["1"]["Per-Box Chargeable Weight"]
    assert "VB " in order_rows["1"]["Per-Box Chargeable Weight"]
    for removed in [
        "Box Type",
        "Assigned Box Length cm",
        "Assigned Box Width cm",
        "Assigned Box Height cm",
        "Box Standardization Note",
        "Distinct SKUs",
        "Warning Summary",
        "Vendor Box ID",
    ]:
        assert removed not in headers
    assert headers.index("Name") > headers.index("SKU Breakdown")

    box_size_rows = _sheet_rows(output_path, "Box Size Summary")
    assert all(not row["Box Type"].startswith("Vendor Box") for row in box_size_rows)
    assert any(row["Box Type"].startswith("VB ") for row in box_size_rows)
    assert "Main SKU Combos" not in box_size_rows[0]

    optimized_rows = _sheet_rows(output_path, "Optimized to Pack")
    assert optimized_rows[0]["Pledge Configuration"] == "1"
    assert int(float(optimized_rows[0]["Total Pledges"])) == 2
    assert "ALL-IN STORAGE SOLUTION [72570] x1" in optimized_rows[0]["All Items"]
    assert "All-in Storage Solution carton:" in optimized_rows[0]["Box 1"]
    assert "VB " in optimized_rows[0]["Box 2"]
    assert "CORE x1" in optimized_rows[0]["Box 2"]




def test_small_split_carton_uses_smaller_vendor_box_across_workbook_tabs(tmp_path, monkeypatch):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {"SKU": "Large", "Product Name": "Large", "Length": "60", "Width": "30", "Height": "30", "Weight kg": "2"},
            {"SKU": "Small Addons", "Product Name": "Small Addons", "Length": "20", "Width": "10", "Height": "4", "Weight kg": "0.5"},
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "1", "SKU": "Large", "Quantity": "1"},
            {"Order ID": "1", "SKU": "Small Addons", "Quantity": "1"},
        ],
    )

    def fake_split_order_into_cartons(items, packing_mode="normal", force_simple_split=False, **kwargs):
        by_sku = {item.canonical_sku: item for item in items}
        large = by_sku["LARGE"]
        small = by_sku["SMALL ADDONS"]
        return SplitResult(
            success=True,
            box_qty=2,
            cartons=[
                SplitCarton(
                    box_number=1,
                    result=OptimizedCartonResult(
                        success=True,
                        length_cm=72,
                        width_cm=34,
                        height_cm=40,
                        chargeable_weight_kg=20,
                        volume_cm3=72 * 34 * 40,
                        placements=[Placement("LARGE", 1, large.padded_dimensions, (0, 0, 0), large.weight_kg)],
                        unplaced_items=[],
                    ),
                ),
                SplitCarton(
                    box_number=2,
                    result=OptimizedCartonResult(
                        success=True,
                        length_cm=34,
                        width_cm=24,
                        height_cm=12,
                        chargeable_weight_kg=3,
                        volume_cm3=34 * 24 * 12,
                        placements=[Placement("SMALL ADDONS", 1, small.padded_dimensions, (0, 0, 0), small.weight_kg)],
                        unplaced_items=[],
                    ),
                ),
            ],
            unplaced_items=[],
        )

    monkeypatch.setattr("box_optimizer.workflow.split_order_into_cartons", fake_split_order_into_cartons)

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={
            "packing_mode": "fast",
            "preserve_region_sheets": False,
            "chargeable_weight_split_savings_threshold_kg": 999,
        },
    )

    order_row = _sheet_rows(output_path, "Order Volume Weights")[0]
    assert "Box Type" not in order_row
    assert "VB 36" in order_row["Box Plan"]
    assert "VB 3 cutdown" in order_row["Box Plan"]
    assert "VB 3 cutdown:" in order_row["Per-Box Chargeable Weight"]

    detail_types = {row["Box Type"] for row in _sheet_rows(output_path, "Multi Box Detail")}
    assert detail_types == {"VB 36 cutdown", "VB 3 cutdown"}

    box_summary_types = {row["Box Type"] for row in _sheet_rows(output_path, "Box Size Summary")}
    assert {"VB 36 cutdown", "VB 3 cutdown"}.issubset(box_summary_types)

    optimized_rows = _sheet_rows(output_path, "Optimized to Pack")
    joined_boxes = " | ".join(value for key, value in optimized_rows[0].items() if key.startswith("Box "))
    assert "VB 36 cutdown:" in joined_boxes
    assert "VB 3 cutdown:" in joined_boxes

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
    assert "Average Chargeable Weight kg" not in rows[0]
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
    assert rows[0]["SKU Breakdown"] == "CORE x1 | ADDON x1"


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
    assert len(order_rows) == 2
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

    rows = _sheet_rows(output_path, "Multi Box Detail")
    by_order = {row["Order ID"]: row for row in rows}
    assert float(by_order["1"]["Length cm"]) == 35
    assert float(by_order["2"]["Length cm"]) == 36
    for sheet_name in ["Multi Box Detail", "Box Size Summary", "Pledge Combination Summary"]:
        for row in _sheet_rows(output_path, sheet_name):
            for column in [key for key in row if key.endswith("Length cm") or key.endswith("Width cm") or key.endswith("Height cm")]:
                assert float(row[column]).is_integer()
                assert "." not in row[column]


def test_single_box_order_packed_actual_weight_totals_all_skus(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {"SKU": "Core", "Product Name": "Core", "Length": "5", "Width": "5", "Height": "5", "Weight kg": "1.04"},
            {"SKU": "Addon", "Product Name": "Addon", "Length": "4", "Width": "4", "Height": "2", "Weight kg": "0.36"},
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
        config={"preserve_region_sheets": False},
    )

    order_row = _sheet_rows(output_path, "Order Volume Weights")[0]
    detail_rows = _sheet_rows(output_path, "Multi Box Detail")
    expected = format_kg_display(packed_actual_weight_kg(1.04 + 0.36))
    assert len(detail_rows) == 1
    assert float(order_row["Packed Actual Weight kg"]) == expected
    assert float(order_row["Packed Actual Weight kg"]) == float(detail_rows[0]["Packed Actual Weight kg"])
    assert float(order_row["Packed Actual Weight kg"]) not in {
        format_kg_display(packed_actual_weight_kg(1.04)),
        format_kg_display(packed_actual_weight_kg(0.36)),
    }


def test_multi_box_order_packed_actual_weight_sums_raw_carton_weights(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {"SKU": "Solo", "Product Name": "Solo", "Length": "8", "Width": "6", "Height": "4", "Weight kg": "1.04"},
            {"SKU": "Addon", "Product Name": "Addon", "Length": "7", "Width": "5", "Height": "3", "Weight kg": "1.04"},
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "1", "SKU": "Solo", "Quantity": "1"},
            {"Order ID": "1", "SKU": "Addon", "Quantity": "1"},
        ],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={
            "sku_rules": {"Solo": {"can_mix_with_other_items": False}},
            "preserve_region_sheets": False,
        },
    )

    order_row = _sheet_rows(output_path, "Order Volume Weights")[0]
    detail_rows = _sheet_rows(output_path, "Multi Box Detail")
    expected = format_kg_display(packed_actual_weight_kg(1.04 + 1.04))
    displayed_detail_sum = sum(float(row["Packed Actual Weight kg"]) for row in detail_rows)
    assert len(detail_rows) == 2
    assert float(order_row["Packed Actual Weight kg"]) == expected
    assert float(order_row["Packed Actual Weight kg"]) >= displayed_detail_sum
    assert float(order_row["Packed Actual Weight kg"]) - displayed_detail_sum <= 0.2


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
    assert headers.index("Backer Number") > headers.index("SKU Breakdown")


def test_country_tab_uses_raw_order_volume_metadata_and_preserves_nonstandard_headers(tmp_path):
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
                "Order ID": "1",
                "SKU": "Core",
                "Quantity": "1",
                "Country": "Hong Kong",
                "Backer Number": "BK-7",
                "Name": "Ada Lovelace",
                "Email": "ada@example.com",
                "Address 1": "1 Algorithm Ave",
                "Address 2": "Unit 9",
                "Add to": "Apartment call box 7",
                "Tax ID number": "HK-TAX",
                "Shipping Method": "Courier",
            }
        ],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={
            "preserve_region_sheets": False,
            "campaign": {"name": "Long Human Campaign", "code": "BARCODE"},
        },
    )

    order_row = _sheet_rows(output_path, "Order Volume Weights")[0]
    country_row = _sheet_rows(output_path, "China-HK")[0]
    metadata_headers = [
        "Backer Number",
        "Name",
        "Email",
        "Address 1",
        "Address 2",
        "Add to",
        "Tax ID number",
        "Shipping Method",
    ]
    assert country_row["SKU"] == order_row["SKU"]
    assert country_row["Quantity"] == order_row["Quantity"]
    assert country_row["Items in box"] == "1"
    assert country_row["Campaign"] == "Long Human Campaign"
    assert country_row["Campaign"] != "BARCODE"
    for header in metadata_headers:
        assert country_row[header] == order_row[header]
    assert "Original Optimized Box Type" not in country_row
    assert "Shipping name" not in country_row
    assert "Address Line 1" not in country_row
    assert "Address Line 2" not in country_row

    hong_kong_xml = _sheet_xml(output_path, "China-HK")
    assert [_inline_cell_text(hong_kong_xml, cell) for cell in ["A1", "B1", "C1", "D1", "E1", "N1", "O1", "P1", "Q1"]] == [
        "Campaign",
        "VFI #",
        "Actual weight g",
        "Volumetric weight kg",
        "SKU",
        "Shipping Method",
        "",
        "Items in box",
        "Items in this box / SKU contents",
    ]
    assert _inline_cell_text(hong_kong_xml, "M2") == "HK-TAX"
    assert _inline_cell_text(hong_kong_xml, "O2") == ""
    assert _inline_cell_text(hong_kong_xml, "Q2") == "(1)  CORE x1"


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
    assert "Assigned Box Length cm" not in row
    assert "Box Standardization Note" not in row
    assert "Distinct SKUs" not in row
    assert "Warning Summary" not in row
    assert "Vendor Box ID" not in row
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
    assert list(row).index("Backer Number") > list(row).index("SKU Breakdown")


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
    assert headers.index("Backer Number") > headers.index("SKU Breakdown")


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
        config={
            "packing_mode": "fast",
            "preserve_region_sheets": False,
            "chargeable_weight_split_savings_threshold_kg": 999,
        },
    )

    row = _sheet_rows(output_path, "Order Volume Weights")[0]
    assert int(float(row["Box Qty"])) == 1
    detail = _sheet_rows(output_path, "Multi Box Detail")
    assert len(detail) == 1
    assert "LARGE x1" in detail[0]["SKUs in Box"]
    assert "SMALL A x1" in detail[0]["SKUs in Box"]
    debug_summary_rows = _sheet_rows(output_path, "Debug Summary")
    assert any(
        summary["Metric"] == "Chargeable Weight Plans Selected" and int(float(summary["Value"])) == 0
        for summary in debug_summary_rows
    )


def test_chargeable_weight_plan_keeps_one_box_when_savings_below_operational_threshold(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {"SKU": "Large", "Product Name": "Large", "Length": "40", "Width": "30", "Height": "15", "Weight kg": "2"},
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
        config={
            "packing_mode": "fast",
            "preserve_region_sheets": False,
            "non_preferred_extra_box_savings_threshold_kg": 0,
            "non_preferred_extra_box_savings_threshold_pct": 0,
        },
    )

    row = _sheet_rows(output_path, "Order Volume Weights")[0]
    assert int(float(row["Box Qty"])) == 1

    debug_summary_rows = _sheet_rows(output_path, "Debug Summary")
    assert any(
        summary["Metric"] == "Chargeable Weight Plans Selected" and int(float(summary["Value"])) == 0
        for summary in debug_summary_rows
    )
    sheet_names = [sheet.sheet_name for sheet in read_workbook(str(output_path))]
    assert "Errors and Warnings" not in sheet_names


def test_company_protection_guardrail_rejects_extra_box_when_margin_gets_worse(monkeypatch):
    lines = [
        OrderLine(order_id="1", raw_sku="Large", canonical_sku="Large", quantity=1, country="US", state_province="CA"),
        OrderLine(order_id="1", raw_sku="Small A", canonical_sku="Small A", quantity=1, country="US", state_province="CA"),
        OrderLine(order_id="1", raw_sku="Small B", canonical_sku="Small B", quantity=1, country="US", state_province="CA"),
    ]
    context = workflow_module.PackingOrderContext(
        order_id="1",
        lines=lines,
        combo="Large x1 | Small A x1 | Small B x1",
        items=[],
        groups=[workflow_module.RuleSplitGroup(lines=lines)],
        cache_key="key",
        first_index=0,
    )
    sku_items = {
        "Large": _sku_item("Large", Dimensions(40, 30, 15), 2),
        "Small A": _sku_item("Small A", Dimensions(6, 5, 2), 0.2),
        "Small B": _sku_item("Small B", Dimensions(5, 5, 2), 0.2),
    }

    def fake_pack_group_records(**kwargs):
        box_qty = len(kwargs["groups"])
        return SplitResult(True, box_qty, [SplitCarton(index + 1, None) for index in range(box_qty)], []), []

    def fake_score(**kwargs):
        box_qty = kwargs["split_result"].box_qty
        if box_qty == 1:
            return workflow_module.CandidatePlanScore(30.0, 1, 1, 30000, package_chargeable_weights_kg=(30.0,))
        return workflow_module.CandidatePlanScore(20.0, box_qty, box_qty, 20000, package_chargeable_weights_kg=(10.0, 10.0))

    monkeypatch.setattr(workflow_module, "_pack_group_records", fake_pack_group_records)
    monkeypatch.setattr(workflow_module, "_score_assigned_split_result", fake_score)

    split_result, warnings = workflow_module._select_chargeable_weight_plan(
        context=context,
        sku_items=sku_items,
        sku_rules={},
        packing_mode="fast",
        cfg={
            **workflow_module.DEFAULT_CONFIG,
            "non_preferred_extra_box_savings_threshold_kg": 0,
            "non_preferred_extra_box_savings_threshold_pct": 0,
            "company_protection_rate_bands": {
                "Zone USA": {"10": 30, "20": 40, "30": 60},
            },
        },
        remaining_budget_seconds=None,
    )

    assert split_result.box_qty == 1
    assert warnings == []


def test_company_protection_guardrail_allows_extra_box_when_margin_improves(monkeypatch):
    lines = [
        OrderLine(order_id="1", raw_sku="Large", canonical_sku="Large", quantity=1, country="US", state_province="CA"),
        OrderLine(order_id="1", raw_sku="Small A", canonical_sku="Small A", quantity=1, country="US", state_province="CA"),
        OrderLine(order_id="1", raw_sku="Small B", canonical_sku="Small B", quantity=1, country="US", state_province="CA"),
    ]
    context = workflow_module.PackingOrderContext(
        order_id="1",
        lines=lines,
        combo="Large x1 | Small A x1 | Small B x1",
        items=[],
        groups=[workflow_module.RuleSplitGroup(lines=lines)],
        cache_key="key",
        first_index=0,
    )
    sku_items = {
        "Large": _sku_item("Large", Dimensions(40, 30, 15), 2),
        "Small A": _sku_item("Small A", Dimensions(6, 5, 2), 0.2),
        "Small B": _sku_item("Small B", Dimensions(5, 5, 2), 0.2),
    }

    def fake_pack_group_records(**kwargs):
        box_qty = len(kwargs["groups"])
        return SplitResult(True, box_qty, [SplitCarton(index + 1, None) for index in range(box_qty)], []), []

    def fake_score(**kwargs):
        box_qty = kwargs["split_result"].box_qty
        if box_qty == 1:
            return workflow_module.CandidatePlanScore(30.0, 1, 1, 30000, package_chargeable_weights_kg=(30.0,))
        return workflow_module.CandidatePlanScore(20.0, box_qty, box_qty, 20000, package_chargeable_weights_kg=(10.0, 10.0))

    monkeypatch.setattr(workflow_module, "_pack_group_records", fake_pack_group_records)
    monkeypatch.setattr(workflow_module, "_score_assigned_split_result", fake_score)

    split_result, warnings = workflow_module._select_chargeable_weight_plan(
        context=context,
        sku_items=sku_items,
        sku_rules={},
        packing_mode="fast",
        cfg={
            **workflow_module.DEFAULT_CONFIG,
            "non_preferred_extra_box_savings_threshold_kg": 0,
            "non_preferred_extra_box_savings_threshold_pct": 0,
            "company_protection_rate_bands": {
                "Zone USA": {"10": 20, "20": 50, "30": 60},
            },
        },
        remaining_budget_seconds=None,
    )

    assert split_result.box_qty == 2
    assert warnings[0].error_type == "ChargeableWeightPlanSelected"


def test_repeat_retail_can_accept_small_margin_giveback_for_customer_savings():
    lines = [
        OrderLine(order_id="1", raw_sku="Large", canonical_sku="Large", quantity=1, country="US", state_province="CA")
    ]
    baseline = workflow_module.CandidatePlanScore(
        total_chargeable_weight_kg=30.0,
        box_qty=1,
        box_type_count=1,
        total_assigned_volume_cm3=30000,
        package_chargeable_weights_kg=(30.0,),
    )
    candidate = workflow_module.CandidatePlanScore(
        total_chargeable_weight_kg=20.0,
        box_qty=2,
        box_type_count=2,
        total_assigned_volume_cm3=20000,
        package_chargeable_weights_kg=(10.0, 10.0),
    )

    assert workflow_module._candidate_beats_baseline(
        candidate,
        baseline,
        1,
        0.5,
        0.05,
        2.0,
        1,
        2,
        1.0,
        0.075,
        3.0,
        0.10,
        lines,
        {
            **workflow_module.DEFAULT_CONFIG,
            "company_protection_rate_bands": {
                "Zone USA": {"10": 30, "20": 55, "30": 60},
            },
            "repeat_retail_max_margin_giveback": 5.1,
            "repeat_retail_min_customer_savings": 5.0,
        },
        "repeat retail accessory split Large",
    )


def test_extra_box_plan_requires_three_kg_savings_per_candidate_extra_box():
    baseline = workflow_module.CandidatePlanScore(
        total_chargeable_weight_kg=10.0,
        box_qty=1,
        box_type_count=1,
        total_assigned_volume_cm3=10000,
    )

    def beats(total_chargeable: float, box_qty: int) -> bool:
        return workflow_module._candidate_beats_baseline(
            workflow_module.CandidatePlanScore(
                total_chargeable_weight_kg=total_chargeable,
                box_qty=box_qty,
                box_type_count=box_qty,
                total_assigned_volume_cm3=box_qty * 1000,
            ),
            baseline,
            1,
            0.5,
            0.05,
            2.0,
            4,
            4,
            1.0,
            0.075,
            3.0,
            0.10,
            [],
            workflow_module.DEFAULT_CONFIG,
        )

    assert not beats(7.7, 2)
    assert beats(7.0, 2)
    assert not beats(6.0, 3)
    assert beats(3.0, 3)
    assert beats(1.2, 3)


def test_three_box_plan_compared_to_two_box_plan_requires_six_kg_savings():
    baseline = workflow_module.CandidatePlanScore(
        total_chargeable_weight_kg=12.0,
        box_qty=2,
        box_type_count=2,
        total_assigned_volume_cm3=12000,
    )

    def beats(total_chargeable: float) -> bool:
        return workflow_module._candidate_beats_baseline(
            workflow_module.CandidatePlanScore(
                total_chargeable_weight_kg=total_chargeable,
                box_qty=3,
                box_type_count=3,
                total_assigned_volume_cm3=3000,
            ),
            baseline,
            1,
            0.5,
            0.05,
            2.0,
            4,
            4,
            1.0,
            0.075,
            3.0,
            0.10,
            [],
            workflow_module.DEFAULT_CONFIG,
        )

    assert not beats(8.0)
    assert beats(5.0)


def test_broader_top_two_dim_candidate_can_win_when_two_extra_boxes_allowed(monkeypatch):
    lines = [
        _order_line("1", "LARGE A"),
        _order_line("1", "LARGE B"),
        _order_line("1", "SMALL A"),
        _order_line("1", "SMALL B"),
    ]
    context = workflow_module.PackingOrderContext(
        order_id="1",
        lines=lines,
        combo="LARGE A x1 | LARGE B x1 | SMALL A x1 | SMALL B x1",
        items=[],
        groups=[workflow_module.RuleSplitGroup(lines=lines)],
        cache_key="key",
        first_index=0,
    )
    sku_items = {
        "LARGE A": _sku_item("LARGE A", Dimensions(50, 30, 16), 2),
        "LARGE B": _sku_item("LARGE B", Dimensions(44, 28, 14), 1.5),
        "SMALL A": _sku_item("SMALL A", Dimensions(6, 5, 2), 0.2),
        "SMALL B": _sku_item("SMALL B", Dimensions(5, 5, 2), 0.2),
    }

    def fake_pack_group_records(**kwargs):
        box_qty = len(kwargs["groups"])
        return SplitResult(True, box_qty, [SplitCarton(index + 1, None) for index in range(box_qty)], []), []

    def fake_score(**kwargs):
        box_qty = kwargs["split_result"].box_qty
        chargeable_by_box_qty = {1: 12.0, 2: 11.2, 3: 5.0}
        return workflow_module.CandidatePlanScore(
            total_chargeable_weight_kg=chargeable_by_box_qty[box_qty],
            box_qty=box_qty,
            box_type_count=box_qty,
            total_assigned_volume_cm3=box_qty * 1000,
        )

    monkeypatch.setattr(workflow_module, "_pack_group_records", fake_pack_group_records)
    monkeypatch.setattr(workflow_module, "_score_assigned_split_result", fake_score)

    split_result, warnings = workflow_module._select_chargeable_weight_plan(
        context=context,
        sku_items=sku_items,
        sku_rules={},
        packing_mode="fast",
        cfg={
            **workflow_module.DEFAULT_CONFIG,
            "max_extra_boxes_per_order": 2,
        },
        remaining_budget_seconds=None,
    )

    assert split_result.box_qty == 3
    assert warnings[0].rule_applied.startswith("split top DIM")


def test_max_extra_box_guardrail_prevents_over_splitting(monkeypatch):
    lines = [
        _order_line("1", "LARGE A"),
        _order_line("1", "LARGE B"),
        _order_line("1", "SMALL A"),
        _order_line("1", "SMALL B"),
    ]
    context = workflow_module.PackingOrderContext(
        order_id="1",
        lines=lines,
        combo="LARGE A x1 | LARGE B x1 | SMALL A x1 | SMALL B x1",
        items=[],
        groups=[workflow_module.RuleSplitGroup(lines=lines)],
        cache_key="key",
        first_index=0,
    )
    sku_items = {
        "LARGE A": _sku_item("LARGE A", Dimensions(50, 30, 16), 2),
        "LARGE B": _sku_item("LARGE B", Dimensions(44, 28, 14), 1.5),
        "SMALL A": _sku_item("SMALL A", Dimensions(6, 5, 2), 0.2),
        "SMALL B": _sku_item("SMALL B", Dimensions(5, 5, 2), 0.2),
    }

    def fake_pack_group_records(**kwargs):
        box_qty = len(kwargs["groups"])
        return SplitResult(True, box_qty, [SplitCarton(index + 1, None) for index in range(box_qty)], []), []

    def fake_score(**kwargs):
        box_qty = kwargs["split_result"].box_qty
        chargeable_by_box_qty = {1: 12.0, 2: 11.8, 3: 9.5}
        return workflow_module.CandidatePlanScore(
            total_chargeable_weight_kg=chargeable_by_box_qty[box_qty],
            box_qty=box_qty,
            box_type_count=box_qty,
            total_assigned_volume_cm3=box_qty * 1000,
        )

    monkeypatch.setattr(workflow_module, "_pack_group_records", fake_pack_group_records)
    monkeypatch.setattr(workflow_module, "_score_assigned_split_result", fake_score)

    split_result, warnings = workflow_module._select_chargeable_weight_plan(
        context=context,
        sku_items=sku_items,
        sku_rules={},
        packing_mode="fast",
        cfg={
            **workflow_module.DEFAULT_CONFIG,
            "max_extra_boxes_per_order": 1,
        },
        remaining_budget_seconds=None,
    )

    assert split_result.box_qty == 1
    assert warnings == []


def test_repeat_retail_batch_candidate_can_use_more_small_cartons(monkeypatch):
    lines = [_order_line("1", "RETAIL BOX", quantity=80)]
    context = workflow_module.PackingOrderContext(
        order_id="1",
        lines=lines,
        combo="RETAIL BOX x80",
        items=[],
        groups=[workflow_module.RuleSplitGroup(lines=lines)],
        cache_key="key",
        first_index=0,
    )
    sku_items = {
        "RETAIL BOX": _sku_item("RETAIL BOX", Dimensions(25, 16, 15), 2),
    }

    def fake_pack_group_records(**kwargs):
        box_qty = len(kwargs["groups"])
        return SplitResult(True, box_qty, [SplitCarton(index + 1, None) for index in range(box_qty)], []), []

    def fake_score(**kwargs):
        box_qty = kwargs["split_result"].box_qty
        chargeable_by_box_qty = {1: 60.0, 2: 58.0, 8: 30.0}
        return workflow_module.CandidatePlanScore(
            total_chargeable_weight_kg=chargeable_by_box_qty.get(box_qty, 45.0),
            box_qty=box_qty,
            box_type_count=box_qty,
            total_assigned_volume_cm3=box_qty * 1000,
        )

    monkeypatch.setattr(workflow_module, "_pack_group_records", fake_pack_group_records)
    monkeypatch.setattr(workflow_module, "_score_assigned_split_result", fake_score)

    split_result, warnings = workflow_module._select_chargeable_weight_plan(
        context=context,
        sku_items=sku_items,
        sku_rules={},
        packing_mode="fast",
        cfg={
            **workflow_module.DEFAULT_CONFIG,
            "max_extra_boxes_per_order": 1,
            "repeat_retail_batch_sizes": [10],
            "repeat_retail_max_extra_boxes_per_order": 10,
            "repeat_retail_min_optimization_seconds": 0,
        },
        remaining_budget_seconds=None,
    )

    assert split_result.box_qty == 8
    assert warnings[0].rule_applied.startswith("repeat retail batch 10")


def test_repeat_retail_batch_candidate_keeps_fewer_boxes_when_savings_low(monkeypatch):
    lines = [_order_line("1", "RETAIL BOX", quantity=80)]
    context = workflow_module.PackingOrderContext(
        order_id="1",
        lines=lines,
        combo="RETAIL BOX x80",
        items=[],
        groups=[workflow_module.RuleSplitGroup(lines=lines)],
        cache_key="key",
        first_index=0,
    )
    sku_items = {
        "RETAIL BOX": _sku_item("RETAIL BOX", Dimensions(25, 16, 15), 2),
    }

    def fake_pack_group_records(**kwargs):
        box_qty = len(kwargs["groups"])
        return SplitResult(True, box_qty, [SplitCarton(index + 1, None) for index in range(box_qty)], []), []

    def fake_score(**kwargs):
        box_qty = kwargs["split_result"].box_qty
        chargeable_by_box_qty = {1: 40.0, 2: 39.0, 8: 39.5}
        return workflow_module.CandidatePlanScore(
            total_chargeable_weight_kg=chargeable_by_box_qty.get(box_qty, 45.0),
            box_qty=box_qty,
            box_type_count=box_qty,
            total_assigned_volume_cm3=box_qty * 1000,
        )

    monkeypatch.setattr(workflow_module, "_pack_group_records", fake_pack_group_records)
    monkeypatch.setattr(workflow_module, "_score_assigned_split_result", fake_score)

    split_result, warnings = workflow_module._select_chargeable_weight_plan(
        context=context,
        sku_items=sku_items,
        sku_rules={},
        packing_mode="fast",
        cfg={
            **workflow_module.DEFAULT_CONFIG,
            "max_extra_boxes_per_order": 1,
            "repeat_retail_batch_sizes": [10],
            "repeat_retail_max_extra_boxes_per_order": 10,
            "repeat_retail_min_optimization_seconds": 0,
        },
        remaining_budget_seconds=None,
    )

    assert split_result.box_qty == 1
    assert warnings == []


def test_repeat_retail_strongbox_addon_candidates_are_generated():
    lines = [
        _order_line("1", "STRONGBOX", quantity=24),
        _order_line("1", "SLEEVES", quantity=24),
        _order_line("1", "PLAYMAT", quantity=6),
    ]
    sku_items = {
        "STRONGBOX": _sku_item("STRONGBOX", Dimensions(25, 16, 15), 2),
        "SLEEVES": _sku_item("SLEEVES", Dimensions(13, 9, 2), 0.12),
        "PLAYMAT": _sku_item("PLAYMAT", Dimensions(38, 8, 8), 0.5),
    }
    sku_rules = {
        "SLEEVES": SKUCampaignRule(key="SLEEVES", no_padding=True),
        "PLAYMAT": SKUCampaignRule(
            key="PLAYMAT",
            no_padding=True,
            wrap_around_largest_item=True,
            wrapped_height_cm=4,
        ),
    }

    candidates = workflow_module._chargeable_candidate_group_sets(
        [workflow_module.RuleSplitGroup(lines=lines)],
        sku_items,
        sku_rules,
        bundle_footprint_tolerance_cm=5,
        cfg={
            **workflow_module.DEFAULT_CONFIG,
            "max_optimization_seconds": 600,
            "repeat_retail_batch_sizes": [8],
        },
    )
    names = [name for name, _groups in candidates]

    assert "repeat retail batch 8 STRONGBOX" in names
    assert "repeat retail distributed 8 STRONGBOX" in names
    assert "repeat retail accessory split 8 STRONGBOX" in names

    distributed = dict(candidates)["repeat retail distributed 8 STRONGBOX"]
    assert len(distributed) == 3
    assert all(
        sum(line.quantity for line in group.lines if line.canonical_sku == "STRONGBOX") == 8
        for group in distributed
    )
    assert sum(
        line.quantity
        for group in distributed
        for line in group.lines
        if line.canonical_sku == "SLEEVES"
    ) == 24
    assert sum(
        line.quantity
        for group in distributed
        for line in group.lines
        if line.canonical_sku == "PLAYMAT"
    ) == 6


def test_oversized_vendor_box_baseline_can_use_two_extra_boxes(monkeypatch):
    lines = [
        _order_line("1", "LARGE A"),
        _order_line("1", "LARGE B"),
        _order_line("1", "SMALL A"),
        _order_line("1", "SMALL B"),
    ]
    context = workflow_module.PackingOrderContext(
        order_id="1",
        lines=lines,
        combo="LARGE A x1 | LARGE B x1 | SMALL A x1 | SMALL B x1",
        items=[],
        groups=[workflow_module.RuleSplitGroup(lines=lines)],
        cache_key="key",
        first_index=0,
    )
    sku_items = {
        "LARGE A": _sku_item("LARGE A", Dimensions(50, 30, 16), 2),
        "LARGE B": _sku_item("LARGE B", Dimensions(44, 28, 14), 1.5),
        "SMALL A": _sku_item("SMALL A", Dimensions(6, 5, 2), 0.2),
        "SMALL B": _sku_item("SMALL B", Dimensions(5, 5, 2), 0.2),
    }

    def fake_pack_group_records(**kwargs):
        box_qty = len(kwargs["groups"])
        return SplitResult(True, box_qty, [SplitCarton(index + 1, None) for index in range(box_qty)], []), []

    def fake_score(**kwargs):
        box_qty = kwargs["split_result"].box_qty
        chargeable_by_box_qty = {1: 30.0, 2: 29.4, 3: 24.0}
        oversized_by_box_qty = {1: 1, 2: 1, 3: 0}
        return workflow_module.CandidatePlanScore(
            total_chargeable_weight_kg=chargeable_by_box_qty[box_qty],
            box_qty=box_qty,
            box_type_count=box_qty,
            total_assigned_volume_cm3=box_qty * 1000,
            oversized_box_count=oversized_by_box_qty[box_qty],
        )

    monkeypatch.setattr(workflow_module, "_pack_group_records", fake_pack_group_records)
    monkeypatch.setattr(workflow_module, "_score_assigned_split_result", fake_score)

    split_result, warnings = workflow_module._select_chargeable_weight_plan(
        context=context,
        sku_items=sku_items,
        sku_rules={},
        packing_mode="fast",
        cfg=workflow_module.DEFAULT_CONFIG,
        remaining_budget_seconds=None,
    )

    assert split_result.box_qty == 3
    assert warnings[0].error_type == "OversizedVendorBoxPlanSelected"





def test_label_generator_uses_vfi_suffix_only_for_multi_box_and_box_specific_contents():
    rows = workflow_module._label_generator_rows(
        [
            {
                "Order ID": "1",
                "VFI #": "39",
                "Order Box ID": "1-1",
                "Box Number": 1,
                "Box Qty": 1,
                "Total Units": 9,
                "Unit Count": 4,
                "Box Type": "VB 25",
                "Chargeable Weight kg": 6.6,
                "SKU Breakdown": "CORE x4 | ADDON x5",
                "SKUs in Box": "CORE x4",
                "Id": "37048314",
                "Address Name": "Samuel Richard Rendall",
                "Address Phone Number": "275279857",
                "Email": "samuel@example.com",
                "Address Line 1": "64 BATKIN ROAD",
                "Address Line 2": "NEW WINDSOR",
                "Address City": "AUCKLAND",
                "Address Postal Code": "600",
                "Full Country": "New Zealand",
                "Address Country": "NZ",
                "Address State": "AUCKLAND",
            },
            {
                "Order ID": "2",
                "VFI #": "39",
                "Order Box ID": "2-2",
                "Box Number": 2,
                "Box Qty": 3,
                "Total Units": 12,
                "Unit Count": 5,
                "Box Type": "VB 36",
                "Chargeable Weight kg": 18.1,
                "SKU Breakdown": "CORE x9 | ADDON x3",
                "SKUs in Box": "ADDON x3 | CORE x2",
            },
        ],
        {"CORE x4 | ADDON x5": 7, "CORE x9 | ADDON x3": 8},
        "OPR",
    )

    assert rows[0]["Pledge Configuration"] == 7
    assert rows[0]["Label numbers"] == "OPR 39"
    assert rows[0]["Total Units"] == 4
    assert rows[0]["SKU Breakdown"] == "CORE x4"
    assert rows[0]["Backer ID"] == "37048314"
    assert rows[0]["Shipping name"] == "Samuel Richard Rendall"
    assert rows[0]["phone"] == "275279857"
    assert rows[0]["email"] == "samuel@example.com"
    assert rows[0]["add 1"] == "64 BATKIN ROAD"
    assert rows[0]["add 2"] == "NEW WINDSOR"
    assert rows[0]["Shipping City"] == "AUCKLAND"
    assert rows[0]["Shipping Postal Code"] == "600"
    assert rows[0]["Country Name"] == "New Zealand"
    assert rows[0]["Ship to Country Code"] == "NZ"
    assert rows[0]["Shipping State"] == "AUCKLAND"
    assert "Box Qty" not in rows[0]
    assert rows[1]["Pledge Configuration"] == 8
    assert rows[1]["Label numbers"] == "OPR 39-2"
    assert rows[1]["Total Units"] == 5
    assert rows[1]["SKU Breakdown"] == "ADDON x3 | CORE x2"


def test_country_sequences_follow_cost_summary_row_order_and_unknown_fallback():
    sequenced_rows, by_order = workflow_module._with_country_sequences(
        [
            {"Order ID": "1", "Country": "Hong Kong", "Code": "HK"},
            {"Order ID": "2", "Country": "Singapore", "Code": "SG"},
            {"Order ID": "3", "Country": "Hong Kong", "Code": "HK"},
            {"Order ID": "4", "Country": "", "Code": ""},
        ]
    )

    assert [row["Country Number"] for row in sequenced_rows] == [
        "Hong Kong 1",
        "Singapore 1",
        "Hong Kong 2",
        "Unknown 1",
    ]
    assert by_order["3"]["sequence"] == 2
    assert by_order["3"]["country_code"] == "HK"


def test_country_package_codes_keep_order_sequence_and_add_carton_suffix():
    _sequenced_rows, by_order = workflow_module._with_country_sequences(
        [
            {"Order ID": "1", "Country": "Hong Kong", "Code": "HK"},
            {"Order ID": "2", "Country": "Singapore", "Code": "SG"},
            {"Order ID": "3", "Country": "Hong Kong", "Code": "HK"},
        ]
    )

    box_rows = workflow_module._with_country_package_codes(
        [
            {"Order ID": "3", "Country": "Hong Kong", "Code": "HK", "Box Number": 1, "Box Qty": 3},
            {"Order ID": "3", "Country": "Hong Kong", "Code": "HK", "Box Number": 2, "Box Qty": 3},
            {"Order ID": "3", "Country": "Hong Kong", "Code": "HK", "Box Number": 3, "Box Qty": 3},
        ],
        by_order,
    )

    assert [row["Country Number"] for row in box_rows] == ["Hong Kong 2", "Hong Kong 2", "Hong Kong 2"]
    assert [row["Country Package Code"] for row in box_rows] == ["HK  2-1", "HK  2-2", "HK  2-3"]
    assert workflow_module._country_package_code("HK", 1, 1, 1) == "HK  1"
    assert workflow_module._country_package_code("SG", 1, 1, 1) == "SG  1"


def test_label_country_code_resolves_from_full_country_names_and_alias_fields():
    rows = [
        {"Country": "Germany"},
        {"Country Name": "France"},
        {"Shipping Country": "Netherlands"},
        {"Ship To Country": "Hong Kong"},
        {"Ship to Country": "China"},
        {"Address Country": "United Kingdom"},
    ]

    assert [workflow_module._country_code_for_label_row(row) for row in rows] == [
        "DE",
        "FR",
        "NL",
        "HK",
        "CN",
        "GB",
    ]


def test_label_country_code_keeps_explicit_code_first_and_does_not_invent_unknowns():
    assert workflow_module._country_code_for_label_row({"Country": "Germany", "Country Code": "FR"}) == "FR"
    assert workflow_module._country_code_for_label_row({"Country": "Atlantis"}) == ""
    assert workflow_module._country_package_code("", 12) == "UN  12"


def test_country_package_codes_use_country_name_fallback_before_un_fallback():
    _sequenced_rows, by_order = workflow_module._with_country_sequences(
        [
            {"Order ID": "1", "Country": "Germany"},
            {"Order ID": "2", "Country": "Germany"},
            {"Order ID": "3", "Country": "France"},
            {"Order ID": "4", "Country": "Netherlands"},
            {"Order ID": "5", "Country": "Atlantis"},
        ]
    )

    box_rows = workflow_module._with_country_package_codes(
        [
            {"Order ID": "1", "Country": "Germany", "Box Number": 1, "Box Qty": 1},
            {"Order ID": "2", "Country": "Germany", "Box Number": 1, "Box Qty": 2},
            {"Order ID": "2", "Country": "Germany", "Box Number": 2, "Box Qty": 2},
            {"Order ID": "3", "Country": "France", "Box Number": 1, "Box Qty": 1},
            {"Order ID": "4", "Country": "Netherlands", "Box Number": 1, "Box Qty": 1},
            {"Order ID": "5", "Country": "Atlantis", "Box Number": 1, "Box Qty": 1},
        ],
        by_order,
    )

    assert [row["Country Package Code"] for row in box_rows] == [
        "DE  1",
        "DE  2-1",
        "DE  2-2",
        "FR  1",
        "NL  1",
        "UN  1",
    ]


def test_cost_summary_removes_country_number_and_keeps_destination_country_and_vfi():
    rows = workflow_module._cost_summary_rows(
        [
            {
                "Backer ID": "B1",
                "VFI #": "1",
                "Country Number": "Hong Kong 1",
                "Shipping name": "Ada",
                "Country": "Hong Kong",
                "Total Units": 1,
            }
        ],
        {},
        workflow_module.CustomerRateSheetSelection(sheet=None, path="", filename="", source="missing"),
    )

    assert "Country Number" not in rows[0]
    assert list(rows[0])[:3] == ["Backer ID", "VFI #", "Shipping name"]
    assert rows[0]["VFI #"] == "1"
    assert rows[0]["Country"] == "Hong Kong"


def test_label_generator_and_label_header_carry_country_package_code_to_continuations():
    label_generator_rows = workflow_module._label_generator_rows(
        [
            {
                "Order ID": "3",
                "VFI #": "115",
                "Box Number": 1,
                "Box Qty": 3,
                "Unit Count": 7,
                "Box Type": "VB 41",
                "Country": "Hong Kong",
                "Code": "HK",
                "Country Number": "Hong Kong 2",
                "Country Package Code": "HK  2-1",
                "SKU Breakdown": " | ".join(f"SKU-{index} x1" for index in range(1, 7)),
                "SKUs in Box": " | ".join(f"SKU-{index} x1" for index in range(1, 7)),
            }
        ],
        {" | ".join(f"SKU-{index} x1" for index in range(1, 7)): 1},
        "VEST",
    )

    assert label_generator_rows[0]["Country Number"] == "Hong Kong 2"
    assert label_generator_rows[0]["Country Package Code"] == "HK  2-1"

    label_rows = workflow_module._labels_rows(
        label_generator_rows,
        {"label_item_lines_per_column": 1, "label_item_column_count": 2},
    )

    assert label_rows[0]["Country Package Code"] == "HK  2-1"
    assert label_rows[1]["Label Continuation"] is True
    assert label_rows[1]["Country Package Code"] == "HK  2-1"
    assert excel_writer_module._label_block_rows(label_rows[0])[0][4] == "HK"
    assert excel_writer_module._label_block_rows(label_rows[1])[0][4] == "HK"


def test_country_package_count_rows_count_physical_packages_not_continuation_labels():
    rows = workflow_module._country_package_count_rows(
        [
            {"Order ID": "1", "Country": "Hong Kong", "Box Number": 1},
            {"Order ID": "1", "Country": "Hong Kong", "Box Number": 2},
            {"Order ID": "2", "Country": "Singapore", "Box Number": 1},
        ]
    )

    assert rows == [
        {"Section": "Country Package Counts", "Metric": "Hong Kong", "Value": 2, "Detail": ""},
        {"Section": "Country Package Counts", "Metric": "Singapore", "Value": 1, "Detail": ""},
    ]


def test_country_scan_sheets_group_barcode_values_by_country_in_label_order():
    sheets = workflow_module._country_scan_sheets(
        [
            {"Country Name": "Hong Kong", "Country Number": "Hong Kong 1", "Label Number": "1", "Label numbers": "VEST 1"},
            {"Country Name": "Singapore", "Country Number": "Singapore 1", "Label Number": "2", "Label numbers": "VEST 2"},
            {"Country Name": "Hong Kong", "Country Number": "Hong Kong 2", "Label Number": "3", "Label numbers": "VEST 3"},
            {"Country Name": "Bad/Name", "Country Number": "Bad/Name 1", "Label Number": "4", "Label numbers": "VEST 4"},
            {"Country Name": "Bad:Name", "Country Number": "Bad:Name 1", "Label Number": "5", "Label numbers": "VEST 5"},
        ]
    )

    assert [row["VFI #"] for row in sheets["China-HK"]] == ["VEST 1", "VEST 3"]
    assert [row["VFI #"] for row in sheets["Non-Hub Countries"]] == ["VEST 4", "VEST 5", "VEST 2"]
    assert "Barcode Value" not in sheets["China-HK"][0]
    assert "Bad_Name" not in sheets
    assert "Bad_Name 2" not in sheets


def test_country_scan_sheets_keep_final_label_order_and_use_vfi_barcode_values():
    sheets = workflow_module._country_scan_sheets(
        [
            {"Country Name": "Hong Kong", "Country Number": "Hong Kong 10", "Label Number": "10", "Label numbers": "VEST 10"},
            {"Country Name": "Hong Kong", "Country Number": "Hong Kong 1", "Label Number": "1", "Label numbers": "VEST 1"},
            {"Country Name": "Hong Kong", "Country Number": "Hong Kong 3", "Label Number": "3", "Label numbers": "VEST 3"},
            {"Country Name": "Hong Kong", "Country Number": "Hong Kong 2", "Label Number": "2", "Label numbers": "VEST 2"},
            {"Country Name": "Hong Kong", "Country Number": "Hong Kong 21-2", "Label Number": "21-2", "Label numbers": "VEST 21-2"},
            {"Country Name": "Hong Kong", "Country Number": "Hong Kong 22-1", "Label Number": "22-1", "Label numbers": "VEST 22-1"},
            {"Country Name": "Hong Kong", "Country Number": "Hong Kong 21-1", "Label Number": "21-1", "Label numbers": "VEST 21-1"},
        ]
    )

    hong_kong_rows = sheets["China-HK"]

    assert "Country Number" not in hong_kong_rows[0]
    assert [row["VFI #"] for row in hong_kong_rows] == [
        "VEST 10",
        "VEST 1",
        "VEST 3",
        "VEST 2",
        "VEST 21-2",
        "VEST 22-1",
        "VEST 21-1",
    ]
    assert list(hong_kong_rows[0])[:2] == ["Campaign", "VFI #"]
    assert "Barcode Value" not in hong_kong_rows[0]
    assert list(hong_kong_rows[0])[2:4] == ["Actual weight g", "Volumetric weight kg"]


def test_country_scan_sheets_group_selected_hub_countries_into_shared_tabs():
    sheets = workflow_module._country_scan_sheets(
        [
            {"Country": "New Zealand", "Barcode/QR Value": "NZ-1", "Total Units": 2},
            {"Country": "Germany", "Barcode/QR Value": "DE-1", "Total Units": 1},
            {"Country": "Australia", "Barcode/QR Value": "AU-1", "Total Units": 3},
            {"Country": "Hong Kong", "Barcode/QR Value": "HK-1", "Total Units": 4},
            {"Country": "China", "Barcode/QR Value": "CN-1", "Total Units": 5},
            {"Country": "United Arab Emirates", "Barcode/QR Value": "AE-1", "Total Units": 6},
            {"Country": "Bahrain", "Barcode/QR Value": "BH-1", "Total Units": 7},
            {"Country": "Oman", "Barcode/QR Value": "OM-1", "Total Units": 8},
            {"Country": "Kuwait", "Barcode/QR Value": "KW-1", "Total Units": 9},
            {"Country": "Saudi Arabia", "Barcode/QR Value": "SA-1", "Total Units": 10},
        ]
    )

    assert [row["VFI #"] for row in sheets["Australia-NZ"]] == ["AU-1", "NZ-1"]
    assert [row["VFI #"] for row in sheets["China-HK"]] == ["CN-1", "HK-1"]
    assert [row["VFI #"] for row in sheets["Middle East Hub"]] == ["BH-1", "KW-1", "OM-1", "SA-1", "AE-1"]
    assert [row["VFI #"] for row in sheets["Non-Hub Countries"]] == ["DE-1"]
    assert "Australia" not in sheets
    assert "New Zealand" not in sheets
    assert "China" not in sheets
    assert "Hong Kong" not in sheets
    assert "Bahrain" not in sheets
    assert sheets["Australia-NZ"][0]["Items in box"] == 3
    assert sheets["China-HK"][0]["Items in box"] == 5
    assert sheets["Middle East Hub"][0]["Items in box"] == 7


def test_label_generator_populates_simplified_chinese_country_from_two_letter_code():
    rows = workflow_module._label_generator_rows(
        [
            {
                "Order ID": "1",
                "VFI #": "1",
                "Box Number": 1,
                "Box Qty": 1,
                "Unit Count": 1,
                "Box Type": "VB 25",
                "SKU Breakdown": "CORE x1",
                "SKUs in Box": "CORE x1",
                "Address Country": "NZ",
            },
            {
                "Order ID": "2",
                "VFI #": "2",
                "Box Number": 1,
                "Box Qty": 1,
                "Unit Count": 1,
                "Box Type": "VB 25",
                "SKU Breakdown": "CORE x1",
                "SKUs in Box": "CORE x1",
                "Address Country": "us",
                "Country Name Chinese": "自定义美国",
            },
            {
                "Order ID": "3",
                "VFI #": "3",
                "Box Number": 1,
                "Box Qty": 1,
                "Unit Count": 1,
                "Box Type": "VB 25",
                "SKU Breakdown": "CORE x1",
                "SKUs in Box": "CORE x1",
                "Address Country": "ZZ",
            },
        ],
        {"CORE x1": 1},
        "OPR",
    )

    assert rows[0]["Country Name Chinese"] == "新西兰"
    assert rows[1]["Country Name Chinese"] == "自定义美国"
    assert rows[2]["Country Name Chinese"] == ""
    assert workflow_module._country_name_zh_hans("HK") == "中国香港"

def test_label_generator_populates_country_code_and_chinese_from_country_name():
    rows = workflow_module._label_generator_rows(
        [
            {
                "Order ID": "1",
                "VFI #": "1",
                "Box Number": 1,
                "Box Qty": 1,
                "Unit Count": 1,
                "Box Type": "VB 25",
                "SKU Breakdown": "CORE x1",
                "SKUs in Box": "CORE x1",
                "Country": "Germany",
                "Country Package Code": "DE  1",
            },
            {
                "Order ID": "2",
                "VFI #": "2",
                "Box Number": 1,
                "Box Qty": 1,
                "Unit Count": 1,
                "Box Type": "VB 25",
                "SKU Breakdown": "CORE x1",
                "SKUs in Box": "CORE x1",
                "Country": "Germany",
                "Country Name Chinese": "è‡ªå®šä¹‰å¾·å›½",
            },
        ],
        {"CORE x1": 1},
        "OPR",
    )

    assert rows[0]["Ship to Country Code"] == "DE"
    assert rows[0]["Country Name Chinese"] == workflow_module._country_name_zh_hans("DE")
    assert rows[0]["Country Package Code"] == "DE  1"
    assert rows[1]["Country Name Chinese"] == "è‡ªå®šä¹‰å¾·å›½"


def test_label_generator_uses_backer_data_aliases_for_addresses_country_code_and_order_id():
    rows = workflow_module._label_generator_rows(
        [
            {
                "Order ID": "Column-1",
                "VFI #": "1",
                "Box Number": 1,
                "Box Qty": 1,
                "Unit Count": 1,
                "Box Type": "VB 33 cutdown",
                "SKU Breakdown": "CORE x1",
                "SKUs in Box": "CORE x1",
                "Order #": "BK-100",
                "Name": "Ralf Kwok",
                "Phone": "98835556",
                "CustomerEmail": "ralf@example.com",
                "Address 1": "Flat 12",
                "Address 2": "1 Example Road",
                "City": "Hong Kong",
                "PostalCode": "0000",
                "Country": "Hong Kong",
                "Code": "HK",
                "State/Province": "HK",
            }
        ],
        {"CORE x1": 1},
        "VEST",
    )

    assert rows[0]["Backer ID"] == "BK-100"
    assert rows[0]["Shipping name"] == "Ralf Kwok"
    assert rows[0]["phone"] == "98835556"
    assert rows[0]["email"] == "ralf@example.com"
    assert rows[0]["add 1"] == "Flat 12"
    assert rows[0]["add 2"] == "1 Example Road"
    assert rows[0]["Shipping City"] == "Hong Kong"
    assert rows[0]["Shipping Postal Code"] == "0000"
    assert rows[0]["Country Name"] == "Hong Kong"
    assert rows[0]["Ship to Country Code"] == "HK"
    assert rows[0]["Country Name Chinese"] == workflow_module._country_name_zh_hans("HK")
    assert rows[0]["Shipping State"] == "HK"


def test_label_country_code_resolves_full_country_name_but_not_unknown_country():
    rows = workflow_module._label_generator_rows(
        [
            {
                "Order ID": "1",
                "VFI #": "1",
                "Box Number": 1,
                "Box Qty": 1,
                "Unit Count": 1,
                "Box Type": "VB 25",
                "SKU Breakdown": "CORE x1",
                "SKUs in Box": "CORE x1",
                "Address Country": "New Zealand",
                "Country": "New Zealand",
            }
        ],
        {"CORE x1": 1},
        "OPR",
    )

    assert rows[0]["Ship to Country Code"] == "NZ"
    assert rows[0]["Country Name Chinese"] == workflow_module._country_name_zh_hans("NZ")

    unknown_rows = workflow_module._label_generator_rows(
        [
            {
                "Order ID": "2",
                "VFI #": "2",
                "Box Number": 1,
                "Box Qty": 1,
                "Unit Count": 1,
                "Box Type": "VB 25",
                "SKU Breakdown": "CORE x1",
                "SKUs in Box": "CORE x1",
                "Address Country": "Atlantis",
                "Country": "Atlantis",
            }
        ],
        {"CORE x1": 1},
        "OPR",
    )

    assert unknown_rows[0]["Ship to Country Code"] == ""
    assert unknown_rows[0]["Country Name Chinese"] == ""


def test_labels_rows_are_generated_one_per_carton_from_label_generator_rows():
    label_rows = workflow_module._labels_rows(
        [
            {
                "Pledge Configuration": 7,
                "Order ID": "1",
                "Total Units": 4,
                "Label numbers": "OPR 39",
                "Box Plan": "VB 25",
                "Per-Box Chargeable Weight": 6.6,
                "SKU Breakdown": "CORE x4",
                "Backer ID": "37048314",
                "Shipping name": "Samuel Richard Rendall",
                "phone": "275279857",
                "email": "samuel@example.com",
                "add 1": "64 BATKIN ROAD",
                "add 2": "NEW WINDSOR",
                "Shipping City": "AUCKLAND",
                "Shipping Postal Code": "600",
                "Country Name": "New Zealand",
                "Ship to Country Code": "NZ",
                "Country Name Chinese": "新西兰",
                "Shipping State": "AUCKLAND",
            },
            {
                "Pledge Configuration": 8,
                "Order ID": "2",
                "Total Units": 5,
                "Label numbers": "OPR 39-2",
                "Box Plan": "VB 36",
                "Per-Box Chargeable Weight": 18.1,
                "SKU Breakdown": "ADDON x3 | CORE x2 | SLEEVE x1",
            },
        ],
        {"campaign": {"name": "Orange Prism", "country_name_chinese": "中国"}},
    )

    assert len(label_rows) == 2
    assert label_rows[0]["Label Number"] == "39"
    assert label_rows[0]["Label Value"] == "OPR 39"
    assert label_rows[0]["Barcode/QR Value"] == "OPR 39"
    assert label_rows[0]["Campaign Name"] == "Orange Prism"
    assert label_rows[0]["Country Name Chinese"] == "新西兰"
    assert label_rows[0]["From"] == workflow_module.LABEL_FROM_LINE
    assert label_rows[0]["Origin"] == "CN"
    assert label_rows[0]["Total Value USD"] == ""
    assert label_rows[0]["To Name"] == "Samuel Richard Rendall"
    assert label_rows[0]["Backer ID"] == "37048314"
    assert label_rows[0]["Address Line 1"] == "64 BATKIN ROAD"
    assert label_rows[0]["Address Line 2"] == "NEW WINDSOR"
    assert label_rows[0]["City"] == "AUCKLAND"
    assert label_rows[0]["State/Province"] == "AUCKLAND"
    assert label_rows[0]["Postal/Zip"] == "600"
    assert label_rows[0]["Country"] == "New Zealand"
    assert label_rows[0]["Country Code"] == "NZ"
    assert label_rows[0]["Phone"] == "275279857"
    assert label_rows[0]["Email"] == "samuel@example.com"
    assert label_rows[0]["Order ID"] == "1"
    assert label_rows[0]["Pledge Configuration"] == 7
    assert label_rows[0]["Carton Box Designation"] == "VB 25"
    assert label_rows[0]["Total Units"] == 4
    assert label_rows[0]["SKU Breakdown"] == "CORE x4"
    assert label_rows[0]["Items to Pack Column 1"] == "CORE x4"
    assert label_rows[0]["Items to Pack Column 2"] == ""
    assert label_rows[0]["On Arrival Note"] == "If product damage is found, please report to Support@VFI.asia in 24 hours"

    assert label_rows[1]["Label Number"] == "39-2"
    assert label_rows[1]["Label Value"] == "OPR 39-2"
    assert label_rows[1]["Barcode/QR Value"] == "OPR 39-2"
    assert label_rows[1]["Country Name Chinese"] == "中国"
    assert label_rows[1]["Total Units"] == 5
    assert label_rows[1]["Items to Pack Column 1"] == "ADDON x3\nCORE x2\nSLEEVE x1"
    assert label_rows[1]["Items to Pack Column 2"] == ""


def test_label_visible_header_shows_number_before_project_without_changing_barcode_value():
    label_rows = workflow_module._labels_rows(
        [
            {
                "Order ID": "1",
                "Label numbers": "emerald 28-1",
                "Ship to Country Code": "HK",
                "Country Package Code": "HK  3-1",
                "SKU Breakdown": "CORE x1",
            }
        ]
    )

    assert label_rows[0]["Barcode/QR Value"] == "emerald 28-1"
    assert excel_writer_module._label_block_rows(label_rows[0])[0] == ["28-1 emerald", "", "", "", "HK", ""]


def test_labels_rows_prefix_item_lines_with_intake_picking_order():
    label_rows = workflow_module._labels_rows(
        [
            {
                "Order ID": "PICK-1",
                "Label numbers": "OPR 1",
                "SKU Breakdown": "ADDON x2 | UNKNOWN x1 | CORE x1 | COIN x3",
            }
        ],
        {},
        sku_pick_order={"CORE": 0, "ADDON": 1, "COIN": 2},
    )

    assert label_rows[0]["Items to Pack Column 1"].splitlines() == [
        "(1)  CORE x1",
        "(2)  ADDON x2",
        "(3)  COIN x3",
        "(?)  UNKNOWN x1",
    ]
    assert label_rows[0]["Items to Pack Column 2"] == ""


def test_labels_rows_print_order_groups_hub_countries_first_then_pledge_stably():
    label_rows = [
        {
            "Pledge Configuration": 2,
            "Country": "United States",
            "Order ID": "A",
            "Label Value": "LC 4",
        },
        {
            "Pledge Configuration": 1,
            "Country": "United States",
            "Order ID": "B",
            "Label Value": "LC 1",
        },
        {
            "Pledge Configuration": 1,
            "Country": "Canada",
            "Order ID": "C",
            "Label Value": "LC 2",
        },
        {
            "Pledge Configuration": 1,
            "Country": "Canada",
            "Order ID": "D",
            "Label Value": "LC 3",
        },
        {
            "Pledge Configuration": 2,
            "Country": "China",
            "Order ID": "E",
            "Label Value": "LC 5",
        },
        {
            "Pledge Configuration": "10",
            "Country": "Germany",
            "Order ID": "F",
            "Label Value": "LC 6",
        },
        {
            "Pledge Configuration": "2",
            "Country": "France",
            "Order ID": "G",
            "Label Value": "LC 7",
        },
    ]

    ordered_rows = workflow_module._labels_rows_in_print_order(label_rows)

    assert [row["Country"] for row in ordered_rows] == [
        "Canada",
        "Canada",
        "China",
        "United States",
        "United States",
        "France",
        "Germany",
    ]
    assert [row["Pledge Configuration"] for row in ordered_rows] == [1, 1, 2, 1, 2, "2", "10"]
    assert [row["Order ID"] for row in ordered_rows] == ["C", "D", "E", "B", "A", "G", "F"]
    assert [row["Label Value"] for row in ordered_rows] == ["LC 2", "LC 3", "LC 5", "LC 1", "LC 4", "LC 7", "LC 6"]


def test_visible_vfi_numbers_are_assigned_after_country_first_print_order_and_drive_scan_values():
    label_rows = [
        {
            "Pledge Configuration": 2,
            "Country": "Hong Kong",
            "Order ID": "HK-C2",
            "Box Number": 1,
            "Box Qty": 1,
            "Label Value": "old",
            "Barcode/QR Value": "old",
        },
        {
            "Pledge Configuration": 1,
            "Country": "China",
            "Order ID": "CN-C1",
            "Box Number": 1,
            "Box Qty": 1,
            "Label Value": "old",
            "Barcode/QR Value": "old",
        },
        {
            "Pledge Configuration": 3,
            "Country": "China",
            "Order ID": "CN-C3",
            "Box Number": 1,
            "Box Qty": 3,
            "Label Value": "old",
            "Barcode/QR Value": "old",
        },
        {
            "Pledge Configuration": 3,
            "Country": "China",
            "Order ID": "CN-C3",
            "Box Number": 2,
            "Box Qty": 3,
            "Label Value": "old",
            "Barcode/QR Value": "old",
        },
        {
            "Pledge Configuration": 3,
            "Country": "China",
            "Order ID": "CN-C3",
            "Box Number": 3,
            "Box Qty": 3,
            "Label Value": "old",
            "Barcode/QR Value": "old",
        },
        {
            "Pledge Configuration": 3,
            "Country": "China",
            "Order ID": "CN-C3",
            "Box Number": 3,
            "Box Qty": 3,
            "Label Continuation": True,
            "Original Label Number": "old continued",
            "Label Value": "old",
            "Barcode/QR Value": "old",
        },
        {
            "Pledge Configuration": 2,
            "Country": "China",
            "Order ID": "CN-C2",
            "Box Number": 1,
            "Box Qty": 1,
            "Label Value": "old",
            "Barcode/QR Value": "old",
        },
    ]

    ordered_rows = workflow_module._labels_rows_in_print_order(label_rows)
    numbered_rows, display_vfi_by_order, _display_label_by_package = workflow_module._assign_visible_vfi_numbers(
        ordered_rows,
        "emerald",
    )

    main_rows = [row for row in numbered_rows if not row.get("Label Continuation")]
    assert [row["Country"] for row in main_rows] == ["China", "China", "China", "China", "China", "Hong Kong"]
    assert [row["Pledge Configuration"] for row in main_rows] == [1, 2, 3, 3, 3, 2]
    assert [row["Barcode/QR Value"] for row in main_rows] == [
        "1 emerald",
        "2 emerald",
        "3 1 of 3 emerald",
        "3 2 of 3 emerald",
        "3 3 of 3 emerald",
        "4 emerald",
    ]
    continuation = next(row for row in numbered_rows if row.get("Label Continuation"))
    assert continuation["Original Label Number"] == "3 3 of 3 emerald"
    assert continuation["Barcode/QR Value"] == "3 3 of 3 emerald"
    assert display_vfi_by_order["CN-C3"] == "3 emerald"

    scan_sheets = workflow_module._country_scan_sheets(numbered_rows, campaign_name="Emerald Long Campaign")
    assert [row["VFI #"] for row in scan_sheets["China-HK"]] == [
        "1 emerald",
        "2 emerald",
        "3 1 of 3 emerald",
        "3 2 of 3 emerald",
        "3 3 of 3 emerald",
        "4 emerald",
    ]
    assert [row["Campaign"] for row in scan_sheets["China-HK"]] == ["Emerald Long Campaign"] * 6
    assert "Country" not in scan_sheets["China-HK"][0]
    assert scan_sheets["China-HK"][0]["Items in this box / SKU contents"] == ""


def test_labels_rows_respect_campaign_item_line_controls_and_report_overflow():
    label_rows = workflow_module._labels_rows(
        [
            {
                "Pledge Configuration": 9,
                "Order ID": "LONG",
                "Total Units": 9,
                "Label numbers": "OPR 40",
                "Box Plan": "VB 39",
                "SKU Breakdown": "A x1 | B x1 | C x1 | D x1 | E x1",
            }
        ],
        {
            "campaign": {"name": "Orange Prism"},
            "label_item_lines_per_column": 2,
            "label_item_column_count": 2,
        },
    )

    assert len(label_rows) == 2
    row = label_rows[0]
    assert row["Items to Pack Column 1"] == "A x1\nB x1"
    assert row["Items to Pack Column 2"] == "C x1\nD x1"
    assert row["Label Item Lines Per Column"] == 2
    assert row["Label Item Column Count"] == 2
    assert row["Label Overflow Items Per Label"] == workflow_module.DEFAULT_LABEL_OVERFLOW_ITEMS_PER_LABEL
    assert row["Overflow Item Count"] == 1
    assert row["Overflow Items"] == ""
    assert row["Overflow Note"] == ""
    assert row["SKU Breakdown"] == "A x1 | B x1 | C x1 | D x1 | E x1"

    continuation = label_rows[1]
    assert continuation["Label Continuation"] is True
    assert continuation["Label Number"] == "40 CONTINUED"
    assert continuation["Original Label Number"] == "40"
    assert continuation["Barcode/QR Value"] == "OPR 40"
    assert continuation["Items to Pack Column 1"] == "E x1"
    assert continuation["Items to Pack Column 2"] == ""


def test_labels_rows_can_use_single_item_column_for_dense_campaigns():
    label_rows = workflow_module._labels_rows(
        [
            {
                "Order ID": "ONE-COL",
                "Label numbers": "OPR 41",
                "SKU Breakdown": "A x1 | B x1 | C x1",
            }
        ],
        {"label_item_lines_per_column": 2, "label_item_column_count": 1},
    )

    assert len(label_rows) == 2
    row = label_rows[0]
    assert row["Items to Pack Column 1"] == "A x1\nB x1"
    assert row["Items to Pack Column 2"] == ""
    assert row["Label Item Column Count"] == 1
    assert row["Overflow Item Count"] == 1
    assert row["Overflow Items"] == ""
    assert label_rows[1]["Items to Pack Column 1"] == "C x1"


def test_labels_rows_support_multiple_continuation_labels():
    label_rows = workflow_module._labels_rows(
        [
            {
                "Order ID": "MULTI-CONTINUED",
                "Label numbers": "OPR 42",
                "SKU Breakdown": " | ".join(f"SKU-{index} x1" for index in range(1, 10)),
            }
        ],
        {
            "label_item_lines_per_column": 2,
            "label_item_column_count": 2,
            "label_overflow_items_per_label": 2,
        },
    )

    assert len(label_rows) == 4
    assert label_rows[0]["Items to Pack Column 1"] == "SKU-1 x1\nSKU-2 x1"
    assert label_rows[0]["Items to Pack Column 2"] == "SKU-3 x1\nSKU-4 x1"
    assert [row["Label Number"] for row in label_rows[1:]] == [
        "42 CONTINUED 1",
        "42 CONTINUED 2",
        "42 CONTINUED 3",
    ]
    assert [row["Items to Pack Column 1"] for row in label_rows[1:]] == [
        "SKU-5 x1\nSKU-6 x1",
        "SKU-7 x1\nSKU-8 x1",
        "SKU-9 x1",
    ]
    assert all(row["Barcode/QR Value"] == "OPR 42" for row in label_rows)


def test_label_generator_expands_bundles_for_label_units_and_sku_lines():
    rows = workflow_module._label_generator_rows(
        [
            {
                "Order ID": "BUNDLE-1",
                "VFI #": "17",
                "Box Number": 1,
                "Box Qty": 1,
                "Unit Count": 1,
                "Label Unit Count": 3,
                "Box Type": "VB 39",
                "Chargeable Weight kg": 4.2,
                "SKU Breakdown": "CORE x1 | ADDON x2",
                "SKUs in Box": "BUNDLE[ADDON x2 | CORE x1] x1",
                "Label SKUs in Box": "ADDON x2 | CORE x1",
            }
        ],
        {"CORE x1 | ADDON x2": 1},
        "EMERALD",
    )

    assert rows[0]["Total Units"] == 3
    assert rows[0]["SKU Breakdown"] == "ADDON x2 | CORE x1"
    assert "BUNDLE[" not in rows[0]["SKU Breakdown"]


def test_label_sku_counts_expand_bundled_placements_without_changing_pack_diagnostics():
    placements = [
        Placement(
            "BUNDLE[CORE x1 | ADDON x2]",
            1,
            Dimensions(20, 10, 5),
            (0, 0, 0),
            2,
        ),
        Placement("COIN", 3, Dimensions(5, 5, 1), (0, 0, 5), 0.2),
    ]

    counts = workflow_module._label_sku_counts_from_placements(placements)

    assert counts == {"ADDON": 2, "COIN": 3, "CORE": 1}
    assert workflow_module._sku_counts_text(counts) == "ADDON x2 | COIN x3 | CORE x1"
    assert (
        workflow_module._sku_counts_text(counts, {"CORE": 0, "ADDON": 1, "COIN": 2})
        == "CORE x1 | ADDON x2 | COIN x3"
    )


def test_label_sku_breakdown_follows_sku_master_order(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {"SKU": "CORE", "Length": "10", "Width": "10", "Height": "2", "Weight kg": "1"},
            {"SKU": "ADDON", "Length": "8", "Width": "8", "Height": "2", "Weight kg": "0.5"},
            {"SKU": "COIN", "Length": "2", "Width": "2", "Height": "1", "Weight kg": "0.1"},
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "1", "SKU": "COIN", "Quantity": "3"},
            {"Order ID": "1", "SKU": "ADDON", "Quantity": "2"},
            {"Order ID": "1", "SKU": "CORE", "Quantity": "1"},
        ],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={"preserve_region_sheets": False},
    )

    label_row = _sheet_rows(output_path, "Label generator")[0]
    assert label_row["SKU Breakdown"] == "CORE x1 | ADDON x2 | COIN x3"
    cost_row = _sheet_rows(output_path, "Cost Summary")[0]
    assert cost_row["SKU Breakdown"] == "CORE x1 | ADDON x2 | COIN x3"
    optimized_row = _sheet_rows(output_path, "Optimized to Pack")[0]
    assert optimized_row["All Items"] == "CORE x1, ADDON x2, COIN x3"
    assert optimized_row["Box 1"].endswith("CORE x1, ADDON x2, COIN x3")


def test_workbook_cost_summary_removes_country_number_and_scan_tabs_use_label_vfi_values(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [{"SKU": "CORE", "Product Name": "Core", "Length": "10", "Width": "8", "Height": "2", "Weight kg": "1"}],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "1", "SKU": "CORE", "Quantity": "1", "Country": "Hong Kong", "Country Code": "HK"},
            {"Order ID": "2", "SKU": "CORE", "Quantity": "1", "Country": "Singapore", "Country Code": "SG"},
            {"Order ID": "3", "SKU": "CORE", "Quantity": "1", "Country": "Hong Kong", "Country Code": "HK"},
        ],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={"packing_mode": "fast", "preserve_region_sheets": False, "campaign": {"code": "TEST"}},
    )

    workbook = read_workbook(str(output_path))
    sheet_names = _workbook_sheet_names(output_path)
    labels_index = sheet_names.index("Labels")
    assert sheet_names[labels_index + 1 : labels_index + 3] == ["China-HK", "Singapore"]

    cost_rows = _sheet_rows(output_path, "Cost Summary")
    assert "Country Number" not in cost_rows[0]
    assert [row["Country"] for row in cost_rows] == ["Hong Kong", "Hong Kong", "Singapore"]
    assert [row["VFI #"] for row in cost_rows] == ["1 TEST", "2 TEST", "3 TEST"]

    summary_rows = _sheet_rows(output_path, "Summary")
    country_counts = {
        row["Metric"]: row["Value"]
        for row in summary_rows
        if row["Section"] == "Country Package Counts"
    }
    assert country_counts == {"Hong Kong": "2", "Singapore": "1"}

    hong_kong_scan = _sheet_rows(output_path, "China-HK")
    singapore_scan = _sheet_rows(output_path, "Singapore")
    assert "Country Number" not in hong_kong_scan[0]
    assert [row["VFI #"] for row in hong_kong_scan] == ["1 TEST", "2 TEST"]
    assert hong_kong_scan[0]["Items in box"] == "1"
    assert "Barcode Value" not in hong_kong_scan[0]
    hong_kong_xml = _sheet_xml(output_path, "China-HK")
    assert _inline_cell_text(hong_kong_xml, "A1") == "Campaign"
    assert _inline_cell_text(hong_kong_xml, "B1") == "VFI #"
    assert _inline_cell_text(hong_kong_xml, "C1") == "Actual weight g"
    assert _inline_cell_text(hong_kong_xml, "D1") == "Volumetric weight kg"
    assert _inline_cell_text(hong_kong_xml, "C2") == ""
    assert _inline_cell_text(hong_kong_xml, "D2") == ""
    hong_kong_root = ElementTree.fromstring(hong_kong_xml)
    hong_kong_widths = {
        int(column.attrib["min"]): float(column.attrib["width"])
        for column in hong_kong_root.findall("main:cols/main:col", _NS)
    }
    assert hong_kong_widths[3] == 13
    assert hong_kong_widths[4] == 16
    assert hong_kong_root.find(".//main:c[@r='D2']", _NS).attrib["s"] == "31"
    assert _inline_cell_text(hong_kong_xml, "O1") == "Items in box"
    assert _inline_cell_text(hong_kong_xml, "P1") == "Items in this box / SKU contents"
    assert singapore_scan[0]["VFI #"] == "3 TEST"
    assert "Backer ID" not in hong_kong_scan[0]
    assert "Order ID" not in hong_kong_scan[0]
    assert hong_kong_scan[0]["Items in this box / SKU contents"] == "(1)  CORE x1"
    assert "Package / carton identifier" not in hong_kong_scan[0]
    assert "Carton index" not in hong_kong_scan[0]
    assert "Total cartons" not in hong_kong_scan[0]
    assert "Pledge Configuration" not in hong_kong_scan[0]


def test_label_generator_carries_first_valid_backer_notes_and_ignores_sku_notes():
    row = {
        "Order ID": "1",
        "Box Qty": 1,
        "SKU Breakdown": "CORE x1",
        "Label SKUs in Box": "CORE x1",
        "Box Type": "VB 4",
        "Chargeable Weight kg": 1,
        "Total Units": 1,
        "Label Unit Count": 1,
        "SKU Notes": "do not use",
        "Customer Notes": "Leave package at front desk",
        "Delivery Notes": "do not concatenate",
    }

    labels = workflow_module._label_generator_rows([row], {"CORE x1": 1}, "TEST")

    assert workflow_module._valid_order_notes_header("Shipping Notes")
    assert workflow_module._valid_order_notes_header("customernotes")
    assert workflow_module._valid_order_notes_header("Delivernotes")
    assert not workflow_module._valid_order_notes_header("SKU Notes")
    assert labels[0]["Notes"] == "Leave package at front desk"


def test_country_scan_sheets_use_active_hub_pricing_and_preserve_intake_metadata(tmp_path):
    rate_sheet = workflow_module.CustomerRateSheet(
        hub=workflow_module.CustomerRateLane(
            rates_by_zone={
                "Zone JP": {1.0: 10.0},
                "Zone SG": {1.0: 12.0},
            },
            zone_by_country={
                "japan": "Zone JP",
                "jp": "Zone JP",
                "singapore": "Zone SG",
                "sg": "Zone SG",
            },
            method_by_country={},
        ),
        express=workflow_module.CustomerRateLane(
            rates_by_zone={},
            zone_by_country={},
            method_by_country={},
        ),
    )
    label_rows = [
        {
            "Country": "Japan",
            "Barcode/QR Value": "1 VFI",
            "Order ID": "JP-1",
            "Items to Pack Column 1": "CORE x1",
            "_Country Scan Metadata": {
                "Backer ID": "B-JP",
                "Name": "Japan Backer",
                "Email": "jp@example.com",
                "Address 1": "1 Tokyo Rd",
                "Address 2": "Suite 7",
                "City": "Tokyo",
                "Postal/Zip": "100",
                "Add to": "Apartment call box 7",
                "Tax ID number": "JP-TAX",
                "Tax ID": "JP-TAX-ALT",
                "Original Optimized Box Type": "VB 4",
            },
        },
        {
            "Country": "Singapore",
            "Barcode/QR Value": "2 VFI",
            "Backer ID": "B-SG",
            "To Name": "Singapore Backer",
            "Email": "sg@example.com",
            "Address Line 1": "2 Marina Way",
            "Address Line 2": "#02-01",
            "City": "Singapore",
            "Postal/Zip": "018956",
            "Order ID": "SG-1",
            "Items to Pack Column 1": "CORE x1",
            "VAT": "SG-VAT",
        },
        {
            "Country": "Germany",
            "Barcode/QR Value": "3 VFI",
            "Backer ID": "B-DE",
            "To Name": "Germany Backer",
            "Order ID": "DE-1",
            "Items to Pack Column 1": "CORE x1",
            "Company": "DE Company",
            "Original Optimized Box Type": "VB 8",
        },
        {
            "Country": "France",
            "Barcode/QR Value": "4 VFI",
            "Backer ID": "B-FR",
            "To Name": "France Backer",
            "Order ID": "FR-1",
            "Items to Pack Column 1": "CORE x1",
        },
    ]

    sheets = workflow_module._country_scan_sheets(label_rows, rate_sheet, "Long Human Campaign")

    assert list(sheets) == ["Japan", "Singapore", "Non-Hub Countries"]
    assert [row["VFI #"] for row in sheets["Non-Hub Countries"]] == ["4 VFI", "3 VFI"]
    assert sheets["Japan"][0]["Campaign"] == "Long Human Campaign"
    assert sheets["Japan"][0]["Add to"] == "Apartment call box 7"
    assert sheets["Japan"][0]["Tax ID number"] == "JP-TAX"
    assert sheets["Japan"][0]["Tax ID"] == "JP-TAX-ALT"
    assert sheets["Japan"][0]["Address 2"] == "Suite 7"
    assert "add 2" not in sheets["Japan"][0]
    assert sheets["Singapore"][0]["Address Line 2"] == "#02-01"
    assert sheets["Singapore"][0]["VAT"] == "SG-VAT"
    assert sheets["Non-Hub Countries"][1]["Company"] == "DE Company"
    for sheet_rows in sheets.values():
        assert "Barcode Value" not in sheet_rows[0]
        assert "Barcode/QR Value" not in sheet_rows[0]
        assert "Package Carton Identifier" not in sheet_rows[0]
        assert "Carton Index" not in sheet_rows[0]
        assert "Total Cartons" not in sheet_rows[0]
        assert "Pledge Configuration" not in sheet_rows[0]
        assert "Original Optimized Box Type" not in sheet_rows[0]

    output_path = tmp_path / "country_scan.xlsx"
    excel_writer_module.write_workbook(str(output_path), country_scan_sheets=sheets)
    japan_xml = _sheet_xml(output_path, "Japan")
    assert _inline_cell_text(japan_xml, "A1") == "Campaign"
    assert _inline_cell_text(japan_xml, "B1") == "VFI #"
    assert _inline_cell_text(japan_xml, "C1") == "Actual weight g"
    assert _inline_cell_text(japan_xml, "D1") == "Volumetric weight kg"
    assert _inline_cell_text(japan_xml, "C2") == ""
    assert _inline_cell_text(japan_xml, "D2") == ""
    assert [
        _inline_cell_text(japan_xml, reference)
        for reference in [
            "A1",
            "B1",
            "C1",
            "D1",
            "E1",
            "F1",
            "G1",
            "H1",
            "I1",
            "J1",
            "K1",
            "L1",
            "M1",
            "N1",
            "O1",
            "P1",
            "Q1",
        ]
    ] == [
        "Campaign",
        "VFI #",
        "Actual weight g",
        "Volumetric weight kg",
        "Backer ID",
        "Name",
        "Email",
        "Address 1",
        "Address 2",
        "City",
        "Postal/Zip",
        "Add to",
        "Tax ID number",
        "Tax ID",
        "",
        "Items in box",
        "Items in this box / SKU contents",
    ]
    assert _inline_cell_text(japan_xml, "I2") == "Suite 7"
    assert _inline_cell_text(japan_xml, "O2") == ""
    assert _inline_cell_text(japan_xml, "P2") == ""
    assert _inline_cell_text(japan_xml, "Q2") == "CORE x1"


def test_country_scan_multi_package_rows_lookup_actual_dimensions_by_package_barcode(tmp_path):
    label_rows = [
        {
            "Country": "Hong Kong",
            "Barcode/QR Value": "15 1 of 2 ITFFKS1",
            "Label Number": "15 1 of 2 ITFFKS1",
            "Order ID": "ORDER-15",
            "Box Number": 1,
            "Box Qty": 2,
        },
        {
            "Country": "Hong Kong",
            "Barcode/QR Value": "15 2 of 2 ITFFKS1",
            "Label Number": "15 2 of 2 ITFFKS1",
            "Order ID": "ORDER-15",
            "Box Number": 2,
            "Box Qty": 2,
        },
    ]
    sheets = workflow_module._country_scan_sheets(label_rows, campaign_name="TEST Campaign")
    output_path = tmp_path / "country_scan_multi_package.xlsx"

    excel_writer_module.write_workbook(
        str(output_path),
        actual_dimensions_rows=workflow_module._actual_dimensions_rows(
            expected_scan_barcodes=["15 1 of 2 ITFFKS1", "15 2 of 2 ITFFKS1"]
        ),
        country_scan_sheets=sheets,
    )

    hong_kong_xml = _sheet_xml(output_path, "China-HK")

    assert _inline_cell_text(hong_kong_xml, "B2") == "15 1 of 2 ITFFKS1"
    assert _inline_cell_text(hong_kong_xml, "B3") == "15 2 of 2 ITFFKS1"
    assert _cell_formula(hong_kong_xml, "C2") == (
        'IFERROR(IF(INDEX(\'Actual Dimensions\'!$A:$A,MATCH($B2,\'Actual Dimensions\'!$A:$A,0))="",'
        '"",INDEX(\'Actual Dimensions\'!$B:$B,MATCH($B2,\'Actual Dimensions\'!$A:$A,0))),"")'
    )
    assert _cell_formula(hong_kong_xml, "D2") == (
        'IFERROR(IF(INDEX(\'Actual Dimensions\'!$A:$A,MATCH($B2,\'Actual Dimensions\'!$A:$A,0))="",'
        '"",INDEX(\'Actual Dimensions\'!$F:$F,MATCH($B2,\'Actual Dimensions\'!$A:$A,0))),"")'
    )
    assert "MATCH($B3,'Actual Dimensions'!$A:$A,0)" in _cell_formula(hong_kong_xml, "C3")
    assert "MATCH($B3,'Actual Dimensions'!$A:$A,0)" in _cell_formula(hong_kong_xml, "D3")
    assert "Expected scan group VFI key" not in _cell_formula(hong_kong_xml, "C2")
    assert "Expected scan group VFI key" not in _cell_formula(hong_kong_xml, "D2")


def test_label_generator_does_not_duplicate_order_notes_for_multi_carton_orders():
    row = {
        "Order ID": "1",
        "Box Qty": 2,
        "SKU Breakdown": "CORE x1",
        "Label SKUs in Box": "CORE x1",
        "Box Type": "VB 4",
        "Chargeable Weight kg": 1,
        "Total Units": 1,
        "Label Unit Count": 1,
        "Shipping Notes": "Leave package at front desk",
    }

    labels = workflow_module._label_generator_rows([row], {"CORE x1": 1}, "TEST")

    assert labels[0]["Notes"] == ""


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


def test_identical_sku_combinations_reuse_packing_plan_and_preserve_metadata(tmp_path, monkeypatch):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {"SKU": "Core", "Product Name": "Core", "Length": "10", "Width": "8", "Height": "4", "Weight kg": "1"},
            {"SKU": "Addon", "Product Name": "Addon", "Length": "6", "Width": "4", "Height": "2", "Weight kg": "0.2"},
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "1001", "SKU": "Core", "Quantity": "1", "Name": "Alice", "Email": "alice@example.com"},
            {"Order ID": "1001", "SKU": "Addon", "Quantity": "1", "Name": "Alice", "Email": "alice@example.com"},
            {"Order ID": "1002", "SKU": "Addon", "Quantity": "1", "Name": "Bob", "Email": "bob@example.com"},
            {"Order ID": "1002", "SKU": "Core", "Quantity": "1", "Name": "Bob", "Email": "bob@example.com"},
        ],
    )
    calls = {"count": 0}

    def fake_split_order_into_cartons(items, packing_mode="normal", force_simple_split=False, **kwargs):
        calls["count"] += 1
        length = 20 + calls["count"]
        placements = [
            Placement(
                canonical_sku=item.canonical_sku,
                quantity=1,
                dimensions=item.padded_dimensions,
                origin=(0, 0, 0),
                weight_kg=item.weight_kg,
            )
            for item in items
        ]
        return SplitResult(
            success=True,
            box_qty=1,
            cartons=[
                SplitCarton(
                    box_number=1,
                    result=OptimizedCartonResult(
                        success=True,
                        length_cm=length,
                        width_cm=12,
                        height_cm=8,
                        chargeable_weight_kg=1,
                        volume_cm3=length * 12 * 8,
                        placements=placements,
                        unplaced_items=[],
                    ),
                )
            ],
            unplaced_items=[],
        )

    monkeypatch.setattr("box_optimizer.workflow.split_order_into_cartons", fake_split_order_into_cartons)

    summary = optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={
            "packing_mode": "fast",
            "use_vendor_box_menu": False,
            "preserve_region_sheets": False,
        },
    )

    assert calls["count"] == 1
    assert summary["parsed_order_count"] == 2
    assert summary["unique_packing_cache_key_count"] == 1
    assert summary["packing_solve_cache_miss_count"] == 1
    assert summary["packing_cache_reuse_count"] == 1
    order_rows = {row["Order ID"]: row for row in _sheet_rows(output_path, "Order Volume Weights")}
    assert order_rows["1001"]["Name"] == "Alice"
    assert order_rows["1002"]["Name"] == "Bob"
    assert [row["VFI #"] for row in sorted(order_rows.values(), key=lambda row: row["Order ID"])] == ["1 VFI", "2 VFI"]
    detail_rows = sorted(_sheet_rows(output_path, "Multi Box Detail"), key=lambda row: row["Order ID"])
    assert [row["Order Box ID"] for row in detail_rows] == ["1001-1", "1002-1"]
    assert {row["Length cm"] for row in detail_rows} == {"23"}
    label_rows = sorted(_sheet_rows(output_path, "Label generator"), key=lambda row: row["Order ID"])
    assert [row["Order ID"] for row in label_rows] == ["1001", "1002"]
    cost_rows = sorted(_sheet_rows(output_path, "Cost Summary"), key=lambda row: row["Backer ID"])
    assert len(cost_rows) == 2
    assert [row["VFI #"] for row in cost_rows] == ["1 VFI", "2 VFI"]
    debug_rows = {
        row["Metric"]: row["Value"]
        for row in _sheet_rows(output_path, "Debug Summary")
        if row["Section"] == "Performance"
    }
    assert debug_rows["Unique Packing Cache Keys"] == "1"
    assert debug_rows["Representative Packing Solves"] == "1"
    assert debug_rows["Packing Cache Reuses"] == "1"
    summary_rows = _sheet_rows(output_path, "Box Size Summary")
    assert len(summary_rows) == 1
    assert int(float(summary_rows[0]["Box Count"])) == 2


def test_same_sku_quantities_with_different_packing_cache_keys_solve_separately(tmp_path, monkeypatch):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {"SKU": "Core", "Product Name": "Core", "Length": "10", "Width": "8", "Height": "4", "Weight kg": "1"},
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "1001", "SKU": "Core", "Quantity": "1", "Country": "United States"},
            {"Order ID": "1002", "SKU": "Core", "Quantity": "1", "Country": "Canada"},
        ],
    )
    calls = {"count": 0}

    def fake_split_order_into_cartons(items, packing_mode="normal", force_simple_split=False, **kwargs):
        calls["count"] += 1
        item = items[0]
        length = 20 + calls["count"]
        return SplitResult(
            success=True,
            box_qty=1,
            cartons=[
                SplitCarton(
                    box_number=1,
                    result=OptimizedCartonResult(
                        success=True,
                        length_cm=length,
                        width_cm=12,
                        height_cm=8,
                        chargeable_weight_kg=1,
                        volume_cm3=length * 12 * 8,
                        placements=[Placement(item.canonical_sku, 1, item.padded_dimensions, (0, 0, 0), item.weight_kg)],
                        unplaced_items=[],
                    ),
                )
            ],
            unplaced_items=[],
        )

    monkeypatch.setattr("box_optimizer.workflow.split_order_into_cartons", fake_split_order_into_cartons)

    summary = optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={
            "packing_mode": "fast",
            "use_vendor_box_menu": False,
            "preserve_region_sheets": False,
            "company_protection_rate_bands": {"default": {"1": 1}},
            "company_protection_country_zones": {"Canada": "Zone Canada", "default": "Zone Other"},
        },
    )

    assert calls["count"] == 2
    assert summary["parsed_order_count"] == 2
    assert summary["unique_packing_cache_key_count"] == 2
    assert summary["packing_solve_cache_miss_count"] == 2
    assert summary["packing_cache_reuse_count"] == 0
    detail_rows = sorted(_sheet_rows(output_path, "Multi Box Detail"), key=lambda row: row["Order ID"])
    assert [row["Order ID"] for row in detail_rows] == ["1001", "1002"]
    assert len(_sheet_rows(output_path, "Label generator")) == 2
    assert len(_sheet_rows(output_path, "Cost Summary")) == 2


def test_all_in_ship_alone_does_not_force_every_sku_into_own_carton(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {"SKU": "All-in Storage Solution [72570]", "Product Name": "All-in Storage Solution [72570]", "Length": "60", "Width": "30", "Height": "30", "Weight kg": "2"},
            {"SKU": "Core", "Product Name": "Core", "Length": "30", "Width": "20", "Height": "7", "Weight kg": "1"},
            {"SKU": "Expansion", "Product Name": "Expansion", "Length": "24", "Width": "16", "Height": "5", "Weight kg": "0.8"},
            {"SKU": "Dice", "Product Name": "Dice", "Length": "8", "Width": "5", "Height": "2", "Weight kg": "0.2"},
            {"SKU": "Tokens", "Product Name": "Tokens", "Length": "10", "Width": "7", "Height": "2", "Weight kg": "0.2"},
            {"SKU": "Sleeves", "Product Name": "Sleeves", "Length": "12", "Width": "8", "Height": "2", "Weight kg": "0.2"},
            {"SKU": "Tray", "Product Name": "Tray", "Length": "18", "Width": "12", "Height": "3", "Weight kg": "0.4"},
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "1", "SKU": "All-in Storage Solution [72570]", "Quantity": "1"},
            {"Order ID": "1", "SKU": "Core", "Quantity": "1"},
            {"Order ID": "1", "SKU": "Expansion", "Quantity": "1"},
            {"Order ID": "1", "SKU": "Dice", "Quantity": "1"},
            {"Order ID": "1", "SKU": "Tokens", "Quantity": "1"},
            {"Order ID": "1", "SKU": "Sleeves", "Quantity": "1"},
            {"Order ID": "1", "SKU": "Tray", "Quantity": "1"},
        ],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={
            "packing_mode": "fast",
            "sku_rules": {
                "All-in Storage Solution [72570]": {
                    "prepacked": True,
                    "no_padding": True,
                    "ships_alone": True,
                    "can_mix_with_other_items": False,
                    "box_type": "All-in Storage Solution carton",
                }
            },
            "preserve_region_sheets": False,
        },
    )

    order_row = _sheet_rows(output_path, "Order Volume Weights")[0]
    assert int(float(order_row["Box Qty"])) == 2
    detail_rows = _sheet_rows(output_path, "Multi Box Detail")
    assert len(detail_rows) == 2
    all_in_row = next(row for row in detail_rows if row["Box Type"] == "All-in Storage Solution carton")
    combined_row = next(row for row in detail_rows if row["Box Type"] != "All-in Storage Solution carton")
    assert all_in_row["SKUs in Box"] == "ALL-IN STORAGE SOLUTION [72570] x1"
    assert "CORE x1" in combined_row["SKUs in Box"]
    assert "EXPANSION x1" in combined_row["SKUs in Box"]
    assert "DICE x1" in combined_row["SKUs in Box"]
    warning_messages = [row["Message"] for row in _sheet_rows(output_path, "Errors and Warnings")]
    assert any("Order split due to ships_alone=true" in message for message in warning_messages)
    assert all("prepacked/forced-box" not in message for message in warning_messages)


def test_ship_as_is_friendly_label_displays_on_printable_labels_only(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    ship_as_is_sku = "Earth Under Siege All In Storage"
    full_carton_name = f"{ship_as_is_sku} shipping carton"
    _write_csv(
        sku_master_path,
        [
            {
                "SKU": ship_as_is_sku,
                "Product Name": ship_as_is_sku,
                "Length": "60",
                "Width": "30",
                "Height": "30",
                "Weight kg": "2",
            },
        ],
    )
    _write_csv(
        orders_path,
        [{"Order ID": "1", "SKU": ship_as_is_sku, "Quantity": "1", "Country": "Germany"}],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={
            "packing_mode": "fast",
            "preserve_region_sheets": False,
            "sku_rules": {
                ship_as_is_sku: {
                    "prepacked": True,
                    "no_padding": True,
                    "ships_alone": True,
                    "can_mix_with_other_items": False,
                    "box_type": full_carton_name,
                    "label_box_type": "All In",
                }
            },
        },
    )

    label_generator_row = _sheet_rows(output_path, "Label generator")[0]
    labels_xml = _sheet_xml(output_path, "Labels")
    multi_box_row = _sheet_rows(output_path, "Multi Box Detail")[0]
    box_size_row = _sheet_rows(output_path, "Box Size Summary")[0]

    assert label_generator_row["Box Plan"] == "All In"
    assert "All In" in labels_xml
    assert full_carton_name not in labels_xml
    assert multi_box_row["Box Type"] == full_carton_name
    assert box_size_row["Box Type"] == full_carton_name


def test_blank_ship_as_is_friendly_label_keeps_carton_name_on_printable_labels():
    rows = workflow_module._label_generator_rows(
        [
            {
                "Order ID": "1",
                "VFI #": "1",
                "Box Number": 1,
                "Box Qty": 1,
                "Unit Count": 1,
                "Box Type": "Earth Under Siege All In Storage shipping carton",
                "Label Box Type": "",
                "SKU Breakdown": "CORE x1",
                "SKUs in Box": "CORE x1",
                "Country": "Germany",
            }
        ],
        {"CORE x1": 1},
        "EUS",
    )

    label_rows = workflow_module._labels_rows(rows)

    assert rows[0]["Box Plan"] == "Earth Under Siege All In Storage shipping carton"
    assert label_rows[0]["Carton Box Designation"] == "Earth Under Siege All In Storage shipping carton"


def test_normal_vendor_box_label_display_ignores_unrelated_friendly_labels():
    rows = workflow_module._label_generator_rows(
        [
            {
                "Order ID": "1",
                "VFI #": "1",
                "Box Number": 1,
                "Box Qty": 1,
                "Unit Count": 1,
                "Box Type": "VB 25",
                "Label Box Type": "",
                "SKU Breakdown": "CORE x1",
                "SKUs in Box": "CORE x1",
                "Country": "Germany",
            }
        ],
        {"CORE x1": 1},
        "EUS",
    )

    label_rows = workflow_module._labels_rows(rows)

    assert rows[0]["Box Plan"] == "VB 25"
    assert label_rows[0]["Carton Box Designation"] == "VB 25"


def test_string_false_can_mix_config_still_forces_separation(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {"SKU": "Solo", "Product Name": "Solo", "Length": "20", "Width": "12", "Height": "5", "Weight kg": "1"},
            {"SKU": "Addon", "Product Name": "Addon", "Length": "8", "Width": "5", "Height": "2", "Weight kg": "0.2"},
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "1", "SKU": "Solo", "Quantity": "1"},
            {"Order ID": "1", "SKU": "Addon", "Quantity": "1"},
        ],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={
            "packing_mode": "fast",
            "sku_rules": {"Solo": {"can_mix_with_other_items": "false"}},
            "preserve_region_sheets": False,
        },
    )

    row = _sheet_rows(output_path, "Order Volume Weights")[0]
    assert int(float(row["Box Qty"])) == 2
    warning_messages = [warning["Message"] for warning in _sheet_rows(output_path, "Errors and Warnings")]
    assert any("can_mix_with_other_items=false" in message for message in warning_messages)


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


def test_balanced_workflow_prioritizes_high_volume_pledge_combinations(tmp_path, monkeypatch):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {"SKU": "Rare", "Product Name": "Rare", "Length": "10", "Width": "8", "Height": "4", "Weight kg": "1"},
            {"SKU": "Common", "Product Name": "Common", "Length": "10", "Width": "8", "Height": "4", "Weight kg": "1"},
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "1", "SKU": "Rare", "Quantity": "1"},
            {"Order ID": "2", "SKU": "Common", "Quantity": "1"},
            {"Order ID": "3", "SKU": "Common", "Quantity": "1"},
        ],
    )
    calls = []

    def fake_split_order_into_cartons(items, packing_mode="normal", force_simple_split=False, **kwargs):
        calls.append((items[0].canonical_sku, packing_mode))
        item = items[0]
        return SplitResult(
            success=True,
            box_qty=1,
            cartons=[
                SplitCarton(
                    box_number=1,
                    result=OptimizedCartonResult(
                        success=True,
                        length_cm=10,
                        width_cm=8,
                        height_cm=4,
                        chargeable_weight_kg=1,
                        volume_cm3=320,
                        placements=[Placement(item.canonical_sku, 1, item.padded_dimensions, (0, 0, 0), item.weight_kg)],
                        unplaced_items=[],
                    ),
                )
            ],
            unplaced_items=[],
        )

    monkeypatch.setattr("box_optimizer.workflow.split_order_into_cartons", fake_split_order_into_cartons)

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={"packing_mode": "balanced", "max_optimization_seconds": 180, "preserve_region_sheets": False},
    )

    assert calls[:2] == [("COMMON", "balanced"), ("RARE", "balanced")]


def test_balanced_workflow_uses_fast_for_uncached_combo_when_budget_is_low(tmp_path, monkeypatch):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {"SKU": "A", "Product Name": "A", "Length": "10", "Width": "8", "Height": "4", "Weight kg": "1"},
            {"SKU": "B", "Product Name": "B", "Length": "10", "Width": "8", "Height": "4", "Weight kg": "1"},
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "1", "SKU": "A", "Quantity": "1"},
            {"Order ID": "2", "SKU": "B", "Quantity": "1"},
        ],
    )
    calls = []
    ticks = iter([0, 0, 100, 100, 100, 100, 100, 100, 100, 100])

    def fake_perf_counter():
        return next(ticks, 100)

    def fake_split_order_into_cartons(items, packing_mode="normal", force_simple_split=False, **kwargs):
        calls.append(packing_mode)
        item = items[0]
        return SplitResult(
            success=True,
            box_qty=1,
            cartons=[
                SplitCarton(
                    box_number=1,
                    result=OptimizedCartonResult(
                        success=True,
                        length_cm=10,
                        width_cm=8,
                        height_cm=4,
                        chargeable_weight_kg=1,
                        volume_cm3=320,
                        placements=[Placement(item.canonical_sku, 1, item.padded_dimensions, (0, 0, 0), item.weight_kg)],
                        unplaced_items=[],
                    ),
                )
            ],
            unplaced_items=[],
        )

    monkeypatch.setattr("box_optimizer.workflow.time.perf_counter", fake_perf_counter)
    monkeypatch.setattr("box_optimizer.workflow.split_order_into_cartons", fake_split_order_into_cartons)

    summary = optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={"packing_mode": "balanced", "max_optimization_seconds": 1, "preserve_region_sheets": False},
    )

    assert summary["orders_processed"] == 2
    assert output_path.read_bytes()[:2] == b"PK"
    assert calls == ["fast", "fast"]


def test_workflow_log_event_includes_plain_message_fields(caplog):
    from box_optimizer import workflow as workflow_module

    with caplog.at_level(logging.INFO, logger="box_optimizer"):
        workflow_module._log_event(
            "order_packing_started",
            order_id="2572092W",
            combo_rank=3,
            pledge_count=84,
            mode="balanced",
            item_count=16,
        )

    message = caplog.records[-1].getMessage()
    assert "order_packing_started" in message
    assert "order_id=2572092W" in message
    assert "combo_rank=3" in message
    assert "pledge_count=84" in message
    assert "mode=balanced" in message
    assert "item_count=16" in message




def test_candidate_plan_selection_message_handles_same_more_and_fewer_boxes():
    baseline = workflow_module.CandidatePlanScore(20.0, 2, 1, 1000)
    same_boxes = workflow_module.CandidatePlanScore(17.6, 2, 1, 900)
    more_boxes = workflow_module.CandidatePlanScore(15.0, 3, 1, 800)
    fewer_boxes = workflow_module.CandidatePlanScore(14.5, 1, 1, 700)

    assert workflow_module._candidate_plan_selection_message(same_boxes, baseline, 2.4) == (
        "Selected alternate 2-box layout; saved 2.4 kg chargeable weight."
    )
    assert workflow_module._candidate_plan_selection_message(more_boxes, baseline, 5.0) == (
        "Selected 3-box plan over 2-box plan; saved 5.0 kg chargeable weight."
    )
    assert workflow_module._candidate_plan_selection_message(fewer_boxes, baseline, 5.5) == (
        "Selected fewer-box plan; saved 1 box and 5.5 kg chargeable weight."
    )


def _cap_test_split_result() -> SplitResult:
    carton_result = OptimizedCartonResult(
        success=True,
        length_cm=73,
        width_cm=36,
        height_cm=43,
        chargeable_weight_kg=23,
        volume_cm3=73 * 36 * 43,
        placements=[],
        unplaced_items=[],
    )
    return SplitResult(True, 1, [SplitCarton(1, carton_result)], [])


def _cap_test_assignment(vendor_box_id=None):
    return workflow_module.StandardizedBoxAssignment(
        order_id="1",
        combination_key="combo",
        box_type="VB 36" if vendor_box_id else "Custom 74x37x44",
        optimized_length_cm=74,
        optimized_width_cm=37,
        optimized_height_cm=44,
        assigned_length_cm=74,
        assigned_width_cm=37,
        assigned_height_cm=44,
        box_standardization_note="",
        placements=[],
        vendor_box_id=vendor_box_id,
        selection_decision="vendor" if vendor_box_id else "optimized",
    )


def test_vendor_box_near_cap_does_not_create_noisy_carton_cap_warning():
    warning = workflow_module._carton_cap_warning_for_order(
        order_id="1",
        combo="Large x1",
        split_result=_cap_test_split_result(),
        assignments_by_key={"1#1": _cap_test_assignment(vendor_box_id="36")},
    )

    assert warning is None


def test_true_optimized_cap_warning_still_appears():
    warning = workflow_module._carton_cap_warning_for_order(
        order_id="1",
        combo="Large x1",
        split_result=_cap_test_split_result(),
        assignments_by_key={"1#1": _cap_test_assignment()},
    )

    assert warning is not None
    assert warning.error_type == "CartonCapWarning"
    assert "reported carton dimensions were capped" in warning.message


def test_large_retail_bulk_review_note_mentions_driver_and_split_reason():
    lines = [
        OrderLine(order_id="27", raw_sku="A", canonical_sku="A", quantity=20),
        OrderLine(order_id="27", raw_sku="B", canonical_sku="B", quantity=10),
    ]
    box_rows = [
        {"Box Qty": 3, "Packed Actual Weight kg": "10", "Dimensional Weight kg (/5000)": "15", "Chargeable Weight g": 15000},
        {"Box Qty": 3, "Packed Actual Weight kg": "10", "Dimensional Weight kg (/5000)": "15", "Chargeable Weight g": 15000},
        {"Box Qty": 3, "Packed Actual Weight kg": "10", "Dimensional Weight kg (/5000)": "15", "Chargeable Weight g": 15000},
    ]
    warning_rows = [
        workflow_module.WorkflowWarning(
            order_id="27",
            stage="packing",
            error_type="ChargeableWeightPlanSelected",
            message="Selected alternate 3-box layout; saved 1 kg chargeable weight.",
        )
    ]

    warning = workflow_module._retail_bulk_review_warning(
        order_id="27",
        lines=lines,
        order_box_rows=box_rows,
        warning_rows=warning_rows,
        combo="A x20 | B x10",
    )

    assert warning is not None
    assert warning.error_type == "RetailBulkReview"
    assert "30 units across 3 boxes" in warning.message
    assert "dimensional weight" in warning.message
    assert "optimized alternate layout" in warning.message



def _vendor_flex_test_split(sku: str) -> SplitResult:
    result = OptimizedCartonResult(
        success=True,
        length_cm=34.2,
        width_cm=32.8,
        height_cm=18,
        chargeable_weight_kg=5,
        volume_cm3=34.2 * 32.8 * 18,
        placements=[Placement(sku, 1, Dimensions(34.2, 32.8, 18), (0, 0, 0), 1)],
        unplaced_items=[],
    )
    return SplitResult(True, 1, [SplitCarton(1, result)], [])


def test_vendor_box_fit_auto_allows_wrap_carton_only():
    cfg = {**workflow_module.DEFAULT_CONFIG, "use_vendor_box_menu": True}
    assignments = workflow_module._standardize_split_result(
        {
            "flex": _vendor_flex_test_split("PLAYMAT"),
            "rigid": _vendor_flex_test_split("RIGID"),
        },
        {"flex": "PLAYMAT x1", "rigid": "RIGID x1"},
        cfg,
        {"PLAYMAT": SKUCampaignRule(key="PLAYMAT", wrap_around_largest_item=True, no_padding=True)},
    )
    by_order = {assignment.order_id: assignment for assignment in assignments}

    assert "fit tolerance" in by_order["flex#1"].box_standardization_note
    assert "fit tolerance" not in by_order["rigid#1"].box_standardization_note


def test_vendor_box_fit_auto_does_not_apply_campaign_flex_to_separate_rigid_carton():
    cfg = {**workflow_module.DEFAULT_CONFIG, "use_vendor_box_menu": True}
    rigid_result = OptimizedCartonResult(
        success=True,
        length_cm=31.4,
        width_cm=31.4,
        height_cm=9.8,
        chargeable_weight_kg=2.3,
        volume_cm3=31.4 * 31.4 * 9.8,
        placements=[Placement("CORE", 1, Dimensions(31.4, 31.4, 9.8), (0, 0, 0), 1.95)],
        unplaced_items=[],
    )
    assignments = workflow_module._standardize_split_result(
        {
            "rigid": SplitResult(True, 1, [SplitCarton(1, rigid_result)], []),
            "flex": _vendor_flex_test_split("PLAYMAT"),
        },
        {"rigid": "CORE x1", "flex": "PLAYMAT x1"},
        cfg,
        {"PLAYMAT": SKUCampaignRule(key="PLAYMAT", wrap_around_largest_item=True, no_padding=True)},
    )
    by_order = {assignment.order_id: assignment for assignment in assignments}

    assert by_order["rigid#1"].vendor_box_id == "15"
    assert "fit tolerance" not in by_order["rigid#1"].box_standardization_note
    assert "fit tolerance" in by_order["flex#1"].box_standardization_note


def test_vendor_box_fit_auto_allows_compressible_carton():
    cfg = {**workflow_module.DEFAULT_CONFIG, "use_vendor_box_menu": True}
    assignments = workflow_module._standardize_split_result(
        {"soft": _vendor_flex_test_split("PLUSH")},
        {"soft": "PLUSH x1"},
        cfg,
        {"PLUSH": SKUCampaignRule(key="PLUSH", compressible=True)},
    )

    assert "fit tolerance" in assignments[0].box_standardization_note


def test_vendor_box_fit_mode_off_disables_flex_for_flexible_carton():
    cfg = {**workflow_module.DEFAULT_CONFIG, "use_vendor_box_menu": True, "vendor_box_fit_mode": "off"}
    assignments = workflow_module._standardize_split_result(
        {"flex": _vendor_flex_test_split("PLAYMAT")},
        {"flex": "PLAYMAT x1"},
        cfg,
        {"PLAYMAT": SKUCampaignRule(key="PLAYMAT", wrap_around_largest_item=True)},
    )

    assert "fit tolerance" not in assignments[0].box_standardization_note


def test_vendor_box_fit_mode_on_allows_flex_for_rigid_carton():
    cfg = {**workflow_module.DEFAULT_CONFIG, "use_vendor_box_menu": True, "vendor_box_fit_mode": "on"}
    assignments = workflow_module._standardize_split_result(
        {"rigid": _vendor_flex_test_split("RIGID")},
        {"rigid": "RIGID x1"},
        cfg,
        {},
    )

    assert "fit tolerance" in assignments[0].box_standardization_note





def _new_rate_sheet_table(hub_zone_usa_rate: float) -> dict[str, list[list[object]]]:
    return {
        "Zone Key": [
            ["HUB"],
            ["KG", 0.5, 1, 1.5],
            ["Zone USA", hub_zone_usa_rate, hub_zone_usa_rate, hub_zone_usa_rate],
            ["Zone 1", 21, 22, 23],
            ["Zone 2", 31, 32, 33],
            ["Zone 3", 41, 42, 43],
            ["Zone 9", "", "", ""],
            *([[""]] * 24),
            ["Express"],
            ["KG", 0.5, 1, 1.5],
            ["Zone 1", 101, 111, 121],
            ["Zone 2", 201, 211, 221],
            ["Zone 3", 301, 311, 321],
            ["Zone 4", 401, 411, 421],
        ],
        "Sheet2": [
            ["", "HUB", "", "", "", "", "", "", "Express", "", "", ""],
            ["", "COUNTRY", "A2 (ISO)", "Ship by", "Zone", "ship by", "ship to contact", "", "COUNTRY", "", "A2 (ISO)", "Zone"],
            ["", "United States", "US", "Hub", "Zone USA", "ship by hub", "hub@example.com", "", "United States", "", "US", "Zone 1"],
        ],
    }


def _write_active_rate_sheet_metadata_for_test(path: Path, original_filename: str) -> None:
    metadata = {
        "original_filename": original_filename,
        "saved_filename": path.name,
        "uploaded_at": "2026-05-31T12:00:00+00:00",
        "size_bytes": path.stat().st_size,
        "sha256": sha256_file(path),
        "validation": {"zone_key": True, "mapping_sheet_count": 1, "sheet_names": ["Zone Key", "Sheet2"]},
        "source": "active_upload",
    }
    rate_sheet_metadata_path().write_text(json.dumps(metadata), encoding="utf-8")


class _FakeRateSyncResponse:
    def __init__(self, payload: bytes):
        self.payload = payload

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return False

    def read(self) -> bytes:
        return self.payload


def _install_fake_rate_sync(
    monkeypatch,
    remote_status: dict,
    workbook_bytes: bytes | None = None,
    fail_current: bool = False,
    fail_download: bool = False,
) -> None:
    def fake_urlopen(request, timeout=0):
        url = getattr(request, "full_url", str(request))
        if "/rates/current" in url:
            if fail_current:
                raise OSError("railway unavailable")
            return _FakeRateSyncResponse(json.dumps(remote_status).encode("utf-8"))
        if "/rates/download" in url:
            if fail_download:
                raise OSError("download failed")
            return _FakeRateSyncResponse(workbook_bytes or b"")
        raise AssertionError(f"Unexpected rate sync URL: {url}")

    monkeypatch.setattr(rate_source_module, "urlopen", fake_urlopen)


def test_cost_summary_uses_customer_rate_sheet_for_zone_and_shipping_fee(tmp_path, monkeypatch):
    monkeypatch.setenv("BOX_OPTIMIZER_RATE_SHEET_DIR", str(tmp_path / "managed_rates"))
    monkeypatch.delenv("BOX_OPTIMIZER_RATE_SYNC_URL", raising=False)
    rate_path = tmp_path / "rates.xlsx"
    _write_xlsx_table(
        rate_path,
        "Zone Key",
        [
            ["HUB", "", "", ""],
            ["KG", 0.5, 1, 1.5],
            ["Zone USA", 11.15, 12.25, 13.35],
            ["Zone 1", 21.0, 22.0, 23.0],
            ["Zone 2", 31.0, 32.0, 33.0],
            ["", "", "", ""],
            ["Zone 0", "Zone USA", "Zone 1", "Zone 2"],
            ["China", "USA", "Australia", "Indonesia"],
        ],
    )

    rows = workflow_module._cost_summary_rows(
        [
            {"Country": "United States", "US State Abbreviation": "CA", "Chargeable Weight kg": 1.1, "Total Units": 1},
            {"Country": "United States", "US State Abbreviation": "HI", "Chargeable Weight kg": 0.7, "Total Units": 3},
            {
                "Country": "Australia",
                "Chargeable Weight kg": 1.0,
                "Total Units": 5,
                "Id": "37043809",
                "Address Name": "Jonathon Adkins",
                "Address Phone Number": "423732691",
                "Email": "harry@example.com",
            },
        ],
        {**workflow_module.DEFAULT_CONFIG, "rate_sheet_path": str(rate_path)},
    )

    assert rows[0]["Hub Shipping Fee"] == 15.35
    assert rows[0]["Express"] == 0
    assert rows[0]["Shipping Method"] == "Review Needed"
    assert "Shipping method mapping missing" in rows[0]["Shipping Rate Note"]
    assert rows[1]["Hub Shipping Fee"] == 24.5
    assert rows[1]["Express"] == 0
    assert rows[2]["Hub Shipping Fee"] == 25.0
    assert rows[2]["Express"] == 0
    assert "Zone" not in rows[0]
    assert "Shipping Method" in rows[0]
    assert rows[2]["Backer ID"] == "37043809"
    assert rows[2]["Shipping name"] == "Jonathon Adkins"
    assert rows[2]["phone"] == "423732691"
    assert rows[2]["email"] == "harry@example.com"
    assert "Customer Shipping Fee" not in rows[0]
    assert "Slow Post" not in rows[0]
    assert "Estimated VFI Cost" not in rows[0]
    assert "Picking Fee" not in rows[0]
    assert "Margin" not in rows[0]


def test_cost_summary_adds_fixed_separate_playmat_charge_to_active_hub_fee(tmp_path, monkeypatch):
    monkeypatch.setenv("BOX_OPTIMIZER_RATE_SHEET_DIR", str(tmp_path / "managed_rates"))
    monkeypatch.delenv("BOX_OPTIMIZER_RATE_SYNC_URL", raising=False)
    rate_path = tmp_path / "rates.xlsx"
    _write_xlsx_table(
        rate_path,
        "Zone Key",
        [
            ["HUB", "", "", ""],
            ["KG", 0.5, 1, 1.5],
            ["Zone USA", 11.15, 12.25, 13.35],
            ["", "", "", ""],
            ["Zone 0", "Zone USA"],
            ["United States", "USA"],
        ],
    )
    cfg = {
        **workflow_module.DEFAULT_CONFIG,
        "rate_sheet_path": str(rate_path),
        "separate_playmat_charge_skus": ["PLAYMAT"],
    }
    sku_rules = workflow_module._parse_sku_rules(cfg)

    rows = workflow_module._cost_summary_rows(
        [
            {
                "Country": "United States",
                "US State Abbreviation": "CA",
                "Chargeable Weight kg": 1.1,
                "Total Units": 3,
                "SKU Breakdown": "PLAYMAT x2 | ADDON x1",
            },
        ],
        cfg,
        sku_rules=sku_rules,
    )

    assert rows[0]["Hub Shipping Fee"] == 27.85
    assert rows[0]["Express"] == 0
    assert "Separate Playmat Charge: 2 units x $6.00 = $12.00" in rows[0]["Shipping Rate Note"]
    assert "Separate Playmat Charge" not in rows[0]
    assert workflow_module._cost_summary_total_cost(rows) == 27.85
    summary_rows = workflow_module._clean_summary_rows(
        {
            "orders_processed": 1,
            "boxes_created": 1,
            "box_types": 1,
            "unmatched_skus": 0,
            "total_shipping_cost": workflow_module._cost_summary_total_cost(rows),
            "total_shipping_cost_detail": "Hub Shipping Fee + Express",
        },
        [],
        [],
        sku_rules,
    )
    total_row = next(row for row in summary_rows if row["Metric"] == "Total Chargeable Cost")
    assert total_row["Value"] == 27.85


def test_cost_summary_adds_fixed_separate_playmat_charge_to_express_fallback(tmp_path, monkeypatch):
    monkeypatch.setenv("BOX_OPTIMIZER_RATE_SHEET_DIR", str(tmp_path / "managed_rates"))
    monkeypatch.delenv("BOX_OPTIMIZER_RATE_SYNC_URL", raising=False)
    rate_path = tmp_path / "Rate and Zone.xlsx"
    zone_key = [[""] for _ in range(37)]
    zone_key[0] = ["HUB"]
    zone_key[1] = ["KG", 0.5, 1, 1.5]
    zone_key[2] = ["Zone 9", "", "", ""]
    zone_key[31] = ["Express"]
    zone_key[32] = ["KG", 0.5, 1, 1.5]
    zone_key[33] = ["Zone 2", 201, 211, 221]
    sheet2 = [
        ["", "HUB", "", "", "", "", "", "", "Express", "", "", ""],
        ["", "COUNTRY", "A2 (ISO)", "Ship by", "Zone", "ship by", "ship to contact", "", "COUNTRY", "", "A2 (ISO)", "Zone"],
        ["", "Fallbackia", "FB", "Hub", "Zone 9", "ship to fallback", "", "", "Fallbackia", "", "FB", "Zone 2"],
    ]
    _write_xlsx_tables(rate_path, {"Zone Key": zone_key, "Sheet2": sheet2})
    cfg = {
        **workflow_module.DEFAULT_CONFIG,
        "rate_sheet_path": str(rate_path),
        "separate_playmat_charge_skus": ["PLAYMAT"],
    }
    sku_rules = workflow_module._parse_sku_rules(cfg)

    rows = workflow_module._cost_summary_rows(
        [
            {
                "Country": "Fallbackia",
                "Chargeable Weight kg": 0.6,
                "Total Units": 1,
                "SKU Breakdown": "PLAYMAT x1",
            },
        ],
        cfg,
        sku_rules=sku_rules,
    )

    assert rows[0]["Hub Shipping Fee"] == 0
    assert rows[0]["Express"] == 219
    assert "Express fallback; hub unavailable." in rows[0]["Shipping Rate Note"]
    assert "Separate Playmat Charge: 1 units x $6.00 = $6.00" in rows[0]["Shipping Rate Note"]


def test_cost_summary_uses_express_fallback_from_rate_and_zone_sheet(tmp_path, monkeypatch):
    monkeypatch.setenv("BOX_OPTIMIZER_RATE_SHEET_DIR", str(tmp_path / "managed_rates"))
    monkeypatch.delenv("BOX_OPTIMIZER_RATE_SYNC_URL", raising=False)
    rate_path = tmp_path / "Rate and Zone.xlsx"
    zone_key = [[""] for _ in range(37)]
    zone_key[0] = ["HUB"]
    zone_key[1] = ["KG", 0.5, 1, 1.5, 2]
    zone_key[2] = ["Zone USA", 11, 12, 13, 14]
    zone_key[3] = ["Zone 1", 21, 22, 23, 24]
    zone_key[4] = ["Zone 2", 31, 32, 33, 34]
    zone_key[5] = ["Zone 3", 41, 42, 43, 44]
    zone_key[6] = ["Zone 9", "", "", "", ""]
    zone_key[31] = ["Express"]
    zone_key[32] = ["KG", 0.5, 1, 1.5, 2]
    zone_key[33] = ["Zone 1", 101, 111, 121, 131]
    zone_key[34] = ["Zone 2", 201, 211, 221, 231]
    zone_key[35] = ["Zone 3", 301, 311, 321, 331]
    zone_key[36] = ["Zone 4", 401, 411, 421, 431]
    sheet2 = [
        ["", "HUB", "", "", "", "", "", "", "Express", "", "", ""],
        ["", "COUNTRY", "A2 (ISO)", "Ship by", "Zone", "ship by", "ship to contact", "", "COUNTRY", "", "A2 (ISO)", "Zone"],
        ["", "Australia", "AU", "Hub", "Zone 1", "ship to Aetherworks", "kickstarter@aetherworks.com.au", "", "Australia", "澳大利亚", "AU", "Zone 2"],
        ["", "Fallbackia", "FB", "Hub", "Zone 9", "ship to fallback", "", "", "Fallbackia", "回退", "FB", "Zone 2"],
        ["", "", "", "", "", "", "", "", "Brazil", "巴西", "BR", "Zone 2"],
    ]
    _write_xlsx_tables(rate_path, {"Zone Key": zone_key, "Sheet2": sheet2})

    rows = workflow_module._cost_summary_rows(
        [
            {"Country": "Australia", "Chargeable Weight kg": 1.0, "Total Units": 1},
            {"Country": "Brazil", "Chargeable Weight kg": 1.1, "Total Units": 1},
            {"Country": "Fallbackia", "Chargeable Weight kg": 0.6, "Total Units": 3},
            {"Country": "Unknown", "Chargeable Weight kg": 1.0, "Total Units": 1},
        ],
        {**workflow_module.DEFAULT_CONFIG, "rate_sheet_path": str(rate_path)},
    )

    assert rows[0]["Hub Shipping Fee"] == 24
    assert rows[0]["Express"] == 0
    assert rows[0]["Shipping Method"] == "ship to Aetherworks: kickstarter@aetherworks.com.au"
    assert rows[1]["Hub Shipping Fee"] == 0
    assert rows[1]["Express"] == 223
    assert rows[1]["Shipping Method"] == "Ship by Express"
    assert rows[1]["Shipping Rate Note"] == "Express fallback; hub unavailable."
    assert rows[2]["Hub Shipping Fee"] == 0
    assert rows[2]["Express"] == 213.5
    assert rows[2]["Shipping Method"] == "Ship by Express"
    assert rows[3]["Hub Shipping Fee"] == 0
    assert rows[3]["Express"] == 0
    assert rows[3]["Shipping Method"] == "Review Needed"
    assert "No hub or express rate found" in rows[3]["Shipping Rate Note"]
    assert "Shipping method unavailable" in rows[3]["Shipping Rate Note"]
    assert "Zone" not in rows[0]
    assert sum(row["Hub Shipping Fee"] + row["Express"] for row in rows) == 460.5
    assert workflow_module._cost_summary_total_cost(rows) == 460.5
    summary_rows = workflow_module._clean_summary_rows(
        {
            "orders_processed": 4,
            "boxes_created": 4,
            "box_types": 1,
            "unmatched_skus": 0,
            "total_shipping_cost": workflow_module._cost_summary_total_cost(rows),
            "total_shipping_cost_detail": "Hub Shipping Fee + Express",
        },
        [],
        [],
        {},
    )
    total_row = next(row for row in summary_rows if row["Metric"] == "Total Chargeable Cost")
    assert total_row["Value"] == 460.5
    assert total_row["Detail"] == "Hub Shipping Fee + Express"


def test_current_rate_and_zone_format_parses_hub_express_and_ignores_other_services(tmp_path, monkeypatch):
    monkeypatch.setenv("BOX_OPTIMIZER_RATE_SHEET_DIR", str(tmp_path / "managed_rates"))
    monkeypatch.delenv("BOX_OPTIMIZER_RATE_SYNC_URL", raising=False)
    rate_path = tmp_path / "Rate and Zone 2026.xlsx"
    zone_key = [[""] for _ in range(45)]
    zone_key[0] = ["HUB"]
    zone_key[1] = ["KG", 0.5, 1, 1.5, 2]
    zone_key[2] = ["Zone 0", 1, 2, 3, 4]
    zone_key[3] = ["Zone USA", 11, 12, 13, 14]
    zone_key[4] = ["Zone 1", 21, 22, 23, 24]
    zone_key[5] = ["Zone 2", 31, 32, 33, 34]
    zone_key[6] = ["Zone 3", 41, 42, 43, 44]
    zone_key[31] = ["Express"]
    zone_key[32] = ["KG", 0.5, 1, 1.5, 2]
    zone_key[33] = ["Zone 1", 101, 111, 121, 131]
    zone_key[34] = ["Zone 2", 201, 211, 221, 231]
    zone_key[35] = ["Zone 3", 301, 311, 321, 331]
    zone_key[36] = ["Zone 4", 401, 411, 421, 431]
    zone_key[39] = ["Slow Post"]
    zone_key[40] = ["KG", 0.5, 1, 1.5, 2]
    zone_key[41] = ["Zone 2", 999, 999, 999, 999]
    sheet2 = [
        ["", "HUB", "", "", "", "", "", "", "Express", "", "", ""],
        ["", "Country", "A2 ISO", "Ship by", "Zone", "hub destination", "ship-to contact", "", "Country", "Chinese country name", "A2 ISO", "Zone"],
        ["", "Canada", "CA", "Hub", "Zone 2", "ship to hub", "hub@example.com", "", "Brazil", "巴西", "BR", "Zone 2"],
        ["", "Australia", "AU", "Hub", "Zone 1", "ship to Aetherworks", "kickstarter@aetherworks.com.au", "", "Mexico", "墨西哥", "MX", "Zone 3"],
    ]
    _write_xlsx_tables(rate_path, {"Zone Key": zone_key, "Sheet2": sheet2})

    rate_sheet = workflow_module._load_customer_rate_sheet(str(rate_path))
    rows = workflow_module._cost_summary_rows(
        [
            {"Country": "Canada", "Chargeable Weight kg": 1.1, "Total Units": 1},
            {"Country": "Brazil", "Chargeable Weight kg": 1.1, "Total Units": 1},
            {"Country": "Unknown", "Chargeable Weight kg": 1.1, "Total Units": 1},
        ],
        {**workflow_module.DEFAULT_CONFIG, "rate_sheet_path": str(rate_path)},
    )

    assert rate_sheet is not None
    assert rate_sheet.hub.rates_by_zone["Zone 2"][1.5] == 33
    assert rate_sheet.express.rates_by_zone["Zone 2"][1.5] == 221
    assert rate_sheet.hub.rates_by_zone["Zone 2"][1.5] != 999
    assert rate_sheet.hub.zone_by_country["canada"] == "Zone 2"
    assert rate_sheet.hub.zone_by_country["ca"] == "Zone 2"
    assert rate_sheet.hub.iso_by_country["canada"] == "CA"
    assert rate_sheet.hub.ship_by_country["canada"] == "Hub"
    assert rate_sheet.hub.ship_to_instruction_by_country["canada"] == "ship to hub"
    assert rate_sheet.hub.ship_to_contact_by_country["canada"] == "hub@example.com"
    assert rate_sheet.express.zone_by_country["brazil"] == "Zone 2"
    assert rate_sheet.express.zone_by_country["br"] == "Zone 2"
    assert rate_sheet.express.iso_by_country["brazil"] == "BR"
    assert rate_sheet.express.chinese_name_by_country["brazil"] == "巴西"
    assert rows[0]["Hub Shipping Fee"] == 35
    assert rows[0]["Express"] == 0
    assert rows[1]["Hub Shipping Fee"] == 0
    assert rows[1]["Express"] == 223
    assert rows[2]["Hub Shipping Fee"] == 0
    assert rows[2]["Express"] == 0
    assert "No hub or express rate found" in rows[2]["Shipping Rate Note"]


def test_active_uploaded_rate_sheet_is_selected_over_default_rate_sheet(tmp_path, monkeypatch):
    monkeypatch.setenv("BOX_OPTIMIZER_RATE_SHEET_DIR", str(tmp_path / "managed_rates"))
    monkeypatch.delenv("BOX_OPTIMIZER_RATE_SYNC_URL", raising=False)
    default_path = tmp_path / "default-rates.xlsx"
    active_path = active_rate_sheet_path()
    _write_xlsx_tables(default_path, _new_rate_sheet_table(90))
    _write_xlsx_tables(active_path, _new_rate_sheet_table(10))
    _write_active_rate_sheet_metadata_for_test(active_path, "railway-rates.xlsx")

    selection = workflow_module._resolve_customer_rate_sheet({**workflow_module.DEFAULT_CONFIG, "rate_sheet_path": str(default_path)})
    rows = workflow_module._cost_summary_rows(
        [{"Country": "United States", "US State Abbreviation": "CA", "Chargeable Weight kg": 1.0, "Total Units": 1}],
        {**workflow_module.DEFAULT_CONFIG, "rate_sheet_path": str(default_path)},
        selection,
    )

    assert selection.source == "active upload"
    assert selection.filename == "railway-rates.xlsx"
    assert selection.checksum_short == sha256_file(active_path)[:12]
    assert rows[0]["Hub Shipping Fee"] == 12


def test_default_rate_sheet_is_used_when_no_active_upload_exists(tmp_path, monkeypatch):
    monkeypatch.setenv("BOX_OPTIMIZER_RATE_SHEET_DIR", str(tmp_path / "managed_rates"))
    monkeypatch.delenv("BOX_OPTIMIZER_RATE_SYNC_URL", raising=False)
    default_path = tmp_path / "default-rates.xlsx"
    _write_xlsx_tables(default_path, _new_rate_sheet_table(30))

    selection = workflow_module._resolve_customer_rate_sheet({**workflow_module.DEFAULT_CONFIG, "rate_sheet_path": str(default_path)})
    rows = workflow_module._cost_summary_rows(
        [{"Country": "United States", "US State Abbreviation": "CA", "Chargeable Weight kg": 1.0, "Total Units": 1}],
        {**workflow_module.DEFAULT_CONFIG, "rate_sheet_path": str(default_path)},
        selection,
    )

    assert selection.source == "default fallback"
    assert selection.filename == "default-rates.xlsx"
    assert rows[0]["Hub Shipping Fee"] == 32


def test_invalid_active_rate_sheet_falls_back_to_default_and_reports_warning(tmp_path, monkeypatch):
    monkeypatch.setenv("BOX_OPTIMIZER_RATE_SHEET_DIR", str(tmp_path / "managed_rates"))
    monkeypatch.delenv("BOX_OPTIMIZER_RATE_SYNC_URL", raising=False)
    default_path = tmp_path / "default-rates.xlsx"
    active_path = active_rate_sheet_path()
    _write_xlsx_tables(default_path, _new_rate_sheet_table(40))
    active_path.write_bytes(b"not a workbook")
    rate_sheet_metadata_path().write_text(json.dumps({"original_filename": "bad-active.xlsx"}), encoding="utf-8")

    selection = workflow_module._resolve_customer_rate_sheet({**workflow_module.DEFAULT_CONFIG, "rate_sheet_path": str(default_path)})
    rows = workflow_module._cost_summary_rows(
        [{"Country": "United States", "US State Abbreviation": "CA", "Chargeable Weight kg": 1.0, "Total Units": 1}],
        {**workflow_module.DEFAULT_CONFIG, "rate_sheet_path": str(default_path)},
        selection,
    )

    assert selection.source == "default fallback"
    assert selection.filename == "default-rates.xlsx"
    assert "Active rate sheet invalid" in selection.warning
    assert rows[0]["Hub Shipping Fee"] == 42


def test_summary_audit_rows_show_rate_sheet_source_and_checksum(tmp_path, monkeypatch):
    monkeypatch.setenv("BOX_OPTIMIZER_RATE_SHEET_DIR", str(tmp_path / "managed_rates"))
    monkeypatch.delenv("BOX_OPTIMIZER_RATE_SYNC_URL", raising=False)
    active_path = active_rate_sheet_path()
    _write_xlsx_tables(active_path, _new_rate_sheet_table(10))
    _write_active_rate_sheet_metadata_for_test(active_path, "current-railway.xlsx")
    selection = workflow_module._resolve_customer_rate_sheet({})
    summary_rows = workflow_module._clean_summary_rows(
        {
            "orders_processed": 1,
            "boxes_created": 1,
            "box_types": 1,
            "unmatched_skus": 0,
            "rate_sheet_filename": selection.filename,
            "rate_sheet_source": selection.source,
            "rate_sheet_audit_detail": f"Uploaded At: {selection.uploaded_at}; Checksum: {selection.checksum_short}",
            "rate_sheet_warning": selection.warning,
        },
        [],
        [],
        {},
    )

    used_row = next(row for row in summary_rows if row["Metric"] == "Rate Sheet Used")
    source_row = next(row for row in summary_rows if row["Metric"] == "Rate Sheet Source")
    assert used_row["Value"] == "current-railway.xlsx"
    assert "Checksum:" in used_row["Detail"]
    assert source_row["Value"] == "active upload"


def test_summary_audit_rows_show_default_fallback_when_no_active_rate_sheet(tmp_path, monkeypatch):
    monkeypatch.setenv("BOX_OPTIMIZER_RATE_SHEET_DIR", str(tmp_path / "managed_rates"))
    monkeypatch.delenv("BOX_OPTIMIZER_RATE_SYNC_URL", raising=False)
    default_path = tmp_path / "default-rates.xlsx"
    _write_xlsx_tables(default_path, _new_rate_sheet_table(10))
    selection = workflow_module._resolve_customer_rate_sheet({**workflow_module.DEFAULT_CONFIG, "rate_sheet_path": str(default_path)})
    summary_rows = workflow_module._clean_summary_rows(
        {
            "orders_processed": 1,
            "boxes_created": 1,
            "box_types": 1,
            "unmatched_skus": 0,
            "rate_sheet_filename": selection.filename,
            "rate_sheet_source": selection.source,
            "rate_sheet_audit_detail": "",
            "rate_sheet_warning": selection.warning,
        },
        [],
        [],
        {},
    )

    used_row = next(row for row in summary_rows if row["Metric"] == "Rate Sheet Used")
    source_row = next(row for row in summary_rows if row["Metric"] == "Rate Sheet Source")
    assert used_row["Value"] == "default-rates.xlsx"
    assert source_row["Value"] == "default fallback"


def test_rate_sheet_sync_url_missing_skips_remote_sync(tmp_path, monkeypatch):
    monkeypatch.setenv("BOX_OPTIMIZER_RATE_SHEET_DIR", str(tmp_path / "managed_rates"))
    monkeypatch.delenv("BOX_OPTIMIZER_RATE_SYNC_URL", raising=False)
    default_path = tmp_path / "default-rates.xlsx"
    _write_xlsx_tables(default_path, _new_rate_sheet_table(30))

    def fail_if_called(*_args, **_kwargs):
        raise AssertionError("remote sync should not be called")

    monkeypatch.setattr(rate_source_module, "urlopen", fail_if_called)
    selection = workflow_module._resolve_customer_rate_sheet({**workflow_module.DEFAULT_CONFIG, "rate_sheet_path": str(default_path)})

    assert selection.source == "default fallback"
    assert selection.sync_status == ""


def test_rate_sheet_sync_downloads_railway_sheet_when_checksum_differs(tmp_path, monkeypatch):
    monkeypatch.setenv("BOX_OPTIMIZER_RATE_SHEET_DIR", str(tmp_path / "managed_rates"))
    monkeypatch.setenv("BOX_OPTIMIZER_RATE_SYNC_URL", "https://railway.example")
    remote_path = tmp_path / "remote-rates.xlsx"
    _write_xlsx_tables(remote_path, _new_rate_sheet_table(10))
    remote_bytes = remote_path.read_bytes()
    remote_sha = sha256_file(remote_path)
    _install_fake_rate_sync(
        monkeypatch,
        {
            "active": True,
            "filename": "railway-rates.xlsx",
            "uploaded_at": "2026-05-31T12:00:00+00:00",
            "sha256": remote_sha,
            "source": "active_upload",
        },
        workbook_bytes=remote_bytes,
    )

    rows = workflow_module._cost_summary_rows(
        [{"Country": "United States", "US State Abbreviation": "CA", "Chargeable Weight kg": 1.0, "Total Units": 1}],
        {},
    )
    metadata = json.loads(rate_sheet_metadata_path().read_text(encoding="utf-8"))

    assert rows[0]["Hub Shipping Fee"] == 12
    assert active_rate_sheet_path().read_bytes() == remote_bytes
    assert metadata["source"] == "remote_sync"
    assert metadata["original_filename"] == "railway-rates.xlsx"
    assert metadata["sha256"] == remote_sha


def test_rate_sheet_sync_does_nothing_when_checksum_matches(tmp_path, monkeypatch):
    monkeypatch.setenv("BOX_OPTIMIZER_RATE_SHEET_DIR", str(tmp_path / "managed_rates"))
    monkeypatch.setenv("BOX_OPTIMIZER_RATE_SYNC_URL", "https://railway.example")
    active_path = active_rate_sheet_path()
    _write_xlsx_tables(active_path, _new_rate_sheet_table(10))
    _write_active_rate_sheet_metadata_for_test(active_path, "cached-rates.xlsx")
    active_bytes = active_path.read_bytes()
    active_sha = sha256_file(active_path)
    _install_fake_rate_sync(
        monkeypatch,
        {"active": True, "filename": "cached-rates.xlsx", "sha256": active_sha},
        fail_download=True,
    )

    selection = workflow_module._resolve_customer_rate_sheet({})

    assert selection.source == "active upload"
    assert selection.sync_status == "Rate Sheet Sync: up to date from Railway."
    assert active_path.read_bytes() == active_bytes


def test_rate_sheet_sync_uses_cached_active_when_railway_unreachable(tmp_path, monkeypatch):
    monkeypatch.setenv("BOX_OPTIMIZER_RATE_SHEET_DIR", str(tmp_path / "managed_rates"))
    monkeypatch.setenv("BOX_OPTIMIZER_RATE_SYNC_URL", "https://railway.example")
    active_path = active_rate_sheet_path()
    _write_xlsx_tables(active_path, _new_rate_sheet_table(10))
    _write_active_rate_sheet_metadata_for_test(active_path, "cached-rates.xlsx")
    _install_fake_rate_sync(monkeypatch, {}, fail_current=True)

    selection = workflow_module._resolve_customer_rate_sheet({})
    rows = workflow_module._cost_summary_rows(
        [{"Country": "United States", "US State Abbreviation": "CA", "Chargeable Weight kg": 1.0, "Total Units": 1}],
        {},
        selection,
    )

    assert selection.source == "active upload"
    assert "failed" in selection.warning
    assert rows[0]["Hub Shipping Fee"] == 12


def test_rate_sheet_sync_failed_download_does_not_corrupt_cached_active(tmp_path, monkeypatch):
    monkeypatch.setenv("BOX_OPTIMIZER_RATE_SHEET_DIR", str(tmp_path / "managed_rates"))
    monkeypatch.setenv("BOX_OPTIMIZER_RATE_SYNC_URL", "https://railway.example")
    active_path = active_rate_sheet_path()
    remote_path = tmp_path / "remote-rates.xlsx"
    _write_xlsx_tables(active_path, _new_rate_sheet_table(10))
    _write_xlsx_tables(remote_path, _new_rate_sheet_table(40))
    _write_active_rate_sheet_metadata_for_test(active_path, "cached-rates.xlsx")
    cached_bytes = active_path.read_bytes()
    _install_fake_rate_sync(
        monkeypatch,
        {"active": True, "filename": "railway-rates.xlsx", "sha256": sha256_file(remote_path)},
        fail_download=True,
    )

    selection = workflow_module._resolve_customer_rate_sheet({})

    assert selection.source == "active upload"
    assert "failed" in selection.warning
    assert active_path.read_bytes() == cached_bytes


def test_rate_sheet_sync_remote_no_active_keeps_default_behavior(tmp_path, monkeypatch):
    monkeypatch.setenv("BOX_OPTIMIZER_RATE_SHEET_DIR", str(tmp_path / "managed_rates"))
    monkeypatch.setenv("BOX_OPTIMIZER_RATE_SYNC_URL", "https://railway.example")
    default_path = tmp_path / "default-rates.xlsx"
    _write_xlsx_tables(default_path, _new_rate_sheet_table(30))
    _install_fake_rate_sync(monkeypatch, {"active": False})

    selection = workflow_module._resolve_customer_rate_sheet({**workflow_module.DEFAULT_CONFIG, "rate_sheet_path": str(default_path)})
    rows = workflow_module._cost_summary_rows(
        [{"Country": "United States", "US State Abbreviation": "CA", "Chargeable Weight kg": 1.0, "Total Units": 1}],
        {**workflow_module.DEFAULT_CONFIG, "rate_sheet_path": str(default_path)},
        selection,
    )

    assert selection.source == "default fallback"
    assert "Railway has no active rate sheet" in selection.warning
    assert rows[0]["Hub Shipping Fee"] == 32


def test_summary_total_cost_sums_cost_summary_row_totals():
    cost_rows = [
        {"Hub Shipping Fee": 10, "Express": 0},
        {"Hub Shipping Fee": "", "Express": 20.5},
        {"Hub Shipping Fee": None, "Express": ""},
    ]
    cost_total = workflow_module._cost_summary_total_cost(cost_rows)
    summary_rows = workflow_module._clean_summary_rows(
        {
            "orders_processed": 4,
            "boxes_created": 4,
            "box_types": 1,
            "unmatched_skus": 0,
            "total_shipping_cost": cost_total,
            "total_shipping_cost_detail": "Hub Shipping Fee + Express",
        },
        [],
        [],
        {},
    )

    assert workflow_module._customer_handling_fee(1) == 2.0
    assert workflow_module._customer_handling_fee(2) == 2.25
    assert workflow_module._customer_handling_fee(5) == 3.0
    assert workflow_module._customer_handling_fee(1, True) == 1.75
    assert workflow_module._customer_handling_fee(2, True) == 2.25
    assert cost_total == 30.5
    total_row = next(row for row in summary_rows if row["Metric"] == "Total Chargeable Cost")
    assert total_row["Value"] == 30.5
    assert total_row["Detail"] == "Hub Shipping Fee + Express"
    assert not any(row["Metric"] == "Estimated Cost" for row in summary_rows)


def test_cost_summary_applies_narrow_prepacked_handling_discount_inside_shipping_fee(tmp_path, monkeypatch):
    monkeypatch.setenv("BOX_OPTIMIZER_RATE_SHEET_DIR", str(tmp_path / "managed_rates"))
    monkeypatch.delenv("BOX_OPTIMIZER_RATE_SYNC_URL", raising=False)
    lane = workflow_module.CustomerRateLane(rates_by_zone={"Zone 1": {1.0: 10.0}}, zone_by_country={})
    assert workflow_module._rate_lane_shipping_fee(1.0, "Zone 1", lane, 1) == 12.0
    assert workflow_module._rate_lane_shipping_fee(1.0, "Zone 1", lane, 2) == 12.25
    assert workflow_module._rate_lane_shipping_fee(1.0, "Zone 1", lane, 1, True) == 11.75
    assert workflow_module._rate_lane_shipping_fee(1.0, "Zone 1", lane, 2, True) == 12.25

    rows = workflow_module._cost_summary_rows(
        [
            {"Country": "Unknown", "Chargeable Weight kg": 1.0, "Total Units": 1},
            {"Country": "Unknown", "Chargeable Weight kg": 1.0, "Total Units": 2},
            {"Country": "Unknown", "Chargeable Weight kg": 1.0, "Total Units": 1, "Prepacked No Touch": True},
            {"Country": "Unknown", "Chargeable Weight kg": 1.0, "Total Units": 2, "Prepacked No Touch": True},
        ],
        {},
    )

    assert "Prepacked No Touch" not in rows[0]
    assert "Picking/Packing Fee" not in rows[0]
    assert "Total Cost" not in rows[0]


def test_summary_box_types_collapse_cutdown_variants_to_base_vb_box():
    box_rows = [
        {"Box Type": "VB33", "Length cm": 40, "Width cm": 30, "Height cm": 20, "Box Count": 1},
        {"Box Type": "VB33 cut down", "Length cm": 40, "Width cm": 30, "Height cm": 12, "Box Count": 2},
        {"Box Type": "VB33 / cutdown 34 x 32 x 12", "Length cm": 34, "Width cm": 32, "Height cm": 12, "Box Count": 3},
    ]
    summary_rows = workflow_module._clean_summary_rows(
        {
            "orders_processed": 6,
            "boxes_created": 6,
            "box_types": len({workflow_module._base_box_type(row["Box Type"]) for row in box_rows}),
            "unmatched_skus": 0,
        },
        box_rows,
        [],
        {},
    )

    box_type_row = next(row for row in summary_rows if row["Metric"] == "Box Types")
    box_needed_rows = [row for row in summary_rows if row["Section"] == "Boxes Needed"]
    assert box_type_row["Value"] == 1
    assert box_needed_rows == [{"Section": "Boxes Needed", "Metric": "VB 33", "Value": 6, "Detail": "40x30x20 cm"}]


def test_summary_boxes_needed_sorts_by_vb_number_not_alphabetically():
    rows = workflow_module._clean_summary_rows(
        {
            "orders_processed": 8,
            "boxes_created": 8,
            "box_types": 8,
            "unmatched_skus": 0,
        },
        [
            {"Box Type": "VB 18", "Length cm": 1, "Width cm": 1, "Height cm": 1, "Box Count": 1},
            {"Box Type": "VB 39", "Length cm": 1, "Width cm": 1, "Height cm": 1, "Box Count": 1},
            {"Box Type": "VB 4", "Length cm": 1, "Width cm": 1, "Height cm": 1, "Box Count": 1},
            {"Box Type": "VB 23-1", "Length cm": 1, "Width cm": 1, "Height cm": 1, "Box Count": 1},
            {"Box Type": "VB 32", "Length cm": 1, "Width cm": 1, "Height cm": 1, "Box Count": 1},
            {"Box Type": "VB 23", "Length cm": 1, "Width cm": 1, "Height cm": 1, "Box Count": 1},
            {"Box Type": "CUSTOM", "Length cm": 1, "Width cm": 1, "Height cm": 1, "Box Count": 1},
            {"Box Type": "VB6", "Length cm": 1, "Width cm": 1, "Height cm": 1, "Box Count": 1},
        ],
        [],
        {},
    )

    box_metrics = [row["Metric"] for row in rows if row["Section"] == "Boxes Needed"]
    assert box_metrics == ["VB 4", "VB 6", "VB 18", "VB 23", "VB 23-1", "VB 32", "VB 39", "CUSTOM"]
    assert box_metrics.index("VB 4") < box_metrics.index("VB 39")
    assert box_metrics.index("VB 23") < box_metrics.index("VB 23-1") < box_metrics.index("VB 32")


def test_phase_a_workbook_presentation_skeleton_and_campaign_cost_summary(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {
                "SKU": "Core",
                "Product Name": "Core Game",
                "Length": "10",
                "Width": "8",
                "Height": "4",
                "Weight kg": "1",
                "Factory Name": "Longhai Printworks",
                "Publisher Upload Column": "keep me",
            }
        ],
    )
    _write_csv(
        orders_path,
        [{"Order ID": "1", "SKU": "Core", "Quantity": "1", "Backer ID": "B-1", "Name": "Ada"}],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={
            "packing_mode": "fast",
            "preserve_region_sheets": False,
            "campaign": {"name": "Launch Campaign", "code": "LC"},
        },
    )

    workbook = read_workbook(str(output_path))
    sheet_names = _workbook_sheet_names(output_path)
    assert sheet_names[:4] == ["Summary", "Cost Summary - Launch Campaign", "Actual Dimensions", "Labels"]
    for required_sheet in [
        "VFI Intake Form",
        "Optimized to Pack",
        "Label generator",
        "Order Volume Weights",
        "Box Size Summary",
    ]:
        assert required_sheet in sheet_names
    assert "Box Consolidation What-If" in sheet_names
    assert "Debug Summary" in sheet_names
    assert sheet_names.index("Debug Summary") > sheet_names.index("Pledge Combination Summary")

    summary_rows = _sheet_rows(output_path, "Summary")
    summary_sections = {row["Section"] for row in summary_rows}
    summary_metrics = {row["Metric"] for row in summary_rows}
    assert "Run Summary" in summary_sections
    assert "Boxes Needed" in summary_sections
    assert "Warning Count" not in summary_metrics
    assert "Multi-box Orders" not in summary_metrics
    assert "Rules Applied" not in summary_metrics
    assert "Chargeable Weight Plans Selected" not in summary_metrics
    assert "Rules Applied Summary" in summary_sections
    assert "Box Not Available - Substituted Up To VB Box X" not in summary_rows[0]
    assert any(row["Metric"] == "Factory" and row["Value"] == "Longhai Printworks" for row in summary_rows)

    debug_summary_rows = _sheet_rows(output_path, "Debug Summary")
    assert any(row["Metric"] == "Warning Count" for row in debug_summary_rows)

    cost_rows = _sheet_rows(output_path, "Cost Summary - Launch Campaign")
    assert "Shipping fee Hub" not in cost_rows[0]
    assert "Customer Shipping Fee" not in cost_rows[0]
    assert "Hub Shipping Fee (USD)" in cost_rows[0]
    assert "Express (USD)" in cost_rows[0]
    assert "Slow Post" not in cost_rows[0]

    intake_rows = _sheet_rows(output_path, "VFI Intake Form")
    assert intake_rows[0]["Publisher Upload Column"] == "keep me"

    label_generator_rows = _sheet_rows(output_path, "Label generator")
    assert label_generator_rows[0]["Order ID"] == "1"
    assert label_generator_rows[0]["Pledge Configuration"] == "1"
    assert label_generator_rows[0]["Label numbers"] == "1 LC"
    assert "Box Qty" not in label_generator_rows[0]
    assert label_generator_rows[0]["Total Units"] == "1"
    assert label_generator_rows[0]["SKU Breakdown"] == "CORE x1"
    labels_xml = _sheet_xml(output_path, "Labels")
    assert "VFI #" not in labels_xml
    assert "Barcode / QR Value" not in labels_xml
    assert "1 LC" in labels_xml
    assert "Longhai Printworks" in labels_xml
    assert "Launch Campaign" in labels_xml
    assert "Config:" in labels_xml
    assert "Pledge Config" not in labels_xml
    assert "CORE" in labels_xml
    assert "Carton Box" not in labels_xml
    assert "Detailed description of contents: Board Games-of paper and plastic,non-electrical" in labels_xml
    assert "City/State/Zip" in labels_xml
    assert "City/State/Post" not in labels_xml
    assert "<pageSetup" in labels_xml


def test_factory_name_detects_header_or_adjacent_vfi_intake_value():
    assert workflow_module._factory_name_from_vfi_intake_rows(
        [{"Factory Name": "Longhai Printworks", "SKU": "CORE"}]
    ) == "Longhai Printworks"
    assert workflow_module._factory_name_from_vfi_intake_rows(
        [{"Field": "Factory", "Value": "Longhai Printworks"}]
    ) == "Longhai Printworks"
    assert workflow_module._factory_name_from_vfi_intake_rows(
        [{"SKU": "CORE"}]
    ) == ""


def test_intake_client_invoice_metadata_detects_raw_cells_far_to_the_right(tmp_path):
    sku_master_path = tmp_path / "sku_master.xlsx"
    _write_xlsx_table(
        sku_master_path,
        "stock",
        [
            _wide_row(43, {0: "SKU", 1: "Item name", 2: "Weight/g", 3: "L-cm", 4: "W-cm", 5: "H-cm"}),
            _wide_row(43, {0: "CORE", 1: "Core Game", 2: "1000", 3: "10", 4: "8", 5: "4"}),
            _wide_row(
                43,
                {
                    8: "Address Line 1:",
                    10: "123 Client Street",
                    11: "FACTORY",
                    12: "WHATZ",
                    13: "VFI USE",
                    14: "Internal review only",
                    15: "Inbound Fee:",
                    16: "$12.50",
                    36: "CAMPAIGN NAME:",
                    38: "Far Right Campaign",
                    39: "Commodity:",
                    40: "Board Games",
                    41: "Accounting EMAIL:",
                    42: "accounting@example.com",
                },
            ),
        ],
    )

    source = read_workbook(str(sku_master_path))[0]
    metadata = source.metadata["intake_summary_metadata"]

    assert metadata["Campaign Name"] == "Far Right Campaign"
    assert metadata["Commodity"] == "Board Games"
    assert metadata["Address Line 1"] == "123 Client Street"
    assert metadata["Factory"] == "WHATZ"
    assert metadata["VFI Use"] == "Internal review only"
    assert metadata["Inbound Fee"] == "$12.50"
    assert metadata["Accounting Email"] == "accounting@example.com"
    assert source.metadata["factory_name"] == "WHATZ"


def test_intake_client_invoice_metadata_appears_in_summary_before_factory(tmp_path):
    sku_master_path = tmp_path / "sku_master.xlsx"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_xlsx_table(
        sku_master_path,
        "stock",
        [
            _wide_row(44, {0: "SKU", 1: "Item name", 2: "Weight/g", 3: "L-cm", 4: "W-cm", 5: "H-cm"}),
            _wide_row(44, {0: "CORE", 1: "Core Game", 2: "1000", 3: "10", 4: "8", 5: "4"}),
            _wide_row(
                44,
                {
                    7: "Country:",
                    8: "United States",
                    9: "VAT/EORI/TAX ID:",
                    10: "VAT-123",
                    11: "EMAIL #3:",
                    36: "INVOICES TO:",
                    37: "Client Finance",
                    38: "EMAIL #2:",
                    39: "finance2@example.com",
                    40: "Additional Information:",
                    41: "Use PO 456",
                    42: "FACTORY",
                    43: "WHATZ",
                },
            ),
        ],
    )
    _write_csv(
        orders_path,
        [{"Order ID": "1", "SKU": "CORE", "Quantity": "1", "Backer ID": "B-1", "Name": "Ada"}],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={"packing_mode": "fast", "preserve_region_sheets": False},
    )

    run_summary_rows = [row for row in _sheet_rows(output_path, "Summary") if row["Section"] == "Run Summary"]
    metrics = [row["Metric"] for row in run_summary_rows]
    values = {row["Metric"]: row["Value"] for row in run_summary_rows}

    assert metrics.index("Invoices To") < metrics.index("Email #2") < metrics.index("Country")
    assert metrics.index("Additional Information") < metrics.index("Factory")
    assert "Email #3" not in metrics
    assert values["Invoices To"] == "Client Finance"
    assert values["Email #2"] == "finance2@example.com"
    assert values["Country"] == "United States"
    assert values["VAT/EORI/TAX ID"] == "VAT-123"
    assert values["Additional Information"] == "Use PO 456"
    assert values["Factory"] == "WHATZ"

    optimized_rows = _sheet_rows(output_path, "Optimized to Pack")
    assert optimized_rows[0]["Total Pledges"] == "1"
    assert "CORE x1" in optimized_rows[0]["All Items"]


def test_fast_production_workbook_keeps_operational_sheets_and_internal_calculations(tmp_path):
    sku_master_path = tmp_path / "sku_master.xlsx"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_xlsx_table(
        sku_master_path,
        "stock",
        [
            _wide_row(44, {0: "SKU", 1: "Item name", 2: "Weight/g", 3: "L-cm", 4: "W-cm", 5: "H-cm"}),
            _wide_row(44, {0: "CORE", 1: "Core Game", 2: "1000", 3: "10", 4: "8", 5: "4"}),
            _wide_row(
                44,
                {
                    7: "Country:",
                    8: "United States",
                    9: "VAT/EORI/TAX ID:",
                    10: "VAT-123",
                    36: "INVOICES TO:",
                    37: "Client Finance",
                    40: "Additional Information:",
                    41: "Use PO 456",
                    42: "FACTORY",
                    43: "WHATZ",
                },
            ),
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "1", "SKU": "CORE", "Quantity": "1", "Backer ID": "B-1", "Name": "Ada", "Country": "United States"},
            {"Order ID": "2", "SKU": "CORE", "Quantity": "1", "Backer ID": "B-2", "Name": "Grace", "Country": "Canada"},
        ],
    )

    summary = optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={
            "packing_mode": "fast",
            "preserve_region_sheets": False,
            "workbook_output_mode": "fast_production",
        },
    )

    sheet_names = _workbook_sheet_names(output_path)
    assert sheet_names[:4] == ["Summary", "Cost Summary", "Actual Dimensions", "Labels"]
    for required_sheet in ["United States", "Canada", "VFI Intake Form", "Optimized to Pack", "Box Size Summary"]:
        assert required_sheet in sheet_names
    for skipped_sheet in [
        "Label generator",
        "Order Volume Weights",
        "Pledge Combination Summary",
        "Packing Detail",
        "Multi Box Detail",
        "Debug Summary",
        "Input Column Mapping",
    ]:
        assert skipped_sheet not in sheet_names
    assert "Errors and Warnings" in sheet_names

    assert summary["workbook_output_mode"] == "fast_production"
    assert summary["country_sheet_count"] == 2
    assert "Label generator" in summary["sheets_skipped"]
    assert "Order Volume Weights" in summary["sheets_skipped"]
    assert summary["qr_images_written"] == 2
    assert summary["boxes_created"] == 2

    summary_rows = _sheet_rows(output_path, "Summary")
    run_summary = {row["Metric"]: row["Value"] for row in summary_rows if row["Section"] == "Run Summary"}
    assert run_summary["Output Mode"] == "fast_production"
    assert run_summary["Invoices To"] == "Client Finance"
    assert run_summary["Factory"] == "WHATZ"
    assert any(row.get("SKU") == "SKU Intake Summary" for row in summary_rows)

    cost_rows = _sheet_rows(output_path, "Cost Summary")
    assert len(cost_rows) == 2
    assert {row["VFI #"] for row in cost_rows} == {"1 VFI", "2 VFI"}
    labels_xml = _sheet_xml(output_path, "Labels")
    assert "B-1" in labels_xml
    assert "B-2" in labels_xml


def test_adjacent_factory_metadata_appears_in_labels_header(tmp_path):
    sku_master_path = tmp_path / "sku_master.xlsx"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_xlsx(
        sku_master_path,
        "stock",
        [
            {"SKU": "CORE", "Item name": "Core Game", "Weight/g": "1000", "L-cm": "10", "W-cm": "8", "H-cm": "4", "Field": "Factory", "Value": "WHATZ"}
        ],
    )
    _write_csv(
        orders_path,
        [{"Order ID": "1", "SKU": "CORE", "Quantity": "1", "Backer ID": "B-1", "Name": "Ada"}],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={
            "packing_mode": "fast",
            "preserve_region_sheets": False,
            "campaign": {"name": "Launch Campaign", "code": "LC"},
        },
    )

    summary_rows = _sheet_rows(output_path, "Summary")
    labels_xml = _sheet_xml(output_path, "Labels")

    assert any(row["Metric"] == "Factory" and row["Value"] == "WHATZ" for row in summary_rows)
    assert 'r="C1" t="inlineStr" s="24"><is><t></t></is></c>' in labels_xml
    assert "WHATZ" in labels_xml


def test_vfi_intake_form_preserves_blank_header_factory_columns(tmp_path):
    sku_master_path = tmp_path / "sku_master.xlsx"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_xlsx_table(
        sku_master_path,
        "stock",
        [
            ["SKU", "Item name", "Weight/g", "L-cm", "W-cm", "H-cm", "", ""],
            ["SKU编号", "商品全名", "重量(kg)", "长", "宽", "高", "", ""],
            ["CORE", "Core Game", "1000", "10", "8", "4", "", ""],
            ["", "", "", "", "", "", "Factory", "WHATZ"],
        ],
    )
    _write_csv(
        orders_path,
        [{"Order ID": "1", "SKU": "CORE", "Quantity": "1", "Backer ID": "B-1", "Name": "Ada"}],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={"packing_mode": "fast", "preserve_region_sheets": False},
    )

    intake_rows = _sheet_rows(output_path, "VFI Intake Form")
    summary_rows = _sheet_rows(output_path, "Summary")

    assert intake_rows[0]["Column G"] == ""
    assert intake_rows[2]["Column G"] == "Factory"
    assert intake_rows[2]["Column H"] == "WHATZ"
    assert any(row["Metric"] == "Factory" and row["Value"] == "WHATZ" for row in summary_rows)


def test_summary_boxes_needed_omits_backup_vendor_box_recommendation_column():
    rows = workflow_module._clean_summary_rows(
        {
            "orders_processed": 1,
            "boxes_created": 1,
            "box_types": 1,
            "unmatched_skus": 0,
        },
        [
            {
                "Box Type": "VB 15 cutdown",
                "Length cm": 35.4,
                "Width cm": 32.4,
                "Height cm": 12,
                "Box Count": 1,
                "Backup Vendor Box": "VB9 / 34x34x12",
            }
        ],
        [],
        {},
    )

    box_row = next(row for row in rows if row["Section"] == "Boxes Needed")
    assert box_row["Metric"] == "VB 15"
    assert "Box Not Available - Substituted Up To VB Box X" not in box_row


def test_summary_boxes_needed_has_no_substitute_column_when_no_valid_backup_exists():
    rows = workflow_module._clean_summary_rows(
        {
            "orders_processed": 1,
            "boxes_created": 1,
            "box_types": 1,
            "unmatched_skus": 0,
        },
        [
            {
                "Box Type": "VB 15",
                "Length cm": 35.4,
                "Width cm": 32.4,
                "Height cm": 12,
                "Box Count": 1,
                "Backup Vendor Box": "N/A",
            }
        ],
        [],
        {},
    )

    box_row = next(row for row in rows if row["Section"] == "Boxes Needed")
    assert "Box Not Available - Substituted Up To VB Box X" not in box_row


def test_box_consolidation_what_if_prefers_used_valid_backup_for_low_volume_box():
    rows = workflow_module._box_consolidation_what_if_rows(
        [
            {
                "Box Type": "VB 34",
                "Backup Vendor Box": "VB35 / 50x40x27.25",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 9,
            },
            {
                "Box Type": "VB 35",
                "Backup Vendor Box": "N/A",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 12,
            },
            {
                "Box Type": "VB 35",
                "Backup Vendor Box": "N/A",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 12,
            },
            {
                "Box Type": "VB 35",
                "Backup Vendor Box": "N/A",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 12,
            },
            {
                "Box Type": "VB 35",
                "Backup Vendor Box": "N/A",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 12,
            },
            {
                "Box Type": "VB 35",
                "Backup Vendor Box": "N/A",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 12,
            },
        ]
    )

    recommended = next(row for row in rows if row["Original VB Box"] == "VB 34")
    assert recommended["Backup VB Box Candidate"] == "VB 35"
    assert recommended["Backup Box Current Quantity"] == 5
    assert recommended["Quantity Proposed to Move"] == 1
    assert recommended["Original Box Quantity After Move"] == 0
    assert recommended["Backup Box Quantity After Move"] == 6
    assert recommended["Recommendation"] == "Recommended"
    assert recommended["Chain Path"] == "VB 34 -> VB 35"
    assert recommended["Final Target VB Box"] == "VB 35"
    assert recommended["Final Fit Validated"] == "Yes"
    assert recommended["Box Type Reduction Impact"] == 1
    assert recommended["Original Box Type Count"] == 2
    assert recommended["Hypothetical Box Type Count"] == 1
    assert recommended["Chargeable Weight Increase per Package"] == 1.9
    assert recommended["Total Chargeable Weight Increase"] == 1.9


def test_box_consolidation_what_if_reports_no_valid_backup_and_does_not_move_high_volume_box():
    rows = workflow_module._box_consolidation_what_if_rows(
        [
            {
                "Box Type": "VB 10",
                "Backup Vendor Box": "N/A",
                "Packed Actual Weight kg": 2,
                "Chargeable Weight kg": 3,
            },
            {
                "Box Type": "VB 20",
                "Backup Vendor Box": "VB30 / 60x40x30",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 10,
            },
            {
                "Box Type": "VB 20",
                "Backup Vendor Box": "VB30 / 60x40x30",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 10,
            },
        ]
    )

    no_backup = next(row for row in rows if row["Original VB Box"] == "VB 10")
    assert no_backup["Recommendation"] == "Rejected"
    assert no_backup["Backup VB Box Candidate"] == "N/A"
    assert no_backup["Final Fit Validated"] == "No"

    high_volume = next(row for row in rows if row["Original VB Box"] == "VB 20")
    assert high_volume["Recommendation"] == "Rejected"
    assert high_volume["Quantity Proposed to Move"] == 0
    assert high_volume["Box Type Reduction Impact"] == 0
    assert high_volume["Final Fit Validated"] == "No"
    assert high_volume["Reason Accepted / Rejected"] == "Skipped: protected high-volume primary box."


def test_box_consolidation_what_if_rejects_chargeable_increase_at_or_above_two_kg():
    rows = workflow_module._box_consolidation_what_if_rows(
        [
            {
                "Box Type": "VB 10",
                "Backup Vendor Box": "VB30 / 50x40x27.5",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 9,
            },
            {
                "Box Type": "VB 11",
                "Backup Vendor Box": "VB30 / 50x40x27.75",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 9,
            },
            {
                "Box Type": "VB 30",
                "Backup Vendor Box": "N/A",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 11,
            },
            {
                "Box Type": "VB 30",
                "Backup Vendor Box": "N/A",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 11,
            },
            {
                "Box Type": "VB 30",
                "Backup Vendor Box": "N/A",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 11,
            },
            {
                "Box Type": "VB 30",
                "Backup Vendor Box": "N/A",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 11,
            },
            {
                "Box Type": "VB 30",
                "Backup Vendor Box": "N/A",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 11,
            },
        ]
    )

    two_kg = next(row for row in rows if row["Original VB Box"] == "VB 10")
    assert two_kg["Chargeable Weight Increase per Package"] == 2.0
    assert two_kg["Recommendation"] == "Rejected"
    assert two_kg["Quantity Proposed to Move"] == 0
    assert two_kg["Reason Accepted / Rejected"] == "Rejected: total chargeable increase >= 2.0 kg."

    over_two_kg = next(row for row in rows if row["Original VB Box"] == "VB 11")
    assert over_two_kg["Chargeable Weight Increase per Package"] == 2.1
    assert over_two_kg["Recommendation"] == "Rejected"
    assert over_two_kg["Reason Accepted / Rejected"] == "Rejected: total chargeable increase >= 2.0 kg."


def test_box_consolidation_what_if_protects_top_quarter_high_volume_sources_and_ties():
    protected = workflow_module._protected_high_volume_boxes(
        Counter(
            {
                "VB 33": 72,
                "VB 39": 41,
                "VB 47": 41,
                "VB 32": 12,
                "VB 23": 4,
                "VB 12": 3,
                "VB 8": 2,
                "VB 3": 1,
            }
        )
    )

    assert protected == {"VB 33", "VB 39", "VB 47"}


def test_box_consolidation_what_if_allows_low_volume_source_to_move_into_protected_target():
    rows = workflow_module._box_consolidation_what_if_rows(
        [
            {
                "Box Type": "VB 12",
                "Backup Vendor Box": "VB33 / 50x40x27.25",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 9,
            },
            *[
                {
                    "Box Type": "VB 33",
                    "Backup Vendor Box": "VB47 / 50x40x29",
                    "Packed Actual Weight kg": 4,
                    "Chargeable Weight kg": 12,
                }
                for _ in range(8)
            ],
            *[
                {
                    "Box Type": f"VB {box_id}",
                    "Backup Vendor Box": "N/A",
                    "Packed Actual Weight kg": 4,
                    "Chargeable Weight kg": 13,
                }
                for box_id in [47] * 7 + [48] * 6 + [49] * 5 + [50] * 4 + [51] * 3 + [52] * 2
            ],
        ]
    )

    low_volume = next(row for row in rows if row["Original VB Box"] == "VB 12")
    protected_source = next(row for row in rows if row["Original VB Box"] == "VB 33")
    assert low_volume["Recommendation"] == "Recommended"
    assert low_volume["Final Target VB Box"] == "VB 33"
    assert protected_source["Recommendation"] == "Rejected"
    assert protected_source["Reason Accepted / Rejected"] == "Skipped: protected high-volume primary box."


def test_box_consolidation_what_if_accepts_chained_substitution_when_total_increase_stays_under_cap():
    rows = workflow_module._box_consolidation_what_if_rows(
        [
            {
                "Box Type": "VB 23",
                "Backup Vendor Box": "VB23.1 / 50x40x25",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 9,
            },
            {
                "Box Type": "VB 23.1",
                "Backup Vendor Box": "VB32 / 50x40x27.25",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 10,
            },
            {
                "Box Type": "VB 32",
                "Backup Vendor Box": "N/A",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 12,
            },
            {
                "Box Type": "VB 32",
                "Backup Vendor Box": "N/A",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 12,
            },
            {
                "Box Type": "VB 32",
                "Backup Vendor Box": "N/A",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 12,
            },
        ]
    )

    first_step = next(row for row in rows if row["Original VB Box"] == "VB 23")
    assert first_step["Backup VB Box Candidate"] == "VB 32"
    assert first_step["Chain Path"] == "VB 23 -> VB 23.1 -> VB 32"
    assert first_step["Final Target VB Box"] == "VB 32"
    assert first_step["Recommendation"] == "Recommended"
    assert first_step["Quantity Proposed to Move"] == 1
    assert first_step["Chargeable Weight Increase per Package"] == 1.9

    chained_step = next(row for row in rows if row["Original VB Box"] == "VB 23.1")
    assert chained_step["Backup VB Box Candidate"] == "VB 32"
    assert chained_step["Recommendation"] == "Rejected"
    assert chained_step["Quantity Proposed to Move"] == 0
    assert chained_step["Box Type Reduction Impact"] == 0
    assert chained_step["Reason Accepted / Rejected"] == "Rejected: conflicts with accepted consolidation path."


def test_box_consolidation_what_if_rejects_chained_substitution_when_total_increase_hits_cap():
    rows = workflow_module._box_consolidation_what_if_rows(
        [
            {
                "Box Type": "VB 23",
                "Backup Vendor Box": "VB23.1 / 50x40x25",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 9,
            },
            {
                "Box Type": "VB 23.1",
                "Backup Vendor Box": "VB32 / 50x40x29",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 10,
            },
            {
                "Box Type": "VB 32",
                "Backup Vendor Box": "N/A",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 12,
            },
            {
                "Box Type": "VB 32",
                "Backup Vendor Box": "N/A",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 12,
            },
            {
                "Box Type": "VB 32",
                "Backup Vendor Box": "N/A",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 12,
            },
        ]
    )

    rejected = next(row for row in rows if row["Original VB Box"] == "VB 23")
    assert rejected["Chain Path"] == "VB 23 -> VB 23.1"
    assert rejected["Final Target VB Box"] == "VB 23.1"
    assert rejected["Recommendation"] == "Recommended"
    assert rejected["Chargeable Weight Increase per Package"] == 1.0
    assert rejected["Reason Accepted / Rejected"] == "Accepted one-step fallback; second step rejected: chargeable increase >= 2.0 kg."


def test_box_consolidation_what_if_rejects_chain_when_final_target_cannot_contain_direct_backup():
    rows = workflow_module._box_consolidation_what_if_rows(
        [
            {
                "Box Type": "VB 23",
                "Backup Vendor Box": "VB23.1 / 50x40x25",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 10,
            },
            {
                "Box Type": "VB 23.1",
                "Backup Vendor Box": "VB32 / 50x35x25",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 10,
            },
            {
                "Box Type": "VB 32",
                "Backup Vendor Box": "N/A",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 11,
            },
            {
                "Box Type": "VB 32",
                "Backup Vendor Box": "N/A",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 11,
            },
            {
                "Box Type": "VB 32",
                "Backup Vendor Box": "N/A",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 11,
            },
        ]
    )

    rejected = next(row for row in rows if row["Original VB Box"] == "VB 23")
    assert rejected["Chain Path"] == "VB 23 -> VB 23.1"
    assert rejected["Final Target VB Box"] == "VB 23.1"
    assert rejected["Recommendation"] == "Recommended"
    assert rejected["Final Fit Validated"] == "Yes"
    assert rejected["Reason Accepted / Rejected"] == "Accepted one-step fallback; second step rejected: final target not valid for original package."


def test_box_consolidation_what_if_does_not_chain_beyond_two_steps():
    rows = workflow_module._box_consolidation_what_if_rows(
        [
            {
                "Box Type": "VB 23",
                "Backup Vendor Box": "VB23.1 / 50x40x25",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 9,
            },
            {
                "Box Type": "VB 23.1",
                "Backup Vendor Box": "VB32 / 50x40x27.25",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 10,
            },
            {
                "Box Type": "VB 32",
                "Backup Vendor Box": "VB39 / 50x40x28",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 11,
            },
            {
                "Box Type": "VB 32",
                "Backup Vendor Box": "VB39 / 50x40x28",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 11,
            },
            {
                "Box Type": "VB 32",
                "Backup Vendor Box": "VB39 / 50x40x28",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 11,
            },
            {
                "Box Type": "VB 39",
                "Backup Vendor Box": "VB47 / 50x40x29",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 12,
            },
            {
                "Box Type": "VB 39",
                "Backup Vendor Box": "VB47 / 50x40x29",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 12,
            },
            {
                "Box Type": "VB 47",
                "Backup Vendor Box": "N/A",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 13,
            },
            {
                "Box Type": "VB 47",
                "Backup Vendor Box": "N/A",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 13,
            },
        ]
    )

    recommended = next(row for row in rows if row["Original VB Box"] == "VB 23")
    assert recommended["Recommendation"] == "Recommended"
    assert recommended["Chain Path"] == "VB 23 -> VB 23.1 -> VB 32"
    assert recommended["Chain Path"].count("->") == 2
    assert "VB 39" not in recommended["Chain Path"]


def test_accepted_box_consolidation_updates_summary_and_labels_without_changing_original_rows(tmp_path, monkeypatch):
    monkeypatch.setenv("BOX_OPTIMIZER_RATE_SHEET_DIR", str(tmp_path / "managed_rates"))
    monkeypatch.delenv("BOX_OPTIMIZER_RATE_SYNC_URL", raising=False)
    box_rows = [
        {
            "Box Type": "VB 23",
            "Backup Vendor Box": "VB32 / 50x40x27.25",
            "Packed Actual Weight kg": 4,
            "Chargeable Weight kg": 9,
            "Chargeable Weight g": 9000,
            "Length cm": 45,
            "Width cm": 35,
            "Height cm": 25,
            "Box Count": 1,
            "SKU Breakdown": "CORE x1",
            "Label SKUs in Box": "CORE x1",
            "Total Units": 1,
            "Unit Count": 1,
            "Label Unit Count": 1,
            "Order ID": "1",
            "Box Number": 1,
            "Box Qty": 1,
            "Country": "United States",
            "Box Standardization Note": "",
        },
        *[
            {
                "Box Type": "VB 32",
                "Backup Vendor Box": "N/A",
                "Packed Actual Weight kg": 4,
                "Chargeable Weight kg": 12,
                "Chargeable Weight g": 12000,
                "Length cm": 50,
                "Width cm": 40,
                "Height cm": 27.25,
                "Box Count": 1,
                "SKU Breakdown": "EXP x1",
                "Label SKUs in Box": "EXP x1",
                "Total Units": 1,
                "Unit Count": 1,
                "Label Unit Count": 1,
                "Order ID": str(index),
                "Box Number": 1,
                "Box Qty": 1,
                "Country": "United States",
                "Box Standardization Note": "",
            }
            for index in range(2, 5)
        ],
    ]

    what_if_rows = workflow_module._box_consolidation_what_if_rows(box_rows)
    operational_rows = workflow_module._box_rows_with_operational_consolidation(box_rows, what_if_rows)
    summary_rows = workflow_module._clean_summary_rows(
        {"orders_processed": 4, "boxes_created": 4, "box_types": 1, "unmatched_skus": 0},
        workflow_module._box_size_summary(operational_rows),
        [],
        {},
    )
    label_rows = workflow_module._label_generator_rows([operational_rows[0]], {"CORE x1": 1}, "TEST")
    cost_rows = workflow_module._cost_summary_rows(
        [
            {
                "Country": "United States",
                "US State Abbreviation": "CA",
                "Chargeable Weight kg": operational_rows[0]["Chargeable Weight kg"],
                "Chargeable Weight g": operational_rows[0]["Chargeable Weight g"],
                "Total Units": operational_rows[0]["Total Units"],
            }
        ],
        {},
    )

    assert box_rows[0]["Box Type"] == "VB 23"
    assert what_if_rows[0]["Workbook Presentation Applied"] == "Yes: Summary, Labels, Cost Summary"
    assert operational_rows[0]["Box Type"] == "VB 32"
    assert operational_rows[0]["Length cm"] == 50
    assert operational_rows[0]["Chargeable Weight kg"] == 10.9
    assert operational_rows[0]["Chargeable Weight g"] == 10900
    box_summary = [row for row in summary_rows if row["Section"] == "Boxes Needed"]
    assert box_summary == [{"Section": "Boxes Needed", "Metric": "VB 32", "Value": 4, "Detail": "50.0x40.0x27.25 cm"}]
    assert "Box Not Available - Substituted Up To VB Box X" not in summary_rows[0]
    assert label_rows[0]["Box Plan"] == "VB 32"
    assert cost_rows[0]["Chargeable Weight kg"] == 10.9
    assert cost_rows[0]["Chargeable Weight g"] == 10900


def test_all_quantity_one_quote_safety_skips_box_consolidation():
    box_rows = [
        {
            "Box Type": "VB 23",
            "Backup Vendor Box": "VB32 / 50x40x27.25",
            "Packed Actual Weight kg": 4,
            "Chargeable Weight kg": 9,
            "Chargeable Weight g": 9000,
            "Length cm": 45,
            "Width cm": 35,
            "Height cm": 25,
            "Box Count": 1,
            "SKU Breakdown": "CORE x1",
            "Label SKUs in Box": "CORE x1",
            "Total Units": 1,
            "Unit Count": 1,
            "Label Unit Count": 1,
            "Order ID": "1",
            "Box Number": 1,
            "Box Qty": 1,
            "Country": "United States",
            "Box Standardization Note": "",
        },
        {
            "Box Type": "VB 32",
            "Backup Vendor Box": "N/A",
            "Packed Actual Weight kg": 4,
            "Chargeable Weight kg": 12,
            "Chargeable Weight g": 12000,
            "Length cm": 50,
            "Width cm": 40,
            "Height cm": 27.25,
            "Box Count": 1,
            "SKU Breakdown": "EXP x1",
            "Label SKUs in Box": "EXP x1",
            "Total Units": 1,
            "Unit Count": 1,
            "Label Unit Count": 1,
            "Order ID": "2",
            "Box Number": 1,
            "Box Qty": 1,
            "Country": "United States",
            "Box Standardization Note": "",
        },
    ]

    assert workflow_module._all_configurations_quantity_one(box_rows)
    what_if_rows = workflow_module._box_consolidation_what_if_rows(
        box_rows,
        skip_reason="Skipped: all configurations quantity 1; using best-fit boxes.",
    )
    operational_rows = workflow_module._box_rows_with_operational_consolidation(box_rows, what_if_rows)
    summary_rows = workflow_module._clean_summary_rows(
        {"orders_processed": 2, "boxes_created": 2, "box_types": 2, "unmatched_skus": 0},
        workflow_module._box_size_summary(operational_rows),
        [],
        {},
    )
    label_rows = workflow_module._label_generator_rows([operational_rows[0]], {"CORE x1": 1}, "TEST")

    skipped = next(row for row in what_if_rows if row["Original VB Box"] == "VB 23")
    assert skipped["Reason Accepted / Rejected"] == "Skipped: all configurations quantity 1; using best-fit boxes."
    assert skipped["Workbook Presentation Applied"] == "No"
    assert operational_rows[0]["Box Type"] == "VB 23"
    assert operational_rows[0]["Chargeable Weight kg"] == 9
    assert {row["Metric"] for row in summary_rows if row["Section"] == "Boxes Needed"} == {"VB 23", "VB 32"}
    assert label_rows[0]["Box Plan"] == "VB 23"


def test_all_quantity_one_quote_safety_detects_repeated_configurations_as_eligible():
    box_rows = [
        {"Order ID": "1", "SKU Breakdown": "CORE x1", "Box Number": 1},
        {"Order ID": "2", "SKU Breakdown": "CORE x1", "Box Number": 1},
    ]

    assert not workflow_module._all_configurations_quantity_one(box_rows)


def test_phase_a_cost_summary_falls_back_when_campaign_name_missing(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [{"SKU": "Core", "Product Name": "Core", "Length": "10", "Width": "8", "Height": "4", "Weight kg": "1"}],
    )
    _write_csv(orders_path, [{"Order ID": "1", "SKU": "Core", "Quantity": "1"}])

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={"packing_mode": "fast", "preserve_region_sheets": False, "campaign": {"code": "CODEONLY"}},
    )

    sheet_names = [sheet.sheet_name for sheet in read_workbook(str(output_path))]
    assert sheet_names[1] == "Cost Summary"
    assert "Cost Summary - CODEONLY" not in sheet_names



def test_phase_b_vfi_numbers_and_cost_summary_follow_final_label_order(tmp_path):
    sku_master_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.csv"
    output_path = tmp_path / "optimized.xlsx"
    _write_csv(
        sku_master_path,
        [
            {"SKU": "A", "Product Name": "Alpha", "Length": "10", "Width": "8", "Height": "4", "Weight kg": "1"},
            {"SKU": "B", "Product Name": "Beta", "Length": "9", "Width": "7", "Height": "3", "Weight kg": "1"},
        ],
    )
    _write_csv(
        orders_path,
        [
            {"Order ID": "1", "SKU": "A", "Quantity": "1", "Backer ID": "B-1", "Name": "One"},
            {"Order ID": "2", "SKU": "B", "Quantity": "1", "Backer ID": "B-2", "Name": "Two"},
            {"Order ID": "3", "SKU": "B", "Quantity": "1", "Backer ID": "B-3", "Name": "Three"},
        ],
    )

    optimize_workbook(
        str(sku_master_path),
        str(orders_path),
        str(output_path),
        config={
            "packing_mode": "fast",
            "preserve_region_sheets": False,
            "campaign": {"name": "Launch Campaign", "code": "LC"},
        },
    )

    optimized_rows = _sheet_rows(output_path, "Optimized to Pack")
    assert optimized_rows[0]["Total Pledges"] == "2"
    assert "B x1" in optimized_rows[0]["All Items"]

    order_rows = _sheet_rows(output_path, "Order Volume Weights")
    assert [row["Order ID"] for row in order_rows] == ["2", "3", "1"]
    assert [row["VFI #"] for row in order_rows] == ["1 LC", "2 LC", "3 LC"]

    cost_rows = _sheet_rows(output_path, "Cost Summary - Launch Campaign")
    assert [row["Backer ID"] for row in cost_rows] == ["B-2", "B-3", "B-1"]
    assert [row["VFI #"] for row in cost_rows] == ["1 LC", "2 LC", "3 LC"]

    label_rows = _sheet_rows(output_path, "Label generator")
    assert [row["Order ID"] for row in label_rows] == ["2", "3", "1"]
    assert "VFI #" not in label_rows[0]
    assert [row["Pledge Configuration"] for row in label_rows] == ["1", "1", "2"]
    assert [row["Label numbers"] for row in label_rows] == ["1 LC", "2 LC", "3 LC"]
    assert "Box Qty" not in label_rows[0]
