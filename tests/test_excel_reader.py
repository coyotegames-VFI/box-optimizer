import csv
import zipfile

from box_optimizer.io.excel_reader import read_intake, read_orders, read_sku_master, read_workbook


def _write_csv(path, rows):
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


def _write_xlsx(path, sheets):
    workbook_sheets = []
    relationships = []
    sheet_files = {}

    for index, (sheet_name, rows) in enumerate(sheets.items(), start=1):
        rel_id = f"rId{index}"
        workbook_sheets.append(
            f'<sheet name="{sheet_name}" sheetId="{index}" r:id="{rel_id}"/>'
        )
        relationships.append(
            f'<Relationship Id="{rel_id}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
        )

        xml_rows = []
        for row_number, row_values in enumerate(rows, start=1):
            cells = [
                _inline_cell(chr(ord("A") + column), row_number, value)
                for column, value in enumerate(row_values)
            ]
            xml_rows.append(f'<row r="{row_number}">{"".join(cells)}</row>')
        sheet_files[f"xl/worksheets/sheet{index}.xml"] = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<sheetData>{"".join(xml_rows)}</sheetData>'
            "</worksheet>"
        )

    workbook = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheets>{"".join(workbook_sheets)}</sheets>'
        "</workbook>"
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f'{"".join(relationships)}'
        "</Relationships>"
    )

    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("xl/workbook.xml", workbook)
        archive.writestr("xl/_rels/workbook.xml.rels", rels)
        for sheet_path, sheet_xml in sheet_files.items():
            archive.writestr(sheet_path, sheet_xml)


def test_read_sku_master_from_csv_infers_columns_and_preserves_metadata(tmp_path):
    path = tmp_path / "sku_master.csv"
    _write_csv(
        path,
        [
            {
                "Item SKU": " abc ",
                "Product Name": "Widget",
                "Length": "10",
                "Width": "5",
                "Depth": "2",
                "Weight lb": "1",
                "Warehouse Note": "fragile",
            }
        ],
    )

    items = read_sku_master(str(path))

    assert len(items) == 1
    assert items[0].canonical_sku == "ABC"
    assert items[0].product_name == "Widget"
    assert items[0].length_cm == 10
    assert items[0].weight_kg == 0.45359237
    assert items[0].metadata["Warehouse Note"] == "fragile"


def test_read_orders_from_csv_infers_columns_and_preserves_shipping_metadata(tmp_path):
    path = tmp_path / "orders.csv"
    _write_csv(
        path,
        [
            {
                "Backer Number": "B-1",
                "Product SKU": "abc",
                "Qty": "2",
                "Country": "US",
                "State": "CA",
                "Region": "West",
                "Shipping Service": "Ground",
            }
        ],
    )

    orders = read_orders(str(path))

    assert len(orders) == 1
    assert orders[0].order_id == "B-1"
    assert orders[0].canonical_sku == "ABC"
    assert orders[0].quantity == 2
    assert orders[0].metadata["Shipping Service"] == "Ground"


def test_read_intake_matches_known_skus_and_creates_unmatched_records(tmp_path):
    sku_path = tmp_path / "sku_master.csv"
    order_path = tmp_path / "orders.csv"
    _write_csv(
        sku_path,
        [
            {
                "SKU": "ABC",
                "Product Name": "Widget",
                "Length": "10",
                "Width": "5",
                "Height": "2",
                "Weight kg": "1",
            }
        ],
    )
    _write_csv(
        order_path,
        [
            {"Order ID": "1", "SKU": "ABC", "Quantity": "1", "Gift Note": "keep"},
            {"Order ID": "2", "SKU": "MISSING", "Quantity": "3", "Gift Note": "preserve"},
        ],
    )

    result = read_intake(str(sku_path), str(order_path))

    assert [line.canonical_sku for line in result.matched_order_lines] == ["ABC"]
    assert len(result.unmatched_skus) == 1
    assert result.unmatched_skus[0].order_line.canonical_sku == "MISSING"
    assert result.unmatched_skus[0].metadata["Gift Note"] == "preserve"


def test_read_xlsx_processes_all_useful_sheets(tmp_path):
    path = tmp_path / "sku_master.xlsx"
    _write_xlsx(
        path,
        {
            "Products A": [
                ["SKU", "Product Name", "Length", "Width", "Height", "Weight g"],
                ["A1", "Alpha", "10", "5", "2", "500"],
            ],
            "Notes": [["Comment"], ["ignore me"]],
            "Products B": [
                ["Item SKU", "Product Name", "Length", "Width", "Height", "Weight kg"],
                ["B1", "Beta", "8", "4", "2", "1"],
            ],
        },
    )

    items = read_sku_master(str(path))

    assert [item.canonical_sku for item in items] == ["A1", "B1"]
    assert {item.metadata["_source_sheet"] for item in items} == {"Products A", "Products B"}


def test_sku_master_skips_translated_subheader_and_preserves_first_duplicate_header_values(tmp_path):
    path = tmp_path / "vestige_style_sku.xlsx"
    _write_xlsx(
        path,
        {
            "stock": [
                ["SKU", "Item name", "", "", "Weight/g", "L-cm", "W-cm", "H-cm", "SKU", "Item name"],
                ["SKU编号", "商品全名", "成本算法", "条码", "重量(kg)", "长", "宽", "高", "", ""],
                ["ONB0100", "Vindication Base Game", "", "", "3640", "39", "32", "11", "", ""],
            ],
        },
    )

    items = read_sku_master(str(path))

    assert len(items) == 1
    assert items[0].raw_sku == "ONB0100"
    assert items[0].product_name == "Vindication Base Game"
    assert items[0].length_cm == 39
    assert items[0].width_cm == 32
    assert items[0].height_cm == 11
    assert items[0].weight_kg == 3.64


def test_read_workbook_extracts_adjacent_factory_metadata(tmp_path):
    path = tmp_path / "factory_metadata.xlsx"
    _write_xlsx(
        path,
        {
            "stock": [
                ["SKU", "Item name", "Weight/g", "L-cm", "W-cm", "H-cm", "", ""],
                ["ONB0100", "Vindication Base Game", "3640", "39", "32", "11", "Factory", "WHATZ"],
            ],
        },
    )

    source = read_workbook(str(path))[0]

    assert source.metadata["factory_name"] == "WHATZ"


def test_read_workbook_preserves_blank_header_columns_for_intake_copy(tmp_path):
    path = tmp_path / "factory_metadata.xlsx"
    _write_xlsx(
        path,
        {
            "stock": [
                ["SKU", "Item name", "Weight/g", "L-cm", "W-cm", "H-cm", "", ""],
                ["SKU编号", "商品全名", "重量(kg)", "长", "宽", "高", "", ""],
                ["ONB0100", "Vindication Base Game", "3640", "39", "32", "11", "", ""],
                ["", "", "", "", "", "", "Factory", "WHATZ"],
            ],
        },
    )

    source = read_workbook(str(path))[0]

    assert "Column G" in source.preserved_rows[0]
    assert source.preserved_rows[0]["Column G"] == ""
    assert source.preserved_rows[2]["Column G"] == "Factory"
    assert source.preserved_rows[2]["Column H"] == "WHATZ"


def test_read_sku_master_supports_vfi_short_cm_dimension_headers(tmp_path):
    path = tmp_path / "vfi_sku_master.xlsx"
    _write_xlsx(
        path,
        {
            "VFI Intake form": [
                ["SKU", "Item name", "SKU/ UPC", "Weight/g", "L-cm", "W-cm", "H-cm"],
                ["SKU description", "Product description", "Barcode", "Weight", "Length", "Width", "Height"],
                ["OPR_GEN_A_002", "One Page Rules Tokens Set (70)", "", "121", "28.5", "14.5", "1"],
            ]
        },
    )

    items = read_sku_master(str(path))

    assert len(items) == 1
    assert items[0].canonical_sku == "OPR_GEN_A_002"
    assert items[0].product_name == "One Page Rules Tokens Set (70)"
    assert (items[0].length_cm, items[0].width_cm, items[0].height_cm) == (28.5, 14.5, 1)
    assert items[0].weight_kg == 0.121

def test_combined_dimension_parsing_supports_three_dimension_formats(tmp_path):
    path = tmp_path / "sku_master.csv"
    _write_csv(
        path,
        [
            {"SKU": "A", "Dimensions": "10 x 5 x 3 cm", "Weight kg": "1"},
            {"SKU": "B", "Dimensions": "10×5×3", "Weight kg": "1"},
        ],
    )

    items = read_sku_master(str(path))

    assert [(item.length_cm, item.width_cm, item.height_cm, item.is_flat) for item in items] == [
        (10, 5, 3, False),
        (10, 5, 3, False),
    ]


def test_combined_dimension_header_unit_applies_to_bare_values(tmp_path):
    path = tmp_path / "sku_master.csv"
    _write_csv(
        path,
        [{"SKU": "MM", "Dimensions (mm)": "294 x 294 x 78", "Weight kg": "1"}],
    )

    item = read_sku_master(str(path))[0]

    assert round(item.length_cm, 1) == 29.4
    assert round(item.width_cm, 1) == 29.4
    assert round(item.height_cm, 1) == 7.8


def test_sku_weight_parses_scientific_notation(tmp_path):
    path = tmp_path / "sku_master.csv"
    _write_csv(
        path,
        [{"SKU": "DICE", "Dimensions (mm)": "10 x 10 x 10", "Weight (KG)": "4.4999999999999998E-2"}],
    )

    item = read_sku_master(str(path))[0]

    assert item.weight_kg == 0.045


def test_combined_dimension_parsing_treats_two_dimensions_as_flat(tmp_path):
    path = tmp_path / "sku_master.csv"
    _write_csv(
        path,
        [{"SKU": "Flat", "Dimensions": "10 x 5", "Weight kg": "1"}],
    )

    item = read_sku_master(str(path))[0]

    assert (item.length_cm, item.width_cm, item.height_cm) == (10, 5, 1)
    assert item.is_flat is True


def test_wide_format_order_parsing_creates_order_lines(tmp_path):
    path = tmp_path / "orders.csv"
    _write_csv(
        path,
        [
            {
                "Order ID": "1001",
                "Country": "US",
                "Core Game": "1",
                "Expansion A": "2",
                "Add-on B": "",
            }
        ],
    )

    lines = read_orders(str(path))

    assert [(line.raw_sku, line.quantity) for line in lines] == [
        ("Core Game", 1),
        ("Expansion A", 2),
    ]
    assert all(line.order_id == "1001" for line in lines)
    assert all(line.country == "US" for line in lines)


def test_wide_format_vfi_backer_metadata_columns_are_not_product_columns(tmp_path):
    path = tmp_path / "backers.xlsx"
    _write_xlsx(
        path,
        {
            "Backer Data": [
                ["VFI #", "Shipping name", "phone", "email", "add 1", "add 2", "Shipping City", "Shipping State", "Shipping Postal Code", "Ship to Country Code", "SKU-A", "SKU-B"],
                ["1", "Garrett Hunter", "2507556593", "person@example.com", "1124 2 Avenue", "", "Ladysmith", "British Columbia", "V9G1J5", "CA", "1", "0"],
                ["2", "Jordan Horn", "6048192565", "j@example.com", "45640 alma avenue", "408", "Chilliwack", "British Columbia", "V2R0P8", "CA", "0", "2"],
            ]
        },
    )

    lines = read_orders(str(path))

    assert [(line.order_id, line.raw_sku, line.quantity) for line in lines] == [
        ("1", "SKU-A", 1),
        ("2", "SKU-B", 2),
    ]
    assert all(line.raw_sku not in {"VFI #", "add 2"} for line in lines)
    assert lines[0].country == "CA"
    assert lines[0].metadata["Shipping name"] == "Garrett Hunter"
    assert lines[1].metadata["add 2"] == "408"


def test_wide_format_metadata_columns_are_not_product_columns(tmp_path):
    path = tmp_path / "orders.csv"
    _write_csv(
        path,
        [
            {
                "Order ID": "1001",
                "Address": "1 Main St",
                "Email": "person@example.com",
                "Country": "US",
                "State": "CA",
                "Postal": "90001",
                "Notes": "Leave at door",
                "Core Game": "1",
            }
        ],
    )

    lines = read_orders(str(path))

    assert len(lines) == 1
    assert lines[0].raw_sku == "Core Game"
    assert lines[0].metadata["Address"] == "1 Main St"
    assert lines[0].metadata["Email"] == "person@example.com"


def test_xlsx_reader_uses_second_row_as_headers_when_first_row_is_title(tmp_path):
    path = tmp_path / "orders.xlsx"
    _write_xlsx(
        path,
        {
            "Orders": [
                ["Campaign export", "", "", ""],
                ["Backer Number", "Name", "Email", "Core Game"],
                ["B-100", "Grace Hopper", "grace@example.com", "1"],
            ]
        },
    )

    lines = read_orders(str(path))

    assert len(lines) == 1
    assert lines[0].raw_sku == "Core Game"
    assert lines[0].metadata["Backer Number"] == "B-100"
    assert lines[0].metadata["Name"] == "Grace Hopper"
    assert lines[0].metadata["Email"] == "grace@example.com"


def test_xlsx_reader_merges_row_one_product_headers_with_row_two_metadata_headers(tmp_path):
    path = tmp_path / "orders.xlsx"
    _write_xlsx(
        path,
        {
            "Orders": [
                ["", "", "Core Game"],
                ["Backer Number", "Email", ""],
                ["B-100", "grace@example.com", "1"],
            ]
        },
    )

    lines = read_orders(str(path))

    assert len(lines) == 1
    assert lines[0].raw_sku == "Core Game"
    assert lines[0].metadata["Backer Number"] == "B-100"
    assert lines[0].metadata["Email"] == "grace@example.com"


def test_wide_format_skips_total_row_without_backer_identity(tmp_path):
    path = tmp_path / "backers.xlsx"
    _write_xlsx(
        path,
        {
            "Backer Data": [
                ["VFI #", "Backer ID", "Shipping name", "email", "Country", "SKU-A", "SKU-B"],
                ["1", "B-1", "Ada", "ada@example.com", "US", "1", "0"],
                ["", "", "", "", "", "12", "8"],
            ]
        },
    )

    lines = read_orders(str(path))

    assert [(line.order_id, line.raw_sku, line.quantity) for line in lines] == [("1", "SKU-A", 1)]




def test_sku_master_expands_merged_dimension_header_and_mm_units(tmp_path):
    sku_path = tmp_path / "sku_master.xlsx"
    _write_xlsx(
        sku_path,
        {
            "SKU INFORMATION": [
                ["SKU", "Total Units Arriving to Warehouse", "Individual Unit Dimensions (mm)", "", "", "Unit Weight (g)", "Unit Weight (kg)"],
                ["MOONCOLBOX", "38", "408", "280", "127", "6700", "6.7"],
            ]
        },
    )

    items = read_sku_master(str(sku_path))

    assert len(items) == 1
    assert items[0].canonical_sku == "MOONCOLBOX"
    assert round(items[0].length_cm, 1) == 40.8
    assert round(items[0].width_cm, 1) == 28.0
    assert round(items[0].height_cm, 1) == 12.7
    assert items[0].weight_kg == 6.7


def test_wide_backer_list_uses_sku_columns_after_merged_dimension_sku_parse(tmp_path):
    sku_path = tmp_path / "sku_master.xlsx"
    orders_path = tmp_path / "orders.xlsx"
    _write_xlsx(
        sku_path,
        {
            "SKU INFORMATION": [
                ["SKU", "Total Units Arriving to Warehouse", "Individual Unit Dimensions (mm)", "", "", "Unit Weight (g)", "Unit Weight (kg)"],
                ["MOONCOLBOX", "38", "408", "280", "127", "6700", "6.7"],
            ]
        },
    )
    _write_xlsx(
        orders_path,
        {
            "PHYSICAL FULFILLMENT": [
                ["Id", "Name", "Email", "Address State", "Address Country", "Order ID", "MOONCOLBOX"],
                ["37043809", "Harry", "harry@example.com", "VIC", "AU", "25150171", "1"],
            ]
        },
    )

    result = read_intake(str(sku_path), str(orders_path))

    assert len(result.sku_items) == 1
    assert len(result.matched_order_lines) == 1
    line = result.matched_order_lines[0]
    assert line.order_id == "25150171"
    assert line.canonical_sku == "MOONCOLBOX"
    assert line.quantity == 1
    assert line.country == "AU"
    assert line.state_province == "VIC"
    assert line.metadata["Name"] == "Harry"
    assert result.unmatched_skus == []

def test_intake_ignores_non_order_sheets_and_only_uses_sku_master_product_columns(tmp_path):
    sku_path = tmp_path / "sku_master.csv"
    orders_path = tmp_path / "orders.xlsx"
    _write_csv(
        sku_path,
        [
            {"SKU": "SKU-A", "Product Name": "Core Game", "Dimensions": "10 x 5 x 3", "Weight kg": "1"},
            {"SKU": "SKU-B", "Product Name": "Expansion", "Dimensions": "8 x 4 x 2", "Weight kg": "0.5"},
        ],
    )
    _write_xlsx(
        orders_path,
        {
            "Backer Data": [
                ["VFI #", "Backer ID", "Shipping name", "email", "Country", "SKU-A", "SKU-B"],
                ["1", "B-1", "Ada", "ada@example.com", "US", "1", "2"],
                ["", "", "", "", "", "20", "30"],
            ],
            "SF Asia": [
                ["Order Id", "Name", "Phone Number", "Email", "Country", "Item Total", "Declared Name", "Unit Price", "No. of Packages"],
                ["0104BIOS1", "Ada", "123", "ada@example.com", "CN", "99", "Board Game", "10", "1"],
                ["0104BIOS", "0", "0", "0", "0", "#REF!", "Board Game", "#REF!", "3"],
            ],
        },
    )

    result = read_intake(str(sku_path), str(orders_path))

    assert [(line.order_id, line.raw_sku, line.quantity) for line in result.order_lines] == [
        ("1", "SKU-A", 1),
        ("1", "SKU-B", 2),
    ]
    assert result.unmatched_skus == []
    assert result.debug["detected_product_quantity_columns"] == ["SKU-A", "SKU-B"]

def test_wide_format_product_header_matches_sku_master_product_name(tmp_path):
    sku_path = tmp_path / "sku_master.csv"
    order_path = tmp_path / "orders.csv"
    _write_csv(
        sku_path,
        [{"SKU": "CG-001", "Product Name": "Core Game", "Dimensions": "10 x 5 x 3", "Weight kg": "1"}],
    )
    _write_csv(
        order_path,
        [{"Order ID": "1001", "Core Game": "1"}],
    )

    result = read_intake(str(sku_path), str(order_path))

    assert len(result.matched_order_lines) == 1
    assert result.matched_order_lines[0].canonical_sku == "CG-001"
    assert result.unmatched_skus == []


def test_real_example_files_parse_nonzero_skus_and_order_lines():
    result = read_intake("examples/sku_master.xlsx", "examples/orders.xlsx")

    assert result.debug["sku_items_parsed"] > 0
    assert result.debug["order_lines_created"] > 0
    assert result.debug["wide_product_columns_detected"] > 0
    assert result.debug["matched"] > 0

