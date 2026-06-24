import zipfile
import struct
from xml.etree import ElementTree

from box_optimizer.io.excel_reader import read_workbook
from box_optimizer.io.excel_writer import (
    MAX_LABELS_PER_SHEET,
    MAX_MANUAL_ROW_BREAKS_PER_SHEET,
    _column_letter,
    workbook_sheet_stats,
    write_workbook,
)


NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
}


def test_write_workbook_returns_target_path():
    assert write_workbook("output.xlsx", rows=[]) == "output.xlsx"


def _workbook_sheet_names(path):
    with zipfile.ZipFile(path) as archive:
        root = ElementTree.fromstring(archive.read("xl/workbook.xml"))
    return [
        sheet.attrib["name"]
        for sheet in root.findall("main:sheets/main:sheet", NS)
    ]


def _worksheet_root(path, sheet_index):
    with zipfile.ZipFile(path) as archive:
        return ElementTree.fromstring(
            archive.read(f"xl/worksheets/sheet{sheet_index}.xml")
        )


def _worksheet_xml(path, sheet_name):
    sheet_names = _workbook_sheet_names(path)
    sheet_index = sheet_names.index(sheet_name) + 1
    with zipfile.ZipFile(path) as archive:
        return archive.read(f"xl/worksheets/sheet{sheet_index}.xml").decode("utf-8")


def _row_break_count(path, sheet_name):
    sheet_names = _workbook_sheet_names(path)
    sheet_index = sheet_names.index(sheet_name) + 1
    root = _worksheet_root(path, sheet_index)
    row_breaks = root.find("main:rowBreaks", NS)
    return 0 if row_breaks is None else int(row_breaks.attrib["count"])


def _label_rows(count, *, with_qr=False):
    return [
        {
            "Label Number": f"LBL-{index:04d}",
            "Barcode/QR Value": f"QR-{index:04d}" if with_qr else "",
            "From": "VFI",
            "To Name": f"Backer {index}",
            "Items to Pack Column 1": "CORE x1",
        }
        for index in range(1, count + 1)
    ]


def _styles_root(path):
    with zipfile.ZipFile(path) as archive:
        return ElementTree.fromstring(archive.read("xl/styles.xml"))


def test_write_workbook_creates_required_tabs_first_in_exact_order(tmp_path):
    path = tmp_path / "report.xlsx"

    write_workbook(
        str(path),
        summary_rows=[{"Orders": 2}],
        order_volume_weights_rows=[{"Order ID": "1", "Chargeable Weight": 2.5}],
        box_size_summary_rows=[{"Box Type": "Box Type 1", "Length": 10}],
        unmatched_skus_rows=[{"SKU": "MISSING"}],
    )

    sheet_names = _workbook_sheet_names(path)

    assert sheet_names[:8] == [
        "Summary",
        "Cost Summary",
        "Labels",
        "VFI Intake Form",
        "Optimized to Pack",
        "Label generator",
        "Order Volume Weights",
        "Box Size Summary",
    ]
    assert sheet_names[8:] == ["Unmatched SKUs"]


def test_write_workbook_fast_production_skips_helper_and_detail_tabs(tmp_path):
    path = tmp_path / "report.xlsx"

    write_workbook(
        str(path),
        summary_rows=[{"Metric": "Orders", "Value": 2}],
        cost_summary_rows=[{"Backer ID": "1", "VFI #": "VFI-1"}],
        labels_rows=[{"Label Number": "VFI-1", "Barcode/QR Value": "VFI-1"}],
        vfi_intake_form_rows=[{"Campaign Name": "Campaign"}],
        optimized_to_pack_rows=[{"Pledge Configuration": 1}],
        label_generator_rows=[{"Order ID": "1"}],
        order_volume_weights_rows=[{"Order ID": "1"}],
        packing_detail_rows=[{"Order ID": "1"}],
        multi_box_detail_rows=[{"Order ID": "1"}],
        pledge_combination_summary_rows=[{"SKU Breakdown": "CORE x1"}],
        debug_summary_rows=[{"Metric": "Debug"}],
        input_column_mapping_rows=[{"workbook": "orders.csv"}],
        errors_and_warnings_rows=[{"Message": "warning"}],
        country_scan_sheets={"US": [{"Barcode/QR Value": "VFI-1"}]},
        workbook_output_mode="fast_production",
    )

    sheet_names = _workbook_sheet_names(path)

    assert sheet_names == [
        "Summary",
        "Cost Summary",
        "Labels",
        "US",
        "VFI Intake Form",
        "Optimized to Pack",
        "Box Size Summary",
        "Errors and Warnings",
    ]
    assert "Label generator" not in sheet_names
    assert "Order Volume Weights" not in sheet_names
    assert "Packing Detail" not in sheet_names
    assert "Multi Box Detail" not in sheet_names
    assert "Pledge Combination Summary" not in sheet_names
    assert "Debug Summary" not in sheet_names
    assert "Errors and Warnings" in sheet_names

    stats = workbook_sheet_stats(
        summary_rows=[{"Metric": "Orders", "Value": 2}],
        cost_summary_rows=[{"Backer ID": "1", "VFI #": "VFI-1"}],
        labels_rows=[{"Label Number": "VFI-1", "Barcode/QR Value": "VFI-1"}],
        vfi_intake_form_rows=[{"Campaign Name": "Campaign"}],
        optimized_to_pack_rows=[{"Pledge Configuration": 1}],
        label_generator_rows=[{"Order ID": "1"}],
        order_volume_weights_rows=[{"Order ID": "1"}],
        packing_detail_rows=[{"Order ID": "1"}],
        multi_box_detail_rows=[{"Order ID": "1"}],
        pledge_combination_summary_rows=[{"SKU Breakdown": "CORE x1"}],
        debug_summary_rows=[{"Metric": "Debug"}],
        input_column_mapping_rows=[{"workbook": "orders.csv"}],
        errors_and_warnings_rows=[{"Message": "warning"}],
        country_scan_sheets={"US": [{"Barcode/QR Value": "VFI-1"}]},
        workbook_output_mode="fast_production",
    )
    assert stats["sheets_written"] == sheet_names
    assert stats["country_sheet_count"] == 1
    assert stats["qr_images_written"] == 1
    assert {"Label generator", "Order Volume Weights", "Packing Detail"} <= set(stats["sheets_skipped"])


def test_write_workbook_creates_single_labels_sheet_for_1000_or_fewer_labels(tmp_path):
    path = tmp_path / "report.xlsx"

    write_workbook(str(path), labels_rows=_label_rows(MAX_LABELS_PER_SHEET))

    sheet_names = _workbook_sheet_names(path)

    assert "Labels" in sheet_names
    assert "Labels 2" not in sheet_names
    assert _row_break_count(path, "Labels") == MAX_LABELS_PER_SHEET - 1


def test_write_workbook_splits_labels_after_1000_label_blocks(tmp_path):
    path = tmp_path / "report.xlsx"

    write_workbook(str(path), labels_rows=_label_rows(MAX_LABELS_PER_SHEET + 1))

    sheet_names = _workbook_sheet_names(path)

    assert "Labels" in sheet_names
    assert "Labels 2" in sheet_names
    assert "Labels 3" not in sheet_names
    assert _row_break_count(path, "Labels") == MAX_LABELS_PER_SHEET - 1
    assert _row_break_count(path, "Labels 2") == 0
    labels_xml = _worksheet_xml(path, "Labels")
    labels_2_xml = _worksheet_xml(path, "Labels 2")
    assert "LBL-1000" in labels_xml
    assert "LBL-1001" not in labels_xml
    assert "LBL-1001" in labels_2_xml


def test_write_workbook_splits_labels_into_third_sheet_after_2000_label_blocks(tmp_path):
    path = tmp_path / "report.xlsx"

    write_workbook(str(path), labels_rows=_label_rows((MAX_LABELS_PER_SHEET * 2) + 1))

    sheet_names = _workbook_sheet_names(path)

    assert "Labels" in sheet_names
    assert "Labels 2" in sheet_names
    assert "Labels 3" in sheet_names
    assert "Labels 4" not in sheet_names
    for sheet_name in ["Labels", "Labels 2", "Labels 3"]:
        assert _row_break_count(path, sheet_name) <= MAX_MANUAL_ROW_BREAKS_PER_SHEET
    assert "LBL-2000" in _worksheet_xml(path, "Labels 2")
    assert "LBL-2001" not in _worksheet_xml(path, "Labels 2")
    assert "LBL-2001" in _worksheet_xml(path, "Labels 3")


def test_write_workbook_fast_production_keeps_split_printable_labels(tmp_path):
    path = tmp_path / "report.xlsx"

    write_workbook(
        str(path),
        labels_rows=_label_rows(MAX_LABELS_PER_SHEET + 1),
        label_generator_rows=[{"Order ID": "1"}],
        country_scan_sheets={"US": [{"Barcode/QR Value": "VFI-1"}]},
        workbook_output_mode="fast_production",
    )

    sheet_names = _workbook_sheet_names(path)

    assert "Labels" in sheet_names
    assert "Labels 2" in sheet_names
    assert sheet_names.index("US") > sheet_names.index("Labels 2")
    assert "Label generator" not in sheet_names


def test_split_labels_sheets_have_drawing_relationships_for_qr_images(tmp_path, monkeypatch):
    path = tmp_path / "report.xlsx"
    monkeypatch.setattr(
        "box_optimizer.io.excel_writer.qr_png",
        lambda _value, scale=8, border=4: b"\x89PNG\r\n\x1a\nFAKE",
    )

    write_workbook(str(path), labels_rows=_label_rows(MAX_LABELS_PER_SHEET + 1, with_qr=True))

    sheet_names = _workbook_sheet_names(path)
    labels_index = sheet_names.index("Labels") + 1
    labels_2_index = sheet_names.index("Labels 2") + 1
    with zipfile.ZipFile(path) as archive:
        names = set(archive.namelist())
        labels_rels = archive.read(f"xl/worksheets/_rels/sheet{labels_index}.xml.rels").decode("utf-8")
        labels_2_rels = archive.read(f"xl/worksheets/_rels/sheet{labels_2_index}.xml.rels").decode("utf-8")
        drawing_1_rels = archive.read("xl/drawings/_rels/drawing1.xml.rels").decode("utf-8")
        drawing_2_rels = archive.read("xl/drawings/_rels/drawing2.xml.rels").decode("utf-8")

    assert "xl/drawings/drawing1.xml" in names
    assert "xl/drawings/drawing2.xml" in names
    assert 'Target="../drawings/drawing1.xml"' in labels_rels
    assert 'Target="../drawings/drawing2.xml"' in labels_2_rels
    assert 'Target="../media/label_qr_1.png"' in drawing_1_rels
    assert 'Target="../media/label_qr_1000.png"' in drawing_1_rels
    assert 'Target="../media/label_qr_1001.png"' in drawing_2_rels
    assert "xl/media/label_qr_1001.png" in names


def test_split_labels_display_country_code_and_preserve_chinese_on_later_sheets(tmp_path):
    path = tmp_path / "report.xlsx"
    labels = _label_rows(MAX_LABELS_PER_SHEET + 1)
    labels[-1].update(
        {
            "Country": "Germany",
            "Country Code": "DE",
            "Country Package Code": "DE  2-1",
            "Country Name Chinese": "\u5fb7\u56fd",
        }
    )

    write_workbook(str(path), labels_rows=labels)

    labels_2_xml = _worksheet_xml(path, "Labels 2")

    assert 'r="E1" t="inlineStr" s="25"><is><t>DE</t></is></c>' in labels_2_xml
    assert "\u5fb7\u56fd" in labels_2_xml
    assert "DE  2-1" not in labels_2_xml
    assert "UN  2-1" not in labels_2_xml


def test_fast_production_split_labels_display_country_code_and_preserve_chinese(tmp_path):
    path = tmp_path / "report.xlsx"
    labels = _label_rows(MAX_LABELS_PER_SHEET + 1)
    labels[-1].update(
        {
            "Country": "France",
            "Country Code": "FR",
            "Country Package Code": "FR  1",
            "Country Name Chinese": "\u6cd5\u56fd",
        }
    )

    write_workbook(
        str(path),
        labels_rows=labels,
        label_generator_rows=[{"Order ID": "1"}],
        workbook_output_mode="fast_production",
    )

    sheet_names = _workbook_sheet_names(path)
    labels_2_xml = _worksheet_xml(path, "Labels 2")

    assert "Labels 2" in sheet_names
    assert "Label generator" not in sheet_names
    assert 'r="E1" t="inlineStr" s="25"><is><t>FR</t></is></c>' in labels_2_xml
    assert "\u6cd5\u56fd" in labels_2_xml
    assert "FR  1" not in labels_2_xml
    assert "UN  1" not in labels_2_xml


def test_split_labels_preserve_friendly_carton_box_designation_on_later_sheets(tmp_path):
    path = tmp_path / "report.xlsx"
    labels = _label_rows(MAX_LABELS_PER_SHEET + 1)
    labels[-1].update(
        {
            "Carton Box Designation": "All In",
            "Items to Pack Column 1": "CORE x1",
        }
    )

    write_workbook(str(path), labels_rows=labels)

    labels_2_xml = _worksheet_xml(path, "Labels 2")

    assert "All In" in labels_2_xml
    assert "Earth Under Siege All In Storage shipping carton" not in labels_2_xml


def test_write_workbook_freezes_headers_applies_filters_and_widths(tmp_path):
    path = tmp_path / "report.xlsx"

    write_workbook(str(path), rows=[{"Order ID": "1", "Length": 10, "Weight": 2}])

    root = _worksheet_root(path, 1)

    assert root.find(".//main:pane", NS).attrib["state"] == "frozen"
    assert root.find(".//main:autoFilter", NS).attrib["ref"] == "A1:C2"
    assert len(root.findall(".//main:col", NS)) == 3


def test_write_workbook_includes_units_in_headers(tmp_path):
    path = tmp_path / "report.xlsx"

    write_workbook(str(path), rows=[{"Length": 10, "Width": 8, "Volume": 800}])

    rows = read_workbook(str(path))[0].rows

    assert rows == [{"Length (cm)": "10", "Width (cm)": "8", "Volume (cm3)": "800"}]




def test_write_workbook_rounds_measure_outputs_to_one_decimal_without_touching_money(tmp_path):
    path = tmp_path / "report.xlsx"

    write_workbook(
        str(path),
        vfi_intake_form_rows=[
            {
                "Length Individual Unit Dimensions (mm)": 40.800000000000004,
                "Unit Weight (kg)": 0.16250000000000001,
                "Customer Shipping Fee": 120.83,
                "Backer ID": "37043626",
            }
        ],
        packing_detail_rows=[{"Placement X cm": 31.400000000000002}],
    )

    intake_rows = next(sheet.rows for sheet in read_workbook(str(path)) if sheet.sheet_name == "VFI Intake Form")
    packing_rows = next(sheet.rows for sheet in read_workbook(str(path)) if sheet.sheet_name == "Packing Detail")

    assert intake_rows[0]["Length Individual Unit Dimensions (mm)"] == "40.8"
    assert intake_rows[0]["Unit Weight (kg)"] == "0.2"
    assert intake_rows[0]["Customer Shipping Fee"] == "120.83"
    assert intake_rows[0]["Backer ID"] == "37043626"
    assert packing_rows[0]["Placement X cm"] == "31.4"


def test_write_workbook_formats_shipping_money_columns_as_usd(tmp_path):
    path = tmp_path / "report.xlsx"

    write_workbook(
        str(path),
        summary_rows=[
            {
                "Section": "Cost Placeholder",
                "Metric": "Total Chargeable Cost",
                "Value": 460.5,
                "Detail": "Hub Shipping Fee + Express",
            }
        ],
        sheets={
            "Cost Summary - Test": [
                {
                    "Shipping Method": "ship to Partner: partner@example.com",
                    "Hub Shipping Fee": 24,
                    "Express": 0,
                }
            ]
        },
    )

    styles = _styles_root(path)
    summary = _worksheet_root(path, 1)
    cost_summary = _worksheet_root(path, 2)

    number_formats = styles.findall("main:numFmts/main:numFmt", NS)
    assert any(
        fmt.attrib["numFmtId"] == "164"
        and fmt.attrib["formatCode"] == '$#,##0.00" (USD)"'
        for fmt in number_formats
    )
    assert any(
        fmt.attrib["numFmtId"] == "165"
        and fmt.attrib["formatCode"] == "$#,##0.00"
        for fmt in number_formats
    )
    cell_formats = styles.findall("main:cellXfs/main:xf", NS)
    assert cell_formats[9].attrib["numFmtId"] == "164"
    assert cell_formats[9].attrib["applyNumberFormat"] == "1"
    assert cell_formats[19].attrib["numFmtId"] == "165"
    assert cell_formats[19].attrib["applyNumberFormat"] == "1"
    assert summary.find(".//main:c[@r='C2']", NS).attrib["s"] == "9"
    assert "Hub Shipping Fee (USD)" in ElementTree.tostring(cost_summary, encoding="unicode")
    assert "Express (USD)" in ElementTree.tostring(cost_summary, encoding="unicode")
    assert cost_summary.find(".//main:c[@r='B2']", NS).attrib["s"] == "19"
    assert cost_summary.find(".//main:c[@r='C2']", NS).attrib["s"] == "19"
    assert cost_summary.find(".//main:c[@r='B2']/main:v", NS).text == "24"
    assert cost_summary.find(".//main:c[@r='C2']/main:v", NS).text == "0"


def test_vfi_intake_form_hides_database_upload_columns_without_deleting_data(tmp_path):
    path = tmp_path / "report.xlsx"
    row = {f"Column {_column_letter(index)}": f"value-{index + 1}" for index in range(28)}

    write_workbook(str(path), vfi_intake_form_rows=[row])

    sheet_names = _workbook_sheet_names(path)
    intake_index = sheet_names.index("VFI Intake Form") + 1
    root = _worksheet_root(path, intake_index)
    columns = root.findall(".//main:cols/main:col", NS)
    by_min = {int(column.attrib["min"]): column for column in columns}

    for column_number in range(12, 28):
        assert by_min[column_number].attrib["hidden"] == "1"
        assert by_min[column_number].attrib["collapsed"] == "1"
    assert "hidden" not in by_min[11].attrib
    assert "hidden" not in by_min[28].attrib
    xml = ElementTree.tostring(root, encoding="unicode")
    assert "value-12" in xml
    assert "value-27" in xml


def test_summary_sku_intake_shortage_remaining_is_red(tmp_path):
    path = tmp_path / "report.xlsx"

    write_workbook(
        str(path),
        summary_rows=[
            {
                "Section": "",
                "Metric": "",
                "Value": "",
                "Detail": "",
                "SKU": "ORDERONLY",
                "Received Quantity": 0,
                "Required Quantity": 4,
                "Remaining": -4,
            }
        ],
    )

    root = _worksheet_root(path, 1)
    styles = _styles_root(path)

    assert root.find(".//main:c[@r='H2']", NS).attrib["s"] == "20"
    cell_formats = styles.findall("main:cellXfs/main:xf", NS)
    assert cell_formats[20].attrib["fontId"] == "8"
    fonts = styles.findall("main:fonts/main:font", NS)
    assert fonts[8].find("main:color", NS).attrib["rgb"] == "FFFF0000"


def test_labels_sheet_follows_campaign_suffixed_cost_summary(tmp_path):
    path = tmp_path / "report.xlsx"

    write_workbook(
        str(path),
        sheets={
            "Cost Summary - Sordane": [{"Order ID": "1", "Hub Shipping Fee": 12}],
        },
        labels_rows=[{"Label Number": "39", "Barcode/QR Value": "OPR 39"}],
    )

    sheet_names = _workbook_sheet_names(path)

    assert sheet_names[1:3] == ["Cost Summary - Sordane", "Labels"]
    assert sheet_names.count("Labels") == 1
    assert sheet_names.count("Cost Summary - Sordane") == 1


def test_country_scan_tabs_are_inserted_after_labels(tmp_path):
    path = tmp_path / "report.xlsx"

    write_workbook(
        str(path),
        sheets={"Cost Summary - Sordane": [{"Order ID": "1", "Hub Shipping Fee": 12}]},
        labels_rows=[{"Label Number": "39", "Barcode/QR Value": "OPR 39"}],
        country_scan_sheets={
            "Hong Kong": [{"Country": "Hong Kong", "VFI #": "OPR 39", "": ""}],
            "Singapore": [{"Country": "Singapore", "VFI #": "OPR 40", "": ""}],
        },
    )

    sheet_names = _workbook_sheet_names(path)

    assert sheet_names[1:5] == ["Cost Summary - Sordane", "Labels", "Hong Kong", "Singapore"]
    hong_kong_rows = next(sheet.rows for sheet in read_workbook(str(path)) if sheet.sheet_name == "Hong Kong")
    assert hong_kong_rows == [{"Country": "Hong Kong", "VFI #": "OPR 39"}]


def test_write_workbook_creates_optional_detail_tabs_when_rows_exist(tmp_path):
    path = tmp_path / "report.xlsx"

    write_workbook(
        str(path),
        packing_detail_rows=[{"Order ID": "1", "X": 0, "Y": 0, "Z": 0}],
        multi_box_detail_rows=[{"Order ID": "1", "Box Qty": 2}],
        input_column_mapping_rows=[{"Input Column": "SKU", "Mapped Field": "sku"}],
        errors_and_warnings_rows=[{"Severity": "Warning", "Message": "Check SKU"}],
    )

    assert _workbook_sheet_names(path) == [
        "Summary",
        "Cost Summary",
        "Labels",
        "VFI Intake Form",
        "Optimized to Pack",
        "Label generator",
        "Order Volume Weights",
        "Box Size Summary",
        "Multi Box Detail",
        "Packing Detail",
        "Input Column Mapping",
        "Errors and Warnings",
    ]


def test_order_volume_weights_leads_with_required_columns_then_metadata(tmp_path):
    path = tmp_path / "report.xlsx"

    write_workbook(
        str(path),
        order_volume_weights_rows=[
            {
                "Order ID": "A-1",
                "Region": "NA",
                "Chargeable Weight kg": 2.5,
                "SKU Breakdown": "Core Game x1",
                "Pledge Level": "Deluxe",
                "Shipping Notes": "Leave at door",
            }
        ],
    )

    rows = next(sheet.rows for sheet in read_workbook(str(path)) if sheet.sheet_name == "Order Volume Weights")
    headers = list(rows[0].keys())

    assert headers[:18] == [
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
    assert headers[18:] == ["Pledge Level", "Shipping Notes"]
    assert rows[0]["Pledge Level"] == "Deluxe"
    assert rows[0]["Shipping Notes"] == "Leave at door"


def test_empty_order_volume_weights_still_has_required_headers(tmp_path):
    path = tmp_path / "report.xlsx"

    write_workbook(str(path))

    sheet_names = _workbook_sheet_names(path)
    order_sheet_index = sheet_names.index("Order Volume Weights") + 1
    with zipfile.ZipFile(path) as archive:
        xml = archive.read(f"xl/worksheets/sheet{order_sheet_index}.xml").decode("utf-8")

    assert "Region" in xml
    assert "SKU Breakdown" in xml


def test_labels_sheet_uses_printable_carrier_block_layout(tmp_path):
    path = tmp_path / "report.xlsx"

    write_workbook(
        str(path),
        labels_rows=[
            {
                "Label Number": "39",
                "Barcode/QR Value": "OPR 39",
                "Factory Name": "Longhai Printworks",
                "Campaign Name": "Project Longbox",
                "Country Code": "NZ",
                "Country Name Chinese": "新西兰",
                "Origin": "CN",
                "From": "No.23 Baosheng Rd.,Bld2, 3rd Floor, Longhai, Fujian, China 363107",
                "To Name": "Ada Lovelace",
                "Backer ID": "B-1",
                "Order ID": "1",
                "Pledge Configuration": 7,
                "Address Line 1": "64 BATKIN ROAD",
                "Notes": "Leave package at front desk",
                "Phone": "9195551234",
                "City": "Kfar Saba",
                "State/Province": "Choose a state",
                "Postal/Zip": "4445811",
                "Country": "New Zealand",
                "Carton Box Designation": "VB 25",
                "Total Units": 4,
                "Items to Pack Column 1": "CORE x2\nADDON x1",
                "Items to Pack Column 2": "SLEEVE x1",
                "On Arrival Note": "If product damage is found, please report to Support@VFI.asia in 24 hours",
            },
            {
                "Label Number": "40-2",
                "Barcode/QR Value": "OPR 40-2",
                "Country Code": "US",
                "Country Name Chinese": "美国",
                "From": "No.23 Baosheng Rd.,Bld2, 3rd Floor, Longhai, Fujian, China 363107",
                "To Name": "Grace Hopper",
                "Carton Box Designation": "VB 36",
                "Items to Pack Column 1": "CORE x1",
            },
        ],
    )

    sheet_names = _workbook_sheet_names(path)
    labels_index = sheet_names.index("Labels") + 1
    with zipfile.ZipFile(path) as archive:
        xml = archive.read(f"xl/worksheets/sheet{labels_index}.xml").decode("utf-8")
        root = ElementTree.fromstring(xml)
        names = set(archive.namelist())
        styles_xml = archive.read("xl/styles.xml").decode("utf-8")

    assert "VFI #" not in xml
    assert "Longhai Printworks" in xml
    assert "Campaign</t>" not in xml
    assert 'r="A1" t="inlineStr" s="24"><is><t>39 OPR</t></is></c>' in xml
    assert 'r="B1" t="inlineStr" s="24"><is><t></t></is></c>' in xml
    assert 'r="C1" t="inlineStr" s="24"><is><t></t></is></c>' in xml
    assert 'r="D1" t="inlineStr" s="24"><is><t></t></is></c>' in xml
    assert 'r="E1" t="inlineStr" s="25"><is><t>NZ</t></is></c>' in xml
    assert 'r="F1" t="inlineStr" s="25"><is><t></t></is></c>' in xml
    assert "Barcode / QR Value" not in xml
    assert "Country CN" not in xml
    assert "Total Value USD" not in xml
    assert "39 OPR" in xml
    assert 'r="B4" t="inlineStr" s="21"><is><t>No.23 Baosheng Rd.,Bld2, 3rd Floor,</t></is></c>' in xml
    assert 'r="B5" t="inlineStr" s="21"><is><t>Longhai, Fujian, China 363107</t></is></c>' in xml
    assert "No.23 Baosheng Rd.,Bld2, 3rd Floor, Longhai" not in xml
    assert "64 BATKIN ROAD" in xml
    assert "Ada Lovelace  Backer ID B-1" in xml
    assert "Notes: Leave package at front desk" in xml
    assert "CORE" in xml
    assert "SLEEVE" in xml
    assert "Detailed description of contents: Board Games-of paper and plastic,non-electrical" in xml
    assert "City/State/Zip" in xml
    assert "City/State/Post" not in xml
    assert "Kfar Saba, 4445811" in xml
    assert "Choose a state" not in xml
    assert "Country" in xml
    assert "Country Code" not in xml
    assert "New Zealand" in xml
    assert "Project Longbox     Config: 7" in xml
    assert "Pledge Config" not in xml
    assert "phone: 9195551234" in xml
    assert "On Arrival: If shipping box is damaged please take photos." in xml
    assert "If product damage is found, please report to Support@VFI.asia in 24 hours" in xml
    assert "Project Longbox" in xml
    assert "If shipping box is damaged please take photos. If damage is found" not in xml
    assert "Origin:" in xml
    assert "Total Units:" not in xml
    assert "Carton Box" not in xml
    assert "Item Count: 4" in xml
    assert "Total Item Count" not in xml
    assert "Backer ID:</t>" not in xml
    assert "Order ID" not in xml
    assert "Email" not in xml
    assert "Weight" not in xml
    assert "Overflow" not in xml
    assert 'r="F3"' not in xml
    assert 'r="B3"' not in xml
    assert 'r="C7"' not in xml
    assert 'r="D7"' not in xml
    assert 'r="F10"' not in xml
    assert 'r="A1" t="inlineStr" s="24"' in xml
    assert 'r="A12" t="inlineStr" s="10"' in xml
    assert 'r="A13" t="inlineStr" s="11"' in xml
    assert 'r="B13" t="inlineStr" s="12"' in xml
    assert 'r="A14" t="inlineStr" s="13"' in xml
    assert 'r="B14" t="inlineStr" s="14"' in xml
    assert 'r="A16" t="inlineStr" s="2"' in xml
    assert 'r="A17" t="inlineStr" s="2"' in xml
    assert 'r="A19" t="inlineStr" s="7"><is><t>Project Longbox     Config: 7</t></is></c>' in xml
    assert 'r="A20" t="inlineStr" s="7"><is><t>Item Count: 4</t></is></c>' in xml
    assert 'r="E20" t="inlineStr" s="17"><is><t>VB 25</t></is></c>' in xml
    assert 'r="F20" t="inlineStr" s="17"><is><t></t></is></c>' in xml
    assert 'r="A21" t="inlineStr" s="7"><is><t>Longhai Printworks</t></is></c>' in xml
    assert 'r="E21" t="inlineStr" s="17"' in xml
    assert 'r="F21" t="inlineStr" s="17"><is><t></t></is></c>' in xml
    assert 'r="B16"' not in xml
    assert '<font><sz val="18"/><name val="Calibri"/></font>' in styles_xml
    assert '<font><b/><sz val="31"/><name val="Calibri"/></font>' in styles_xml
    assert '<font><sz val="15"/><name val="Calibri"/></font>' in styles_xml
    assert '<xf numFmtId="0" fontId="6" fillId="0" borderId="0" xfId="0" applyFont="1"/>' in styles_xml
    assert '<alignment vertical="top" wrapText="1"/>' in styles_xml
    assert '<borders count="3">' in styles_xml
    assert '<top style="thick"><color auto="1"/></top><bottom style="thick"><color auto="1"/></bottom>' in styles_xml
    assert '<left style="thick"><color auto="1"/></left><right style="thick"><color auto="1"/></right>' in styles_xml
    assert 's="11"' in xml
    assert 's="12"' in xml
    assert "<pageSetup" in xml
    assert "<rowBreaks" in xml
    assert '<mergeCell ref="A12:F12"/>' in xml
    assert '<mergeCell ref="A1:D1"/>' in xml
    assert '<mergeCell ref="E1:F1"/>' in xml
    assert '<mergeCell ref="A1:E1"/>' not in xml
    assert '<mergeCell ref="B13:C13"/>' in xml
    assert '<mergeCell ref="E13:F13"/>' in xml
    assert '<mergeCell ref="B14:C14"/>' in xml
    assert '<mergeCell ref="E14:F14"/>' in xml
    assert '<mergeCell ref="E10:F10"/>' in xml
    assert '<mergeCell ref="E7:F9"/>' in xml
    assert 'r="E7" t="inlineStr" s="22"><is><t>Notes: Leave package at front desk</t></is></c>' in xml
    assert '<mergeCell ref="A19:F19"/>' in xml
    assert '<mergeCell ref="A20:B20"/>' in xml
    assert '<mergeCell ref="E20:F20"/>' in xml
    assert '<mergeCell ref="A21:D21"/>' in xml
    assert '<mergeCell ref="E21:F21"/>' in xml
    assert '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' in xml
    assert '<printOptions horizontalCentered="1"/>' in xml
    assert 'orientation="portrait" fitToWidth="1" fitToHeight="0"' in xml
    assert xml.index("<pageMargins") < xml.index("<pageSetup") < xml.index("<rowBreaks") < xml.index("<drawing")
    assert f"xl/worksheets/_rels/sheet{labels_index}.xml.rels" in names
    assert "xl/drawings/drawing1.xml" in names
    assert "xl/drawings/_rels/drawing1.xml.rels" in names
    assert "xl/media/label_qr_1.png" in names
    assert "xl/media/label_qr_2.png" in names
    rows = root.findall(".//main:sheetData/main:row", NS)
    cols = root.findall(".//main:cols/main:col", NS)
    assert rows[0].attrib["r"] == "1"
    assert rows[0].attrib["ht"] == "36"
    assert rows[5].attrib["ht"] == "10"
    assert rows[17].attrib["ht"] == "10"
    assert rows[13].attrib["ht"] == "22"
    assert rows[14].attrib["ht"] == "22"
    assert rows[18].attrib["ht"] == "28"
    assert rows[19].attrib["ht"] == "28"
    assert rows[20].attrib["ht"] == "28"
    assert int(rows[-1].attrib["r"]) > 20
    assert len(cols) == 6
    assert [col.attrib["width"] for col in cols] == ["12", "22", "14", "14", "22", "14"]


def test_labels_sheet_leaves_notes_area_unblocked_when_note_is_blank(tmp_path):
    path = tmp_path / "report.xlsx"

    write_workbook(
        str(path),
        labels_rows=[
            {
                "Label Number": "39",
                "Barcode/QR Value": "OPR 39",
                "Country Code": "NZ",
                "From": "No.23 Baosheng Rd.,Bld2, 3rd Floor, Longhai, Fujian, China 363107",
                "To Name": "Ada Lovelace",
                "Address Line 1": "64 BATKIN ROAD",
                "Address Line 2": "Apartment with a long delivery instruction line",
                "Notes": "",
                "Items to Pack Column 1": "CORE x1",
            },
        ],
    )

    sheet_names = _workbook_sheet_names(path)
    labels_index = sheet_names.index("Labels") + 1
    with zipfile.ZipFile(path) as archive:
        xml = archive.read(f"xl/worksheets/sheet{labels_index}.xml").decode("utf-8")

    assert "Notes:" not in xml
    assert '<mergeCell ref="E7:F9"/>' not in xml
    assert 'r="E7"' not in xml
    assert 'r="F7"' not in xml
    assert 'r="E8"' not in xml
    assert 'r="F8"' not in xml
    assert 'r="E9"' not in xml
    assert 'r="F9"' not in xml


def test_labels_sheet_top_right_header_uses_country_code_only_in_ef_merge(tmp_path):
    path = tmp_path / "report.xlsx"

    write_workbook(
        str(path),
        labels_rows=[
            {
                "Label Number": "42",
                "Barcode/QR Value": "OPR 42",
                "Country Package Code": "DE  2-1",
                "Country Name Chinese": "\u5fb7\u56fd",
                "Items to Pack Column 1": "CORE x1",
            },
        ],
    )

    xml = _worksheet_xml(path, "Labels")

    assert 'r="E1" t="inlineStr" s="25"><is><t>DE</t></is></c>' in xml
    assert 'r="F1" t="inlineStr" s="25"><is><t></t></is></c>' in xml
    assert '<mergeCell ref="A1:D1"/>' in xml
    assert '<mergeCell ref="E1:F1"/>' in xml
    assert "DE  2-1" not in xml
    assert "UN  2-1" not in xml
    assert "\u5fb7\u56fd" in xml


def test_labels_sheet_splits_long_address_one_before_notes_box(tmp_path):
    path = tmp_path / "report.xlsx"
    long_address = "no.18,4/F,Block B,Hi-tech Industrial Centre,491-501Castle Peak Road Tsuen Wan"

    write_workbook(
        str(path),
        vfi_intake_form_rows=[{"Address Line 1": long_address}],
        labels_rows=[
            {
                "Label Number": "39",
                "Barcode/QR Value": "OPR 39",
                "Country Code": "HK",
                "Address Line 1": long_address,
                "Address Line 2": "",
                "Notes": "Leave package at front desk",
                "Items to Pack Column 1": "CORE x1",
            },
        ],
    )

    sheet_names = _workbook_sheet_names(path)
    labels_index = sheet_names.index("Labels") + 1
    with zipfile.ZipFile(path) as archive:
        xml = archive.read(f"xl/worksheets/sheet{labels_index}.xml").decode("utf-8")

    intake_rows = next(sheet.rows for sheet in read_workbook(str(path)) if sheet.sheet_name == "VFI Intake Form")

    assert "no.18,4/F,Block B,Hi-tech Industrial Centre," in xml
    assert "491-501Castle Peak Road Tsuen Wan" in xml
    assert "Notes: Leave package at front desk" in xml
    assert '<mergeCell ref="E7:F9"/>' in xml
    assert intake_rows[0]["Address Line 1"] == long_address


def test_labels_sheet_preserves_existing_address_two_and_prioritizes_address_over_notes(tmp_path):
    path = tmp_path / "report.xlsx"
    long_address = "no.18,4/F,Block B,Hi-tech Industrial Centre,491-501Castle Peak Road Tsuen Wan"

    write_workbook(
        str(path),
        labels_rows=[
            {
                "Label Number": "39",
                "Barcode/QR Value": "OPR 39",
                "Country Code": "HK",
                "Address Line 1": long_address,
                "Address Line 2": "Receiving office beside the rear service entrance",
                "Notes": "Call customer before delivery",
                "Items to Pack Column 1": "CORE x1",
            },
        ],
    )

    sheet_names = _workbook_sheet_names(path)
    labels_index = sheet_names.index("Labels") + 1
    with zipfile.ZipFile(path) as archive:
        xml = archive.read(f"xl/worksheets/sheet{labels_index}.xml").decode("utf-8")

    assert "no.18,4/F,Block B,Hi-tech Industrial Centre," in xml
    assert "491-501Castle Peak Road Tsuen Wan, Receiving office beside the rear service entrance" in xml
    assert "Notes:" not in xml
    assert '<mergeCell ref="E7:F9"/>' not in xml
    assert 'r="E7"' not in xml


def test_labels_sheet_keeps_short_address_lines_unchanged(tmp_path):
    path = tmp_path / "report.xlsx"

    write_workbook(
        str(path),
        labels_rows=[
            {
                "Label Number": "39",
                "Barcode/QR Value": "OPR 39",
                "Address Line 1": "64 BATKIN ROAD",
                "Address Line 2": "Unit 2",
                "Items to Pack Column 1": "CORE x1",
            },
        ],
    )

    sheet_names = _workbook_sheet_names(path)
    labels_index = sheet_names.index("Labels") + 1
    with zipfile.ZipFile(path) as archive:
        xml = archive.read(f"xl/worksheets/sheet{labels_index}.xml").decode("utf-8")

    assert "64 BATKIN ROAD" in xml
    assert "Unit 2" in xml


def test_labels_sheet_prints_continuation_label_for_overflow_items(tmp_path):
    path = tmp_path / "report.xlsx"

    write_workbook(
        str(path),
        labels_rows=[
            {
                "Label Number": "39",
                "Barcode/QR Value": "OPR 39",
                "Factory Name": "Longhai Printworks",
                "Country Code": "NZ",
                "From": "No.23 Baosheng Rd.,Bld2, 3rd Floor, Longhai, Fujian, China 363107",
                "To Name": "Ada Lovelace",
                "Address Line 1": "64 BATKIN ROAD",
                "Carton Box Designation": "VB 25",
                "Items to Pack Column 1": "CORE x1",
            },
            {
                "Label Number": "39 CONTINUED",
                "Barcode/QR Value": "OPR 39",
                "Factory Name": "Longhai Printworks",
                "Country Code": "NZ",
                "Label Continuation": True,
                "Original Label Number": "39",
                "Pledge Configuration": "4",
                "Items to Pack Column 1": "(7)  EXTRA-1 x1\n(8)  EXTRA-2 x1\n(9)  EXTRA-3 x1",
            },
        ],
    )

    sheet_names = _workbook_sheet_names(path)
    labels_index = sheet_names.index("Labels") + 1
    with zipfile.ZipFile(path) as archive:
        xml = archive.read(f"xl/worksheets/sheet{labels_index}.xml").decode("utf-8")
        names = set(archive.namelist())
        drawing_xml = archive.read("xl/drawings/drawing1.xml").decode("utf-8")
        styles_xml = archive.read("xl/styles.xml").decode("utf-8")

    assert "VFI #" not in xml
    assert "39 CONTINUED" not in xml
    assert "Continuation for" in xml
    assert "CONTINUED  Config: 4" in xml
    assert "(7)  EXTRA-1" in xml
    assert "(8)  EXTRA-2" in xml
    assert "(9)  EXTRA-3" in xml
    assert '<c r="A24" t="inlineStr" s="13"><is><t>1</t></is></c>' in xml
    assert '<c r="D24" t="inlineStr" s="13"><is><t>1</t></is></c>' in xml
    assert '<c r="A25" t="inlineStr" s="13"><is><t>1</t></is></c>' in xml
    assert '<c r="D25" t="inlineStr" s="13"><is><t></t></is></c>' in xml
    assert xml.count("64 BATKIN ROAD") == 1
    assert '<c r="A21" t="inlineStr" s="24"><is><t>39</t></is></c>' in xml
    assert '<c r="B21" t="inlineStr" s="24"><is><t></t></is></c>' in xml
    assert '<c r="C21" t="inlineStr" s="24"><is><t></t></is></c>' in xml
    assert '<c r="D21" t="inlineStr" s="24"><is><t></t></is></c>' in xml
    assert '<c r="E21" t="inlineStr" s="25"><is><t>NZ</t></is></c>' in xml
    assert '<c r="F21" t="inlineStr" s="25"><is><t></t></is></c>' in xml
    assert '<c r="B22" t="inlineStr" s="2"><is><t>39  CONTINUED  Config: 4</t></is></c>' in xml
    assert '<mergeCell ref="A21:D21"/>' in xml
    assert '<mergeCell ref="E21:F21"/>' in xml
    assert '<mergeCell ref="A21:E21"/>' not in xml
    assert '<mergeCell ref="C21:E21"/>' not in xml
    assert '<mergeCell ref="C21:D21"/>' not in xml
    assert '<mergeCell ref="B23:C23"/>' in xml
    assert '<mergeCell ref="E23:F23"/>' in xml
    assert '<mergeCell ref="B24:C24"/>' in xml
    assert '<mergeCell ref="E24:F24"/>' in xml
    assert '<mergeCell ref="B25:C25"/>' in xml
    assert '<mergeCell ref="E25:F25"/>' in xml
    assert "xl/media/label_qr_1.png" in names
    assert "xl/media/label_qr_2.png" not in names
    assert "QR Code 1" in drawing_xml
    assert "QR Code 2" not in drawing_xml
    assert '<alignment horizontal="right"/>' in styles_xml
    assert "<rowBreaks" in xml


def test_labels_sheet_embeds_qr_pngs_for_barcode_values(tmp_path):
    path = tmp_path / "report.xlsx"

    write_workbook(
        str(path),
        labels_rows=[
            {"Label Number": "39", "Barcode/QR Value": "OPR 39", "From": "VFI"},
        ],
    )

    sheet_names = _workbook_sheet_names(path)
    labels_index = sheet_names.index("Labels") + 1
    with zipfile.ZipFile(path) as archive:
        labels_xml = archive.read(f"xl/worksheets/sheet{labels_index}.xml").decode("utf-8")
        drawing_xml = archive.read("xl/drawings/drawing1.xml").decode("utf-8")
        drawing_rels = archive.read("xl/drawings/_rels/drawing1.xml.rels").decode("utf-8")
        content_types = archive.read("[Content_Types].xml").decode("utf-8")
        qr_png = archive.read("xl/media/label_qr_1.png")

    assert '<drawing r:id="rId1"/>' in labels_xml
    assert '<xdr:col>3</xdr:col>' in drawing_xml
    assert '<xdr:row>0</xdr:row>' in drawing_xml
    assert '<xdr:colOff>457200</xdr:colOff>' in drawing_xml
    assert '<xdr:rowOff>57150</xdr:rowOff>' in drawing_xml
    assert '<xdr:ext cx="1524000" cy="1524000"/>' in drawing_xml
    assert 'Target="../media/label_qr_1.png"' in drawing_rels
    assert 'ContentType="image/png"' in content_types
    assert qr_png.startswith(b"\x89PNG\r\n\x1a\n")
    assert struct.unpack(">I", qr_png[16:20])[0] >= 200
