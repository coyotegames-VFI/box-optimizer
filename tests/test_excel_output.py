import zipfile
from xml.etree import ElementTree

from box_optimizer.io.excel_reader import read_workbook
from box_optimizer.io.excel_writer import write_workbook


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
        "VFI Intake Form",
        "Optimized to Pack",
        "Label generator",
        "Labels",
        "Order Volume Weights",
        "Box Size Summary",
    ]
    assert sheet_names[8:] == ["Unmatched SKUs"]


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
        "VFI Intake Form",
        "Optimized to Pack",
        "Label generator",
        "Labels",
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
