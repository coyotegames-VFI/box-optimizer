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

    assert sheet_names[:3] == [
        "Summary",
        "Order Volume Weights",
        "Box Size Summary",
    ]
    assert sheet_names[3:] == ["Unmatched SKUs"]


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

    rows = read_workbook(str(path))[1].rows
    headers = list(rows[0].keys())

    assert headers[:20] == [
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
        "Box Type",
        "Assigned Box Length cm",
        "Assigned Box Width cm",
        "Assigned Box Height cm",
        "Box Standardization Note",
        "Distinct SKUs",
        "SKU Breakdown",
        "Box Plan",
        "Warning Summary",
    ]
    assert headers[20:] == ["Pledge Level", "Shipping Notes"]
    assert rows[0]["Pledge Level"] == "Deluxe"
    assert rows[0]["Shipping Notes"] == "Leave at door"


def test_empty_order_volume_weights_still_has_required_headers(tmp_path):
    path = tmp_path / "report.xlsx"

    write_workbook(str(path))

    with zipfile.ZipFile(path) as archive:
        xml = archive.read("xl/worksheets/sheet2.xml").decode("utf-8")

    assert "Region" in xml
    assert "SKU Breakdown" in xml
