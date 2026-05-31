import zipfile
import struct
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
    cell_formats = styles.findall("main:cellXfs/main:xf", NS)
    assert cell_formats[9].attrib["numFmtId"] == "164"
    assert cell_formats[9].attrib["applyNumberFormat"] == "1"
    assert summary.find(".//main:c[@r='C2']", NS).attrib["s"] == "9"
    assert cost_summary.find(".//main:c[@r='B2']", NS).attrib["s"] == "9"
    assert cost_summary.find(".//main:c[@r='C2']", NS).attrib["s"] == "9"

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

    assert "VFI #" in xml
    assert "Longhai Printworks" in xml
    assert "Campaign</t>" not in xml
    assert 'r="C1" t="inlineStr" s="7"><is><t>NZ</t></is></c>' in xml
    assert 'r="E1" t="inlineStr" s="7"><is><t></t></is></c>' in xml
    assert 'r="F1" t="inlineStr" s="7"><is><t></t></is></c>' in xml
    assert "Barcode / QR Value" not in xml
    assert "Country CN" not in xml
    assert "Total Value USD" not in xml
    assert "OPR 39" in xml
    assert "No.23 Baosheng Rd.,Bld2, 3rd Floor, Longhai" in xml
    assert "Fujian, China 363107" in xml
    assert "Ada Lovelace  Backer ID B-1" in xml
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
    assert 'r="F3"' in xml
    assert 'r="B3"' not in xml
    assert 'r="C7"' not in xml
    assert 'r="D7"' not in xml
    assert 'r="F10"' not in xml
    assert 'r="A1" t="inlineStr" s="7"' in xml
    assert 'r="A12" t="inlineStr" s="10"' in xml
    assert 'r="A13" t="inlineStr" s="11"' in xml
    assert 'r="B13" t="inlineStr" s="12"' in xml
    assert 'r="A14" t="inlineStr" s="13"' in xml
    assert 'r="B14" t="inlineStr" s="14"' in xml
    assert 'r="A16" t="inlineStr" s="2"' in xml
    assert 'r="A17" t="inlineStr" s="2"' in xml
    assert 'r="A19" t="inlineStr" s="7"' in xml
    assert 'r="A20" t="inlineStr" s="7"><is><t>Item Count: 4</t></is></c>' in xml
    assert 'r="C20" t="inlineStr" s="7"><is><t>Longhai Printworks</t></is></c>' in xml
    assert 'r="E20" t="inlineStr" s="7"><is><t>VB 25</t></is></c>' in xml
    assert 'r="B16"' not in xml
    assert '<font><b/><sz val="20"/><name val="Calibri"/></font>' in styles_xml
    assert '<xf numFmtId="0" fontId="6" fillId="0" borderId="0" xfId="0" applyFont="1"/>' in styles_xml
    assert '<borders count="3">' in styles_xml
    assert '<top style="thick"><color auto="1"/></top><bottom style="thick"><color auto="1"/></bottom>' in styles_xml
    assert '<left style="thick"><color auto="1"/></left><right style="thick"><color auto="1"/></right>' in styles_xml
    assert 's="11"' in xml
    assert 's="12"' in xml
    assert "<pageSetup" in xml
    assert "<rowBreaks" in xml
    assert '<mergeCell ref="A12:F12"/>' in xml
    assert '<mergeCell ref="B13:C13"/>' in xml
    assert '<mergeCell ref="E13:F13"/>' in xml
    assert '<mergeCell ref="B14:C14"/>' in xml
    assert '<mergeCell ref="E14:F14"/>' in xml
    assert '<mergeCell ref="E10:F10"/>' in xml
    assert '<mergeCell ref="A19:F19"/>' in xml
    assert '<mergeCell ref="A20:B20"/>' in xml
    assert '<mergeCell ref="C20:D20"/>' in xml
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
    assert rows[0].attrib["ht"] == "20"
    assert rows[5].attrib["ht"] == "10"
    assert rows[17].attrib["ht"] == "10"
    assert rows[18].attrib["ht"] == "20"
    assert rows[19].attrib["ht"] == "23"
    assert int(rows[-1].attrib["r"]) > 20
    assert len(cols) == 6
    assert [col.attrib["width"] for col in cols] == ["12", "22", "14", "14", "22", "14"]


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

    assert xml.count("VFI #") == 2
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
    assert '<c r="C21" t="inlineStr" s="7"><is><t>CONTINUED  Config: 4</t></is></c>' in xml
    assert '<c r="D21" t="inlineStr" s="7"><is><t></t></is></c>' in xml
    assert '<c r="E21" t="inlineStr" s="7"><is><t></t></is></c>' in xml
    assert '<c r="F21" t="inlineStr" s="7"><is><t>NZ</t></is></c>' in xml
    assert '<mergeCell ref="C21:E21"/>' in xml
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
    assert '<xdr:col>4</xdr:col>' in drawing_xml
    assert '<xdr:row>0</xdr:row>' in drawing_xml
    assert '<xdr:colOff>57150</xdr:colOff>' in drawing_xml
    assert '<xdr:rowOff>57150</xdr:rowOff>' in drawing_xml
    assert '<xdr:ext cx="1524000" cy="1524000"/>' in drawing_xml
    assert 'Target="../media/label_qr_1.png"' in drawing_rels
    assert 'ContentType="image/png"' in content_types
    assert qr_png.startswith(b"\x89PNG\r\n\x1a\n")
    assert struct.unpack(">I", qr_png[16:20])[0] >= 200
