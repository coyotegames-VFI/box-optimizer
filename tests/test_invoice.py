import json
from pathlib import Path

from box_optimizer.invoice import pay_to_lines, pay_to_uses_display_lines
from box_optimizer.workflow import _config, _invoice_payload


FAKE_US_PAY_TO_ENV = {
    "INVOICE_PAY_TO_US_ACCOUNT_HOLDER": "TEST US ACCOUNT HOLDER",
    "INVOICE_PAY_TO_US_BANK_NAME": "TEST US BANK",
    "INVOICE_PAY_TO_US_ACCOUNT_NUMBER": "TEST-US-ACCOUNT-0000",
    "INVOICE_PAY_TO_US_ROUTING_NUMBER": "TEST-US-ROUTING-0000",
    "INVOICE_PAY_TO_US_SWIFT_CODE": "TESTUSXX",
    "INVOICE_PAY_TO_US_IBAN": "TEST-US-IBAN-0000",
    "INVOICE_PAY_TO_US_ADDRESS": "TEST US ADDRESS",
    "INVOICE_PAY_TO_US_EMAIL": "test-us@example.invalid",
    "INVOICE_PAY_TO_US_EXTRA_LINES": "TEST US EXTRA LINE 1|TEST US EXTRA LINE 2",
}


FAKE_CN_PAY_TO_ENV = {
    "INVOICE_PAY_TO_CN_ACCOUNT_HOLDER": "TEST CN ACCOUNT HOLDER",
    "INVOICE_PAY_TO_CN_BANK_NAME": "TEST CN BANK",
    "INVOICE_PAY_TO_CN_ACCOUNT_NUMBER": "TEST-CN-ACCOUNT-0000",
    "INVOICE_PAY_TO_CN_SWIFT_CODE": "TESTCNXX",
    "INVOICE_PAY_TO_CN_CNAPS": "TEST-CN-CNAPS-0000",
    "INVOICE_PAY_TO_CN_BENEFICIARY_ADDRESS": "TEST CN BENEFICIARY ADDRESS",
    "INVOICE_PAY_TO_CN_POSTAL_CODE": "TEST-CN-POSTAL",
    "INVOICE_PAY_TO_CN_PHONE": "TEST-CN-PHONE",
    "INVOICE_PAY_TO_CN_EXTRA_LINES": "TEST CN EXTRA LINE 1|TEST CN EXTRA LINE 2",
}


FAKE_US_DISPLAY_LINES_ENV = {
    "INVOICE_PAY_TO_US_DISPLAY_LINES": json.dumps(
        [
            "TEST US Payment Method Header",
            "Label: TEST US VALUE",
            "",
            "TEST US Second Payment Method",
            "Label: TEST US VALUE 2",
        ]
    )
}


FAKE_CN_DISPLAY_LINES_ENV = {
    "INVOICE_PAY_TO_CN_DISPLAY_LINES": json.dumps(
        [
            "TEST CN Payment Method Header",
            "Label: TEST CN VALUE",
            "",
            "TEST CN Second Payment Method",
            "Label: TEST CN VALUE 2",
        ]
    )
}


def _payload(config, metadata=None, cost_rows=None, ship_order_count=1):
    payload, warning = _invoice_payload(
        _config(config),
        intake_summary_metadata=metadata or {},
        cost_summary_rows=cost_rows
        or [
            {
                "Backer ID": "B1",
                "VFI #": "VFI 1",
                "Final weight kg": "",
                "Final cost": "",
                "Scan note": "",
            }
        ],
        cost_sheet_name="Cost Summary",
        ship_order_count=ship_order_count,
    )
    assert warning == ""
    return payload


def test_config_defaults_do_not_include_invoice():
    cfg = _config({})
    assert cfg["include_invoice"] is False
    assert cfg["invoice_variant"] == "US"


def test_invoice_payload_uses_invoices_to_before_company_fallback():
    payload = _payload(
        {"include_invoice": True, "company": "Fallback Company"},
        {"Invoices To": "Client Finance"},
    )

    assert payload.bill_to == "Client Finance"


def test_invoice_payload_pulls_email_from_intake_metadata():
    payload = _payload(
        {"include_invoice": True},
        {"Email": "client@example.invalid", "Accounting Email": "accounting@example.invalid"},
    )

    assert payload.email == "client@example.invalid"


def test_invoice_payload_pulls_address_lines_from_intake_metadata():
    payload = _payload(
        {"include_invoice": True},
        {
            "Address Line 1": "123 Client Street",
            "Address Line 2": "Suite 400",
            "Postal Code": "99999",
            "Country": "United States",
        },
    )

    assert payload.address_lines == ("123 Client Street", "Suite 400", "99999", "United States")


def test_invoice_payload_missing_email_and_address_remain_blank_without_blocking_generation():
    payload = _payload({"include_invoice": True}, {"Invoices To": "Client Finance"})

    assert payload is not None
    assert payload.email == ""
    assert payload.address_lines == ("", "")


def test_invoice_payload_falls_back_to_company_client_or_publisher():
    payload = _payload({"include_invoice": True, "client": "Client Co"}, {})

    assert payload.bill_to == "Client Co"


def test_invoice_payload_leaves_bill_to_blank_when_no_source_exists():
    payload = _payload({"include_invoice": True}, {})

    assert payload.bill_to == ""


def test_invoice_payload_uses_backer_order_count_not_carton_count():
    payload = _payload({"include_invoice": True}, ship_order_count=1)

    assert payload.ship_order_count == 1


def test_invoice_payload_detects_final_cost_final_weight_and_scan_note_columns_by_header():
    payload = _payload(
        {"include_invoice": True},
        cost_rows=[
            {
                "Backer ID": "B1",
                "Quoted shipping cost": 999,
                "Scan note": "",
                "Final cost": "",
                "Final weight kg": "",
            }
        ],
    )

    assert payload.scan_note_column == "C"
    assert payload.final_cost_column == "D"
    assert payload.final_weight_column == "E"


def test_invoice_payload_parses_inbound_fee_when_available_and_leaves_blank_otherwise():
    with_fee = _payload({"include_invoice": True}, {"Inbound Fee": "$12.50"})
    with_extra_decimals = _payload({"include_invoice": True}, {"Inbound Fee": "$12.349"})
    without_fee = _payload({"include_invoice": True}, {})

    assert with_fee.inbound_fee == 12.5
    assert with_extra_decimals.inbound_fee == 12.34
    assert without_fee.inbound_fee == ""


def test_invoice_payload_detects_canada_and_mx_manual_charge_flags_from_country():
    canada = _payload({"include_invoice": True}, cost_rows=[{"Country": "Canada", "Final cost": ""}])
    mexico = _payload({"include_invoice": True}, cost_rows=[{"Country": "Mexico", "Final cost": ""}])
    mx = _payload({"include_invoice": True}, cost_rows=[{"Country": "MX", "Final cost": ""}])
    both = _payload(
        {"include_invoice": True},
        cost_rows=[{"Country": "Canada", "Final cost": ""}, {"Country": "MX", "Final cost": ""}],
    )
    neither = _payload({"include_invoice": True}, cost_rows=[{"Country": "United States", "Final cost": ""}])

    assert canada.include_canada_ocean_tax is True
    assert canada.include_mx_import_tax is False
    assert mexico.include_canada_ocean_tax is False
    assert mexico.include_mx_import_tax is True
    assert mx.include_mx_import_tax is True
    assert both.include_canada_ocean_tax is True
    assert both.include_mx_import_tax is True
    assert neither.include_canada_ocean_tax is False
    assert neither.include_mx_import_tax is False


def test_invalid_invoice_variant_skips_invoice_and_returns_warning():
    payload, warning = _invoice_payload(
        _config({"include_invoice": True, "invoice_variant": "EU"}),
        intake_summary_metadata={},
        cost_summary_rows=[],
        cost_sheet_name="Cost Summary",
        ship_order_count=1,
    )

    assert payload is None
    assert "Invalid invoice_variant" in warning


def test_us_pay_to_variant_uses_us_runtime_placeholder_env_fields(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    lines, incomplete = pay_to_lines(
        "US",
        environ={"INVOICE_PAY_TO_US_ACCOUNT_HOLDER": "US_TEST_ACCOUNT"},
    )

    assert ("Account holder", "US_TEST_ACCOUNT") in lines
    assert incomplete is True


def test_cn_pay_to_variant_uses_cn_runtime_placeholder_env_fields(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    lines, incomplete = pay_to_lines(
        "CN",
        environ={"INVOICE_PAY_TO_CN_ACCOUNT_HOLDER": "CN_TEST_ACCOUNT"},
    )

    assert ("Account holder", "CN_TEST_ACCOUNT") in lines
    assert incomplete is True


def test_us_pay_to_complete_fake_env_clears_incomplete_warning(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    lines, incomplete = pay_to_lines("US", environ=FAKE_US_PAY_TO_ENV)

    assert incomplete is False
    assert ("Account holder", "TEST US ACCOUNT HOLDER") in lines
    assert ("Bank name", "TEST US BANK") in lines
    assert ("Account number", "TEST-US-ACCOUNT-0000") in lines
    assert ("Routing number", "TEST-US-ROUTING-0000") in lines
    assert ("SWIFT code", "TESTUSXX") in lines
    assert ("IBAN", "TEST-US-IBAN-0000") in lines
    assert ("Address", "TEST US ADDRESS") in lines
    assert ("Email", "test-us@example.invalid") in lines
    assert ("", "TEST US EXTRA LINE 1") in lines
    assert ("", "TEST US EXTRA LINE 2") in lines
    assert not (tmp_path / "local_reference" / "invoice_config.json").exists()


def test_cn_pay_to_complete_fake_env_clears_incomplete_warning(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    lines, incomplete = pay_to_lines("CN", environ=FAKE_CN_PAY_TO_ENV)

    assert incomplete is False
    assert ("Account holder", "TEST CN ACCOUNT HOLDER") in lines
    assert ("Bank name", "TEST CN BANK") in lines
    assert ("Account number", "TEST-CN-ACCOUNT-0000") in lines
    assert ("SWIFT code", "TESTCNXX") in lines
    assert ("CNAPS", "TEST-CN-CNAPS-0000") in lines
    assert ("Beneficiary address", "TEST CN BENEFICIARY ADDRESS") in lines
    assert ("Postal code", "TEST-CN-POSTAL") in lines
    assert ("Phone", "TEST-CN-PHONE") in lines
    assert ("", "TEST CN EXTRA LINE 1") in lines
    assert ("", "TEST CN EXTRA LINE 2") in lines
    assert not (tmp_path / "local_reference" / "invoice_config.json").exists()


def test_pay_to_lines_supports_local_reference_invoice_config_structure(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    config_path = tmp_path / "local_reference" / "invoice_config.json"
    config_path.parent.mkdir()
    config_path.write_text(
        """{
  "US": {
    "account_holder": "TEST US ACCOUNT HOLDER",
    "bank_name": "TEST US BANK",
    "account_number": "TEST-US-ACCOUNT-0000",
    "routing_number": "TEST-US-ROUTING-0000",
    "swift_code": "TESTUSXX",
    "iban": "TEST-US-IBAN-0000",
    "address": "TEST US ADDRESS",
    "email": "test-us@example.invalid",
    "extra_lines": "TEST US EXTRA LINE 1|TEST US EXTRA LINE 2"
  },
  "CN": {
    "account_holder": "TEST CN ACCOUNT HOLDER",
    "bank_name": "TEST CN BANK",
    "account_number": "TEST-CN-ACCOUNT-0000",
    "swift_code": "TESTCNXX",
    "cnaps": "TEST-CN-CNAPS-0000",
    "beneficiary_address": "TEST CN BENEFICIARY ADDRESS",
    "postal_code": "TEST-CN-POSTAL",
    "phone": "TEST-CN-PHONE",
    "extra_lines": "TEST CN EXTRA LINE 1|TEST CN EXTRA LINE 2"
  }
}
""",
        encoding="utf-8",
    )

    us_lines, us_incomplete = pay_to_lines("US", environ={})
    cn_lines, cn_incomplete = pay_to_lines("CN", environ={})

    assert us_incomplete is False
    assert cn_incomplete is False
    assert ("Account holder", "TEST US ACCOUNT HOLDER") in us_lines
    assert ("Account holder", "TEST CN ACCOUNT HOLDER") in cn_lines
    assert ("", "TEST US EXTRA LINE 2") in us_lines
    assert ("", "TEST CN EXTRA LINE 2") in cn_lines


def test_us_env_display_lines_render_and_suppress_generic_labels(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    lines, incomplete = pay_to_lines("US", environ={**FAKE_US_PAY_TO_ENV, **FAKE_US_DISPLAY_LINES_ENV})

    assert lines == (
        ("TEST US Payment Method Header", ""),
        ("Label: TEST US VALUE", ""),
        ("", ""),
        ("TEST US Second Payment Method", ""),
        ("Label: TEST US VALUE 2", ""),
    )
    assert incomplete is False
    assert ("Account holder", "TEST US ACCOUNT HOLDER") not in lines
    assert pay_to_uses_display_lines("US", environ=FAKE_US_DISPLAY_LINES_ENV) is True


def test_cn_env_display_lines_render_and_suppress_generic_labels(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    lines, incomplete = pay_to_lines("CN", environ={**FAKE_CN_PAY_TO_ENV, **FAKE_CN_DISPLAY_LINES_ENV})

    assert lines == (
        ("TEST CN Payment Method Header", ""),
        ("Label: TEST CN VALUE", ""),
        ("", ""),
        ("TEST CN Second Payment Method", ""),
        ("Label: TEST CN VALUE 2", ""),
    )
    assert incomplete is False
    assert ("Account holder", "TEST CN ACCOUNT HOLDER") not in lines
    assert pay_to_uses_display_lines("CN", environ=FAKE_CN_DISPLAY_LINES_ENV) is True


def test_env_display_lines_take_precedence_over_local_config_display_lines(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    config_path = tmp_path / "local_reference" / "invoice_config.json"
    config_path.parent.mkdir()
    config_path.write_text(
        """{
  "US": {
    "display_lines": ["TEST LOCAL DISPLAY LINE"],
    "account_holder": "TEST LOCAL ACCOUNT HOLDER"
  }
}
""",
        encoding="utf-8",
    )

    lines, incomplete = pay_to_lines("US", environ=FAKE_US_DISPLAY_LINES_ENV)

    assert lines[0] == ("TEST US Payment Method Header", "")
    assert ("TEST LOCAL DISPLAY LINE", "") not in lines
    assert ("Account holder", "TEST LOCAL ACCOUNT HOLDER") not in lines
    assert incomplete is False


def test_invalid_env_display_lines_fall_back_without_crashing(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    config_path = tmp_path / "local_reference" / "invoice_config.json"
    config_path.parent.mkdir()
    config_path.write_text(
        """{
  "US": {
    "display_lines": ["TEST LOCAL DISPLAY LINE"]
  }
}
""",
        encoding="utf-8",
    )

    json_lines, json_incomplete = pay_to_lines(
        "US",
        environ={"INVOICE_PAY_TO_US_DISPLAY_LINES": "not-json"},
    )
    non_string_lines, non_string_incomplete = pay_to_lines(
        "US",
        environ={"INVOICE_PAY_TO_US_DISPLAY_LINES": json.dumps(["TEST ENV LINE", 123])},
    )

    assert json_lines == (("TEST LOCAL DISPLAY LINE", ""),)
    assert json_incomplete is False
    assert non_string_lines == (("TEST LOCAL DISPLAY LINE", ""),)
    assert non_string_incomplete is False
    assert pay_to_uses_display_lines(
        "US",
        environ={"INVOICE_PAY_TO_US_DISPLAY_LINES": "not-json"},
    ) is True


def test_blank_env_display_lines_leave_pay_to_incomplete_warning(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    empty_lines, empty_incomplete = pay_to_lines(
        "US",
        environ={"INVOICE_PAY_TO_US_DISPLAY_LINES": "[]", **FAKE_US_PAY_TO_ENV},
    )
    blank_lines, blank_incomplete = pay_to_lines(
        "CN",
        environ={"INVOICE_PAY_TO_CN_DISPLAY_LINES": json.dumps(["", "   "]), **FAKE_CN_PAY_TO_ENV},
    )

    assert empty_lines == ()
    assert empty_incomplete is True
    assert blank_lines == (("", ""), ("   ", ""))
    assert blank_incomplete is True
    assert pay_to_uses_display_lines("US", environ={"INVOICE_PAY_TO_US_DISPLAY_LINES": "[]"}) is True


def test_us_pay_to_display_lines_take_precedence_over_generic_labels(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    config_path = tmp_path / "local_reference" / "invoice_config.json"
    config_path.parent.mkdir()
    config_path.write_text(
        """{
  "US": {
    "display_lines": [
      "TEST US Payment Method 1",
      "TEST US line 1",
      "",
      "TEST US Payment Method 2",
      "TEST US line 2"
    ],
    "account_holder": "TEST GENERIC ACCOUNT HOLDER"
  }
}
""",
        encoding="utf-8",
    )

    lines, incomplete = pay_to_lines("US", environ={})

    assert lines == (
        ("TEST US Payment Method 1", ""),
        ("TEST US line 1", ""),
        ("", ""),
        ("TEST US Payment Method 2", ""),
        ("TEST US line 2", ""),
    )
    assert incomplete is False
    assert ("Account holder", "TEST GENERIC ACCOUNT HOLDER") not in lines
    assert pay_to_uses_display_lines("US") is True


def test_cn_pay_to_display_lines_clear_incomplete_warning(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    lines, incomplete = pay_to_lines(
        "CN",
        config={
            "invoice_pay_to": {
                "CN": {
                    "display_lines": [
                        "TEST CN T/T Bank Transfer",
                        "TEST CN line 1",
                    ]
                }
            }
        },
        environ={},
    )

    assert lines == (("TEST CN T/T Bank Transfer", ""), ("TEST CN line 1", ""))
    assert incomplete is False
    assert pay_to_uses_display_lines(
        "CN",
        config={
            "invoice_pay_to": {
                "CN": {
                    "display_lines": [
                        "TEST CN T/T Bank Transfer",
                        "TEST CN line 1",
                    ]
                }
            }
        },
    ) is True


def test_pay_to_payment_blocks_flatten_to_display_lines_with_blank_separator(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    lines, incomplete = pay_to_lines(
        "US",
        config={
            "invoice_pay_to": {
                "US": {
                    "payment_blocks": [
                        {"title": "TEST US Method A", "lines": ["TEST US A1"]},
                        {"title": "TEST US Method B", "lines": ["TEST US B1", "TEST US B2"]},
                    ]
                }
            }
        },
        environ={},
    )

    assert lines == (
        ("TEST US Method A", ""),
        ("TEST US A1", ""),
        ("", ""),
        ("TEST US Method B", ""),
        ("TEST US B1", ""),
        ("TEST US B2", ""),
    )
    assert incomplete is False


def test_pay_to_incomplete_warning_remains_for_incomplete_fallback(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)

    _lines, incomplete = pay_to_lines("US", environ={"INVOICE_PAY_TO_US_ACCOUNT_HOLDER": "TEST ONLY"})

    assert incomplete is True
    assert pay_to_uses_display_lines("US") is False


def test_fake_pay_to_fixtures_do_not_use_real_sensitive_values():
    all_values = [*FAKE_US_PAY_TO_ENV.values(), *FAKE_CN_PAY_TO_ENV.values()]

    assert all("TEST" in value or value.endswith("@example.invalid") for value in all_values)
    assert any("TEST-US-ACCOUNT-0000" in value for value in all_values)
    assert any("TEST-CN-ACCOUNT-0000" in value for value in all_values)


def test_local_reference_remains_gitignored():
    assert "local_reference/" in Path(".gitignore").read_text(encoding="utf-8")


def test_invoice_tests_do_not_contain_sensitive_fixture_values():
    fixture_values = {"US_TEST_ACCOUNT", "CN_TEST_ACCOUNT"}
    assert fixture_values == {"US_TEST_ACCOUNT", "CN_TEST_ACCOUNT"}
