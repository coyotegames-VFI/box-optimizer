import base64
import csv
import re
from io import StringIO
from pathlib import Path

from fastapi.testclient import TestClient

from box_optimizer.api import app


def _csv_bytes(rows: list[dict]) -> bytes:
    buffer = StringIO()
    writer = csv.DictWriter(buffer, fieldnames=list(rows[0].keys()))
    writer.writeheader()
    writer.writerows(rows)
    return buffer.getvalue().encode("utf-8")


def _sku_master_file():
    return (
        "sku_master.csv",
        _csv_bytes(
            [
                {
                    "SKU": "Core Game",
                    "Product Name": "Core Game",
                    "Length": "5",
                    "Width": "5",
                    "Height": "5",
                    "Weight kg": "1",
                }
            ]
        ),
        "text/csv",
    )


def _orders_file():
    return (
        "orders.csv",
        _csv_bytes(
            [
                {
                    "Order ID": "1001",
                    "SKU": "Core Game",
                    "Quantity": "1",
                    "Region": "NA",
                    "Country": "US",
                    "State": "CA",
                }
            ]
        ),
        "text/csv",
    )


def _orders_file_two_orders():
    return (
        "orders.csv",
        _csv_bytes(
            [
                {
                    "Order ID": "1001",
                    "SKU": "Core Game",
                    "Quantity": "1",
                    "Region": "NA",
                    "Country": "US",
                    "State": "CA",
                },
                {
                    "Order ID": "1002",
                    "SKU": "Core Game",
                    "Quantity": "1",
                    "Region": "NA",
                    "Country": "US",
                    "State": "NY",
                },
            ]
        ),
        "text/csv",
    )


def _base64_text(content: bytes) -> str:
    return base64.b64encode(content).decode("ascii")


def _base64_csv_payload(config_json=None) -> dict:
    sku_name, sku_content, _sku_type = _sku_master_file()
    orders_name, orders_content, _orders_type = _orders_file()
    payload = {
        "sku_master_filename": sku_name,
        "sku_master_base64": _base64_text(sku_content),
        "orders_filename": orders_name,
        "orders_base64": _base64_text(orders_content),
    }
    if config_json is not None:
        payload["config_json"] = config_json
    return payload


def _base64_example_workbook_payload(config_json=None) -> dict:
    root = Path(__file__).resolve().parents[1]
    payload = {
        "sku_master_filename": "sku_master.xlsx",
        "sku_master_base64": _base64_text((root / "examples" / "sku_master.xlsx").read_bytes()),
        "orders_filename": "orders.xlsx",
        "orders_base64": _base64_text((root / "examples" / "orders.xlsx").read_bytes()),
    }
    if config_json is not None:
        payload["config_json"] = config_json
    return payload


def _mixed_sku_master_file():
    return (
        "sku_master.csv",
        _csv_bytes(
            [
                {
                    "SKU": "Core Game",
                    "Product Name": "Core Game",
                    "Length": "5",
                    "Width": "5",
                    "Height": "5",
                    "Weight kg": "1",
                },
                {
                    "SKU": "Oversized",
                    "Product Name": "Oversized",
                    "Length": "80",
                    "Width": "10",
                    "Height": "10",
                    "Weight kg": "3",
                },
            ]
        ),
        "text/csv",
    )


def _mixed_orders_file():
    return (
        "orders.csv",
        _csv_bytes(
            [
                {
                    "Order ID": "1001",
                    "SKU": "Core Game",
                    "Quantity": "1",
                    "Region": "NA",
                    "Country": "US",
                    "State": "CA",
                },
                {
                    "Order ID": "1002",
                    "SKU": "Oversized",
                    "Quantity": "1",
                    "Region": "NA",
                    "Country": "US",
                    "State": "NY",
                },
            ]
        ),
        "text/csv",
    )


def test_health_returns_ok():
    client = TestClient(app)

    response = client.get("/health")

    assert response.status_code == 200
    assert response.json() == {"status": "ok"}


def test_openapi_json_is_exposed():
    client = TestClient(app)

    response = client.get("/openapi.json")

    assert response.status_code == 200
    assert response.json()["info"]["title"] == "box_optimizer"


def test_openapi_includes_base64_action_endpoints_with_explicit_schemas():
    client = TestClient(app)

    response = client.get("/openapi.json")

    assert response.status_code == 200
    schema = response.json()
    assert "/inspect_base64" in schema["paths"]
    assert "/optimize_base64" in schema["paths"]
    assert "/jobs/{job_id}" in schema["paths"]
    assert "/jobs/{job_id}/download" in schema["paths"]
    assert "status" in schema["components"]["schemas"]["HealthResponse"]["properties"]
    assert "git_commit" in schema["components"]["schemas"]["VersionResponse"]["properties"]
    assert "properties" in schema["components"]["schemas"]["InspectSummaryResponse"]
    assert "sku_items" in schema["components"]["schemas"]["InspectSummaryResponse"]["properties"]
    optimize_schema = schema["components"]["schemas"]["OptimizeBase64Response"]
    assert set(optimize_schema["properties"]) >= {"filename", "content_type", "workbook_base64", "summary"}
    request_schema = schema["components"]["schemas"]["Base64WorkbookRequest"]
    assert set(request_schema["properties"]) >= {
        "sku_master_filename",
        "sku_master_base64",
        "orders_filename",
        "orders_base64",
        "config_json",
    }


    job_schema = schema["components"]["schemas"]["JobStatusResponse"]
    assert set(job_schema["properties"]) >= {"job_id", "status", "created_at", "expires_at", "summary", "download_url", "error"}


def test_version_returns_json():
    client = TestClient(app)

    response = client.get("/version")

    assert response.status_code == 200
    body = response.json()
    assert body["app"] == "box_optimizer"
    assert body["version"] == "0.1.0"
    assert "timestamp" in body
    assert "git_commit" in body


def test_inspect_returns_counts(monkeypatch):
    monkeypatch.delenv("BOX_OPTIMIZER_API_KEY", raising=False)
    client = TestClient(app)

    response = client.post(
        "/inspect",
        files={
            "sku_master_file": _sku_master_file(),
            "orders_file": _orders_file(),
        },
    )

    assert response.status_code == 200
    body = response.json()
    assert body["sku_items"] == 1
    assert body["order_rows"] == 1
    assert body["order_lines"] == 1
    assert body["matched"] == 1
    assert body["unmatched"] == 0
    assert "elapsed_seconds" in body


def test_inspect_respects_max_orders(monkeypatch):
    monkeypatch.delenv("BOX_OPTIMIZER_API_KEY", raising=False)
    client = TestClient(app)

    response = client.post(
        "/inspect",
        files={
            "sku_master_file": _sku_master_file(),
            "orders_file": _orders_file_two_orders(),
        },
        data={"config_json": '{"max_orders": 1}'},
    )

    assert response.status_code == 200
    body = response.json()
    assert body["order_rows"] == 2
    assert body["order_lines"] == 1
    assert body["matched"] == 1


def test_config_json_accepts_sku_rules_and_inspect_reports_rule_keys(monkeypatch):
    monkeypatch.delenv("BOX_OPTIMIZER_API_KEY", raising=False)
    client = TestClient(app)

    response = client.post(
        "/inspect",
        files={
            "sku_master_file": _sku_master_file(),
            "orders_file": _orders_file(),
        },
        data={
            "config_json": (
                '{"sku_rules":{"Core Game":{"no_padding":true},'
                '"Missing Rule":{"prepacked":true}}}'
            ),
        },
    )

    assert response.status_code == 200
    body = response.json()
    assert body["matched_rule_keys"] == ["Core Game"]
    assert body["unmatched_rule_keys"] == ["Missing Rule"]


def test_inspect_base64_accepts_xlsx_workbooks(monkeypatch):
    monkeypatch.delenv("BOX_OPTIMIZER_API_KEY", raising=False)
    client = TestClient(app)

    response = client.post(
        "/inspect_base64",
        json=_base64_example_workbook_payload(
            {"max_orders": 1, "debug": True, "packing_mode": "fast"}
        ),
    )

    assert response.status_code == 200
    body = response.json()
    assert body["sku_items"] > 0
    assert body["order_rows"] > 0
    assert body["matched"] > 0
    assert "elapsed_seconds" in body


def test_inspect_base64_accepts_config_json_string(monkeypatch):
    monkeypatch.delenv("BOX_OPTIMIZER_API_KEY", raising=False)
    client = TestClient(app)

    response = client.post(
        "/inspect_base64",
        json=_base64_csv_payload('{"max_orders": 1, "packing_mode": "fast"}'),
    )

    assert response.status_code == 200
    body = response.json()
    assert body["order_lines"] == 1
    assert body["matched"] == 1


def test_optimize_base64_returns_json_workbook(monkeypatch):
    monkeypatch.delenv("BOX_OPTIMIZER_API_KEY", raising=False)
    client = TestClient(app)

    response = client.post(
        "/optimize_base64",
        json=_base64_example_workbook_payload(
            {
                "max_orders": 1,
                "packing_mode": "fast",
                "output_granularity": "order_summary",
                "preserve_region_sheets": False,
            }
        ),
    )

    assert response.status_code == 200
    body = response.json()
    assert body["filename"] == "optimized_shipping_plan.xlsx"
    assert body["content_type"] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    assert base64.b64decode(body["workbook_base64"])[:2] == b"PK"
    assert body["summary"]["orders_processed"] == 1
    assert "elapsed_seconds" in body["summary"]


def test_optimize_base64_uses_campaign_specific_download_filename(monkeypatch):
    monkeypatch.delenv("BOX_OPTIMIZER_API_KEY", raising=False)
    client = TestClient(app)

    response = client.post(
        "/optimize_base64",
        json=_base64_csv_payload(
            {
                "packing_mode": "fast",
                "campaign": {"name": "Dark Horizon", "code": 'OPR: Wave/2?'},
            }
        ),
    )

    assert response.status_code == 200
    assert response.json()["filename"] == "OPR Wave2 Shipping Plan.xlsx"


def test_optimize_base64_falls_back_to_default_download_filename(monkeypatch):
    monkeypatch.delenv("BOX_OPTIMIZER_API_KEY", raising=False)
    client = TestClient(app)

    response = client.post("/optimize_base64", json=_base64_csv_payload({"packing_mode": "fast"}))

    assert response.status_code == 200
    assert response.json()["filename"] == "optimized_shipping_plan.xlsx"


def test_inspect_base64_requires_api_key_when_environment_variable_is_set(monkeypatch):
    monkeypatch.setenv("BOX_OPTIMIZER_API_KEY", "secret")
    client = TestClient(app)

    response = client.post("/inspect_base64", json=_base64_csv_payload())

    assert response.status_code == 401


def test_optimize_base64_accepts_authorization_bearer_header(monkeypatch):
    monkeypatch.setenv("BOX_OPTIMIZER_API_KEY", " secret ")
    client = TestClient(app)

    response = client.post(
        "/optimize_base64",
        headers={"Authorization": "Bearer  secret "},
        json=_base64_csv_payload({"packing_mode": "fast"}),
    )

    assert response.status_code == 200
    assert base64.b64decode(response.json()["workbook_base64"])[:2] == b"PK"


def test_optimize_requires_api_key_when_environment_variable_is_set(monkeypatch):
    monkeypatch.setenv("BOX_OPTIMIZER_API_KEY", "secret")
    client = TestClient(app)

    response = client.post(
        "/optimize",
        files={
            "sku_master_file": _sku_master_file(),
            "orders_file": _orders_file(),
        },
    )

    assert response.status_code == 401


def test_optimize_rejects_invalid_api_key(monkeypatch):
    monkeypatch.setenv("BOX_OPTIMIZER_API_KEY", "secret")
    client = TestClient(app)

    response = client.post(
        "/optimize",
        headers={"X-API-Key": "wrong"},
        files={
            "sku_master_file": _sku_master_file(),
            "orders_file": _orders_file(),
        },
    )

    assert response.status_code == 401


def test_optimize_accepts_x_api_key_header(monkeypatch):
    monkeypatch.setenv("BOX_OPTIMIZER_API_KEY", " secret ")
    client = TestClient(app)

    response = client.post(
        "/optimize",
        headers={"X-API-Key": " secret "},
        files={
            "sku_master_file": _sku_master_file(),
            "orders_file": _orders_file(),
        },
        data={
            "config_json": '{"standardization_tolerance_cm": 2}',
        },
    )

    assert response.status_code == 200
    assert response.headers["content-type"].startswith(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    assert response.content[:2] == b"PK"


def test_optimize_file_response_uses_campaign_specific_download_filename(monkeypatch):
    monkeypatch.delenv("BOX_OPTIMIZER_API_KEY", raising=False)
    client = TestClient(app)

    response = client.post(
        "/optimize",
        files={
            "sku_master_file": _sku_master_file(),
            "orders_file": _orders_file(),
        },
        data={"config_json": '{"packing_mode":"fast","campaign":{"name":"Sordane"}}'},
    )

    assert response.status_code == 200
    assert "Sordane%20Shipping%20Plan.xlsx" in response.headers["content-disposition"]


def test_optimize_accepts_authorization_bearer_header(monkeypatch):
    monkeypatch.setenv("BOX_OPTIMIZER_API_KEY", " secret ")
    client = TestClient(app)

    response = client.post(
        "/optimize",
        headers={"Authorization": "Bearer  secret "},
        files={
            "sku_master_file": _sku_master_file(),
            "orders_file": _orders_file(),
        },
    )

    assert response.status_code == 200
    assert response.headers["content-type"].startswith(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    assert response.content[:2] == b"PK"


def test_inspect_accepts_x_api_key_header(monkeypatch):
    monkeypatch.setenv("BOX_OPTIMIZER_API_KEY", " secret ")
    client = TestClient(app)

    response = client.post(
        "/inspect",
        headers={"X-API-Key": " secret "},
        files={
            "sku_master_file": _sku_master_file(),
            "orders_file": _orders_file(),
        },
    )

    assert response.status_code == 200
    assert response.json()["matched"] == 1


def test_optimize_passes_max_orders_and_fast_mode(monkeypatch):
    monkeypatch.delenv("BOX_OPTIMIZER_API_KEY", raising=False)
    captured = {}

    def fake_optimize_workbook(*, sku_master_path, orders_path, output_path, config):
        captured["config"] = config
        with open(output_path, "wb") as output:
            output.write(b"PK fake workbook")
        return {"orders_processed": 1}

    monkeypatch.setattr("box_optimizer.api.optimize_workbook", fake_optimize_workbook)
    client = TestClient(app)

    response = client.post(
        "/optimize",
        files={
            "sku_master_file": _sku_master_file(),
            "orders_file": _orders_file_two_orders(),
        },
        data={"config_json": '{"max_orders": 1, "packing_mode": "fast"}'},
    )

    assert response.status_code == 200
    assert captured["config"]["max_orders"] == 1
    assert captured["config"]["packing_mode"] == "fast"




def test_optimize_accepts_balanced_mode_and_max_optimization_seconds(monkeypatch):
    monkeypatch.delenv("BOX_OPTIMIZER_API_KEY", raising=False)
    captured = {}

    def fake_optimize_workbook(*, sku_master_path, orders_path, output_path, config):
        captured["config"] = config
        with open(output_path, "wb") as output:
            output.write(b"PK fake workbook")
        return {"orders_processed": 1}

    monkeypatch.setattr("box_optimizer.api.optimize_workbook", fake_optimize_workbook)
    client = TestClient(app)

    response = client.post(
        "/optimize",
        files={
            "sku_master_file": _sku_master_file(),
            "orders_file": _orders_file_two_orders(),
        },
        data={"config_json": '{"packing_mode": "balanced", "max_optimization_seconds": 120}'},
    )

    assert response.status_code == 200
    assert captured["config"]["packing_mode"] == "balanced"
    assert captured["config"]["max_optimization_seconds"] == 120

def test_optimize_returns_xlsx_when_one_order_fails_packing(monkeypatch):
    monkeypatch.delenv("BOX_OPTIMIZER_API_KEY", raising=False)
    client = TestClient(app)

    response = client.post(
        "/optimize",
        files={
            "sku_master_file": _mixed_sku_master_file(),
            "orders_file": _mixed_orders_file(),
        },
        data={"config_json": '{"max_orders": 5, "packing_mode": "fast"}'},
    )

    assert response.status_code == 200
    assert response.headers["content-type"].startswith(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    assert response.content[:2] == b"PK"


def test_optimize_rejects_invalid_config_json(monkeypatch):
    monkeypatch.delenv("BOX_OPTIMIZER_API_KEY", raising=False)
    client = TestClient(app)

    response = client.post(
        "/optimize",
        files={
            "sku_master_file": _sku_master_file(),
            "orders_file": _orders_file(),
        },
        data={"config_json": "{not-json"},
    )

    assert response.status_code == 400


def _extract_job_id(html_text: str) -> str:
    match = re.search(r"Job ID:</strong> ([a-f0-9]{32})", html_text)
    assert match is not None
    return match.group(1)


def test_upload_page_uses_employee_friendly_labels(monkeypatch):
    monkeypatch.delenv("BOX_OPTIMIZER_UPLOAD_TOKEN", raising=False)
    client = TestClient(app)

    response = client.get("/upload")

    assert response.status_code == 200
    assert '<form action="/upload" method="post" enctype="multipart/form-data">' in response.text
    assert '<button type="submit">Run Optimization</button>' in response.text
    assert "Put your SKU file here" in response.text
    assert 'name="sku_master_file"' in response.text
    assert "Put your orders file here" in response.text
    assert 'name="orders_file"' in response.text
    assert 'name="config_json"' in response.text
    assert 'name="campaign_name"' in response.text
    assert 'name="packing_mode_choice"' in response.text
    assert 'name="ship_as_is_skus"' in response.text
    assert 'name="no_padding_skus"' in response.text
    assert 'name="wrap_around_skus"' in response.text
    assert 'name="compressible_skus"' in response.text
    assert 'id="config_generation_error"' in response.text
    assert "Optional packing instructions generated by GPT" in response.text
    assert "config_json</label>" not in response.text
    assert "recommended for Railway uploads" in response.text
    assert "Campaign setup" in response.text
    assert "Ship-as-is / do not touch packages" in response.text
    assert 'value="railway_fast" selected' in response.text
    assert "Vendor measured-fit mode" not in response.text
    assert 'name="vendor_flex_mode"' not in response.text
    assert "language variants must be listed as separate SKUs" in response.text
    assert "These items also receive no item-level padding by default" in response.text
    assert '&quot;packing_mode&quot;: &quot;fast&quot;' in response.text


def test_manual_vendor_box_fit_mode_remains_advanced_json_only(monkeypatch):
    monkeypatch.delenv("BOX_OPTIMIZER_UPLOAD_TOKEN", raising=False)
    client = TestClient(app)

    response = client.get('/upload?job_config={"vendor_box_fit_mode":"on"}')

    assert response.status_code == 200
    assert "Vendor measured-fit mode" not in response.text
    assert 'name="vendor_flex_mode"' not in response.text
    assert "vendor_box_fit_mode" in response.text


def test_upload_page_preserves_token_in_form_action_and_hidden_field(monkeypatch):
    monkeypatch.delenv("BOX_OPTIMIZER_UPLOAD_TOKEN", raising=False)
    client = TestClient(app)

    response = client.get("/upload?upload_token=local-upload-test")

    assert response.status_code == 200
    assert '<form action="/upload?upload_token=local-upload-test" method="post" enctype="multipart/form-data">' in response.text
    assert 'type="hidden" name="upload_token" value="local-upload-test"' in response.text
    assert '<button type="submit">Run Optimization</button>' in response.text

def test_upload_workflow_creates_job_status_and_download(monkeypatch, tmp_path):
    monkeypatch.delenv("BOX_OPTIMIZER_UPLOAD_TOKEN", raising=False)
    monkeypatch.setenv("BOX_OPTIMIZER_JOBS_DIR", str(tmp_path / "jobs"))
    client = TestClient(app)

    response = client.post(
        "/upload",
        files={
            "sku_master_file": ("campaign sku list.csv", _sku_master_file()[1], "text/csv"),
            "orders_file": ("daily orders export.csv", _orders_file()[1], "text/csv"),
        },
        data={"config_json": '{"packing_mode":"fast","output_granularity":"order_summary","preserve_region_sheets":false,"campaign":{"code":"Dark Horizon"}}'},
    )

    assert response.status_code == 200
    assert "Optimization Results" in response.text
    assert "Time to completed:</strong>" in response.text
    assert "Status:</strong> completed" in response.text
    assert "Orders processed:</strong> 1" in response.text
    job_id = _extract_job_id(response.text)

    status_response = client.get(f"/jobs/{job_id}")
    assert status_response.status_code == 200
    status = status_response.json()
    assert status["job_id"] == job_id
    assert status["status"] == "completed"
    assert status["summary"]["orders_processed"] == 1
    assert status["download_url"] == f"/jobs/{job_id}/download"

    download_response = client.get(f"/jobs/{job_id}/download")
    assert download_response.status_code == 200
    assert download_response.headers["content-type"].startswith(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    assert "Dark%20Horizon%20Shipping%20Plan.xlsx" in download_response.headers["content-disposition"]
    assert download_response.content[:2] == b"PK"


def test_upload_workflow_uses_default_friendly_config_when_blank(monkeypatch, tmp_path):
    monkeypatch.delenv("BOX_OPTIMIZER_UPLOAD_TOKEN", raising=False)
    monkeypatch.setenv("BOX_OPTIMIZER_JOBS_DIR", str(tmp_path / "jobs"))
    captured = {}

    def fake_optimize_workbook(*, sku_master_path, orders_path, output_path, config):
        captured["config"] = config
        with open(output_path, "wb") as output:
            output.write(b"PK fake workbook")
        return {
            "orders_processed": 1,
            "boxes_created": 1,
            "box_types": 1,
            "unmatched_skus": 0,
            "warnings": [],
            "warning_count": 0,
            "multi_box_order_count": 0,
            "rules_applied_count": 0,
        }

    monkeypatch.setattr("box_optimizer.api.optimize_workbook", fake_optimize_workbook)
    client = TestClient(app)

    response = client.post(
        "/upload",
        files={
            "sku_master_file": ("anything.xlsx", b"fake sku", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            "orders_file": ("anything else.xlsx", b"fake orders", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        },
        data={"config_json": "{}"},
    )

    assert response.status_code == 200
    assert captured["config"]["debug"] is True
    assert captured["config"]["packing_mode"] == "fast"
    assert captured["config"]["output_granularity"] == "order_summary"
    assert captured["config"]["preserve_region_sheets"] is False




def test_upload_workflow_generates_structured_rules_server_side(monkeypatch, tmp_path):
    monkeypatch.delenv("BOX_OPTIMIZER_UPLOAD_TOKEN", raising=False)
    monkeypatch.setenv("BOX_OPTIMIZER_JOBS_DIR", str(tmp_path / "jobs"))
    captured = {}

    def fake_optimize_workbook(*, sku_master_path, orders_path, output_path, config):
        captured["config"] = config
        with open(output_path, "wb") as output:
            output.write(b"PK fake workbook")
        return {
            "orders_processed": 1,
            "boxes_created": 1,
            "box_types": 1,
            "unmatched_skus": 0,
            "warnings": [],
            "warning_count": 0,
            "multi_box_order_count": 0,
            "rules_applied_count": 0,
        }

    monkeypatch.setattr("box_optimizer.api.optimize_workbook", fake_optimize_workbook)
    client = TestClient(app)

    response = client.post(
        "/upload",
        files={
            "sku_master_file": ("sku.xlsx", b"fake sku", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            "orders_file": ("orders.xlsx", b"fake orders", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        },
        data={
            "campaign_name": "Test Campaign",
            "campaign_code": "TC-1",
            "campaign_notes": "Internal note",
            "packing_mode_choice": "local_power_balanced_300",
            "ship_as_is_skus": "SHIP-001",
            "no_padding_skus": "DICE-001",
            "wrap_around_skus": "PLAYMAT-001",
            "wrapped_height_cm": "5",
            "compressible_skus": "PLUSH-001",
            "compressed_height_ratio": "0.55",
            "compressed_volume_ratio": "0.7",
        },
    )

    assert response.status_code == 200
    config = captured["config"]
    assert config["campaign"] == {"name": "Test Campaign", "code": "TC-1", "notes": "Internal note"}
    assert config["packing_mode"] == "balanced"
    assert config["max_optimization_seconds"] == 300
    assert config["sku_rules"]["SHIP-001"] == {
        "prepacked": True,
        "no_padding": True,
        "ships_alone": True,
        "can_mix_with_other_items": False,
        "box_type": "SHIP-001 shipping carton",
    }
    assert config["sku_rules"]["DICE-001"] == {"no_padding": True}
    assert config["sku_rules"]["PLAYMAT-001"] == {
        "wrap_around_largest_item": True,
        "wrapped_height_cm": 5.0,
        "no_padding": True,
    }
    assert config["sku_rules"]["PLUSH-001"] == {
        "compressible": True,
        "compressed_height_ratio": 0.55,
        "compressed_volume_ratio": 0.7,
    }


def test_upload_workflow_manual_config_overrides_structured_fields(monkeypatch, tmp_path):
    monkeypatch.delenv("BOX_OPTIMIZER_UPLOAD_TOKEN", raising=False)
    monkeypatch.setenv("BOX_OPTIMIZER_JOBS_DIR", str(tmp_path / "jobs"))
    captured = {}

    def fake_optimize_workbook(*, sku_master_path, orders_path, output_path, config):
        captured["config"] = config
        with open(output_path, "wb") as output:
            output.write(b"PK fake workbook")
        return {
            "orders_processed": 1,
            "boxes_created": 1,
            "box_types": 1,
            "unmatched_skus": 0,
            "warnings": [],
            "warning_count": 0,
            "multi_box_order_count": 0,
            "rules_applied_count": 0,
        }

    monkeypatch.setattr("box_optimizer.api.optimize_workbook", fake_optimize_workbook)
    client = TestClient(app)

    response = client.post(
        "/upload",
        files={
            "sku_master_file": ("sku.xlsx", b"fake sku", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            "orders_file": ("orders.xlsx", b"fake orders", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        },
        data={
            "packing_mode_choice": "railway_fast",
            "ship_as_is_skus": "SHIP-001",
            "config_json": '{"packing_mode":"balanced","max_optimization_seconds":60,"sku_rules":{"SHIP-001":{"box_type":"Manual carton"}}}',
        },
    )

    assert response.status_code == 200
    config = captured["config"]
    assert config["packing_mode"] == "balanced"
    assert config["max_optimization_seconds"] == 60
    assert config["sku_rules"]["SHIP-001"]["prepacked"] is True
    assert config["sku_rules"]["SHIP-001"]["box_type"] == "Manual carton"


def test_upload_workflow_returns_clear_error_when_files_missing(monkeypatch):
    monkeypatch.delenv("BOX_OPTIMIZER_UPLOAD_TOKEN", raising=False)
    client = TestClient(app)

    response = client.post("/upload", data={"packing_mode_choice": "railway_fast"})

    assert response.status_code == 400
    assert "both SKU and orders files are required" in response.text
    assert "FastAPI received POST:</strong> yes" in response.text

def test_power_upload_redirects_to_single_upload_page(monkeypatch):
    monkeypatch.delenv("BOX_OPTIMIZER_ENABLE_POWER_UPLOAD", raising=False)
    monkeypatch.delenv("BOX_OPTIMIZER_UPLOAD_TOKEN", raising=False)
    client = TestClient(app)

    response = client.get("/power-upload", follow_redirects=False)

    assert response.status_code == 303
    assert response.headers["location"] == "/upload?mode=local_power_balanced_300"


def test_power_upload_redirect_preserves_token_and_local_power_defaults(monkeypatch):
    monkeypatch.delenv("BOX_OPTIMIZER_ENABLE_POWER_UPLOAD", raising=False)
    monkeypatch.delenv("BOX_OPTIMIZER_UPLOAD_TOKEN", raising=False)
    client = TestClient(app)

    redirect = client.get("/power-upload?upload_token=local-upload-test", follow_redirects=False)
    assert redirect.status_code == 303
    assert redirect.headers["location"] == "/upload?mode=local_power_balanced_300&upload_token=local-upload-test"

    response = client.get(redirect.headers["location"])
    assert response.status_code == 200
    assert "Campaign setup" in response.text
    assert "Ship-as-is / do not touch packages" in response.text
    assert "Vendor measured-fit mode" not in response.text
    assert 'name="vendor_flex_mode"' not in response.text
    assert 'value="local_power_balanced_300" selected' in response.text
    assert 'action="/upload?upload_token=local-upload-test"' in response.text
    assert '&quot;packing_mode&quot;: &quot;balanced&quot;' in response.text
    assert '&quot;max_optimization_seconds&quot;: 300' in response.text


def test_power_upload_post_redirects_to_upload(monkeypatch):
    monkeypatch.delenv("BOX_OPTIMIZER_ENABLE_POWER_UPLOAD", raising=False)
    monkeypatch.delenv("BOX_OPTIMIZER_UPLOAD_TOKEN", raising=False)
    client = TestClient(app)

    response = client.post("/power-upload", follow_redirects=False)

    assert response.status_code == 307
    assert response.headers["location"] == "/upload"



def test_local_upload_launcher_opens_single_upload_page():
    script_path = Path(__file__).resolve().parents[1] / "start-local-upload.ps1"

    script_text = script_path.read_text(encoding="utf-8")

    assert 'BOX_OPTIMIZER_API_KEY="local-test-key"' in script_text
    assert 'BOX_OPTIMIZER_UPLOAD_TOKEN="local-upload-test"' in script_text
    assert 'BOX_OPTIMIZER_ENABLE_POWER_UPLOAD="true"' in script_text
    assert "http://127.0.0.1:8000/upload?upload_token=local-upload-test" in script_text
    assert "http://127.0.0.1:8000/power-upload" not in script_text
    assert "/health" in script_text


def test_upload_page_and_jobs_can_require_upload_token(monkeypatch, tmp_path):
    monkeypatch.setenv("BOX_OPTIMIZER_UPLOAD_TOKEN", "team-token")
    monkeypatch.setenv("BOX_OPTIMIZER_JOBS_DIR", str(tmp_path / "jobs"))
    client = TestClient(app)

    denied = client.get("/upload")
    assert denied.status_code == 403
    assert "Invalid or missing upload access token" in denied.text

    allowed = client.get("/upload?upload_token=team-token&job_config={%22packing_mode%22:%22fast%22}")
    assert allowed.status_code == 200
    assert "team-token" in allowed.text
    assert "packing_mode" in allowed.text

    upload_response = client.post(
        "/upload",
        files={
            "sku_master_file": ("sku list.csv", _sku_master_file()[1], "text/csv"),
            "orders_file": ("orders list.csv", _orders_file()[1], "text/csv"),
        },
        data={"upload_token": "team-token", "config_json": '{"packing_mode":"fast"}'},
    )
    assert upload_response.status_code == 200
    job_id = _extract_job_id(upload_response.text)

    assert client.get(f"/jobs/{job_id}").status_code == 403
    assert client.get(f"/jobs/{job_id}?upload_token=team-token").status_code == 200
    assert client.get(f"/jobs/{job_id}/download?upload_token=team-token").content[:2] == b"PK"
