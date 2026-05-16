import csv
from io import StringIO

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
