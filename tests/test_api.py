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


def test_optimize_returns_generated_excel_workbook(monkeypatch):
    monkeypatch.setenv("BOX_OPTIMIZER_API_KEY", "secret")
    client = TestClient(app)

    response = client.post(
        "/optimize",
        headers={"X-API-Key": "secret"},
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
