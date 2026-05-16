"""Column mapping helpers."""

import re


_COLUMN_ALIASES = {
    "sku": {
        "sku",
        "itemsku",
        "productsku",
        "product",
        "productid",
        "item",
        "itemnumber",
    },
    "product_name": {
        "productname",
        "name",
        "itemname",
        "description",
        "productdescription",
    },
    "length": {"length", "lengthcm", "lengthin", "l"},
    "width": {"width", "widthcm", "widthin", "w"},
    "height": {"height", "heightcm", "heightin", "h", "depth", "depthcm", "depthin"},
    "weight": {
        "weight",
        "weightkg",
        "weightg",
        "weightlb",
        "weightlbs",
        "weightoz",
        "kg",
        "g",
        "lb",
        "lbs",
        "oz",
    },
    "order_id": {
        "orderid",
        "ordernumber",
        "orderno",
        "backernumber",
        "backer",
        "backerid",
    },
    "quantity": {"quantity", "qty", "count"},
    "region": {"region"},
    "country": {"country", "shipcountry", "shippingcountry"},
    "state_province": {
        "state",
        "province",
        "stateprovince",
        "shipstate",
        "shippingstate",
    },
}


def map_columns(row: dict, mapping: dict[str, str]) -> dict:
    """Map source row keys to normalized output keys."""
    return {target: row.get(source) for source, target in mapping.items()}


def normalize_column_name(name: object) -> str:
    """Normalize a source column name for inference."""
    return re.sub(r"[^a-z0-9]+", "", str(name or "").strip().lower())


def infer_columns(headers: list[str]) -> dict[str, str]:
    """Infer canonical field names from common source column names."""
    mapping = {}
    normalized_headers = [(header, normalize_column_name(header)) for header in headers]

    for canonical, aliases in _COLUMN_ALIASES.items():
        for header, normalized in normalized_headers:
            if normalized in aliases and canonical not in mapping:
                mapping[canonical] = header
                break

    return mapping


def infer_weight_unit(header: str | None, default: str = "kg") -> str:
    """Infer a weight unit from a column name."""
    normalized = normalize_column_name(header)
    if normalized.endswith("kg") or normalized == "kg":
        return "kg"
    if normalized.endswith("g") or normalized == "g":
        return "g"
    if normalized.endswith("lb") or normalized.endswith("lbs") or normalized in {"lb", "lbs"}:
        return "lb"
    if normalized.endswith("oz") or normalized == "oz":
        return "oz"
    return default


def infer_dimension_unit(header: str | None, default: str = "cm") -> str:
    """Infer a dimension unit from a column name."""
    normalized = normalize_column_name(header)
    if normalized.endswith("mm"):
        return "mm"
    if normalized.endswith("cm"):
        return "cm"
    if normalized.endswith("in"):
        return "in"
    if normalized.endswith("ft"):
        return "ft"
    return default
