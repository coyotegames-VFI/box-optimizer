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
    "length": {"length", "lengthcm", "lengthmm", "lengthin", "l", "lcm", "lmm", "lin"},
    "width": {"width", "widthcm", "widthmm", "widthin", "w", "wcm", "wmm", "win"},
    "height": {"height", "heightcm", "heightmm", "heightin", "h", "hcm", "hmm", "hin", "depth", "depthcm", "depthmm", "depthin"},
    "dimensions": {
        "dimensions",
        "dimensionsmm",
        "dimensionscm",
        "dimensionsin",
        "dimensionsft",
        "dimension",
        "dimensionmm",
        "dimensioncm",
        "dimensionin",
        "dimensionft",
        "size",
        "sizemm",
        "sizecm",
        "sizein",
        "sizeft",
        "productdimensions",
        "productdimensionsmm",
        "productdimensionscm",
        "productdimensionsin",
        "productdimensionsft",
        "dims",
        "dimsmm",
        "dimscm",
        "dimsin",
        "dimsft",
        "lxwxh",
        "lwh",
    },
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
        "vfi",
        "vfiid",
        "vfinumber",
    },
    "quantity": {"quantity", "qty", "count"},
    "region": {"region"},
    "country": {
        "country",
        "countrycode",
        "shipcountry",
        "shippingcountry",
        "shiptocountry",
        "shiptocountrycode",
        "addresscountry",
        "fullcountry",
    },
    "state_province": {
        "state",
        "province",
        "stateprovince",
        "shipstate",
        "shippingstate",
        "addressstate",
        "addressprovince",
    },
}

METADATA_COLUMN_ALIASES = {
    "orderid",
    "ordernumber",
    "orderno",
    "backernumber",
    "backer",
    "backerid",
    "backerkitid",
    "name",
    "firstname",
    "lastname",
    "fullname",
    "email",
    "phone",
    "company",
    "vfi",
    "vfiid",
    "vfinumber",
    "shippingname",
    "address",
    "add1",
    "add2",
    "address1",
    "address2",
    "addressline1",
    "addressline2",
    "city",
    "shippingcity",
    "state",
    "province",
    "stateprovince",
    "country",
    "postal",
    "postalcode",
    "shippingpostalcode",
    "zipcode",
    "zip",
    "shippingservice",
    "servicetype",
    "declaredname",
    "itemtotal",
    "unitprice",
    "noofpackages",
    "shippingmethod",
    "shipservice",
    "shipmethod",
    "shipmentmethod",
    "shippingaddress",
    "shippingaddress1",
    "shippingaddress2",
    "shiptocountry",
    "shiptocountrycode",
    "shipaddress",
    "shipaddress1",
    "shipaddress2",
    "addresstype",
    "shippingaddresstype",
    "region",
    "notes",
    "note",
    "comments",
    "comment",
    "pledge",
    "pledgename",
    "pledgelevel",
    "reward",
    "addons",
    "add ons",
    "status",
    "fulfillmentstatus",
    "fulfilmentstatus",
    "tracking",
    "trackingnumber",
    "total",
    "amount",
    "currency",
    "date",
    "created",
    "updated",
}

_METADATA_WORD_HINTS = {
    "address",
    "backer",
    "billing",
    "city",
    "country",
    "customer",
    "email",
    "fulfillment",
    "fulfilment",
    "name",
    "order",
    "phone",
    "postal",
    "province",
    "ship",
    "shipping",
    "state",
    "status",
    "street",
    "tracking",
    "zip",
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

    for canonical in ["length", "width", "height"]:
        if canonical in mapping:
            continue
        for header, normalized in normalized_headers:
            if normalized.startswith(canonical) and ("dimension" in normalized or "dims" in normalized):
                mapping[canonical] = header
                break

    if "weight" not in mapping:
        for header, normalized in normalized_headers:
            if "weight" in normalized and any(unit in normalized for unit in ["kg", "g", "lb", "lbs", "oz"]):
                mapping["weight"] = header
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


def is_metadata_column(header: str) -> bool:
    """Return whether a header is likely order/customer metadata."""
    normalized = normalize_column_name(header)
    if normalized in METADATA_COLUMN_ALIASES:
        return True
    if "notes" in normalized:
        return True
    words = {
        normalize_column_name(word)
        for word in re.split(r"[^A-Za-z0-9]+", str(header or ""))
        if word
    }
    return bool(words & _METADATA_WORD_HINTS)

