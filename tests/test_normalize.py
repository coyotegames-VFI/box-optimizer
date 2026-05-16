import pytest

from box_optimizer.models import Dimensions
from box_optimizer.normalize import normalize_dimensions, normalize_sku


def test_normalize_sku_strips_and_uppercases():
    assert normalize_sku(" abc-123 ") == "ABC-123"


def test_normalize_dimensions_converts_inches_to_cm_and_sorts():
    result = normalize_dimensions(2, 10, 4, unit="in")

    assert result.dimensions == Dimensions(length=25.4, width=10.16, height=5.08)
    assert result.original_dimensions == (2, 10, 4)
    assert result.original_unit == "in"


def test_normalize_dimensions_uses_one_cm_height_for_flat_items():
    result = normalize_dimensions(30, 20, unit="cm")

    assert result.dimensions == Dimensions(length=30, width=20, height=1)
    assert result.original_dimensions == (30, 20)


def test_normalize_dimensions_flat_height_is_one_cm_after_unit_conversion():
    result = normalize_dimensions(10, 2, unit="in")

    assert result.dimensions == Dimensions(length=25.4, width=5.08, height=1)


def test_normalize_dimensions_rejects_unknown_units():
    with pytest.raises(ValueError, match="Unsupported dimension unit"):
        normalize_dimensions(1, 2, 3, unit="yard")
