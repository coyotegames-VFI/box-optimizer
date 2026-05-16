import pytest

from box_optimizer.models import Dimensions
from box_optimizer.weights import (
    chargeable_weight_kg,
    dimensional_weight_kg,
    normalize_weight,
    packed_actual_weight_kg,
    total_weight,
)


def test_total_weight_sums_values():
    assert total_weight([1.5, 2.0, 3.25]) == 6.75


@pytest.mark.parametrize(
    ("weight", "unit", "expected_kg"),
    [
        (2, "kg", 2),
        (2500, "g", 2.5),
        (10, "lb", 4.5359237),
        (16, "oz", 0.45359237),
    ],
)
def test_normalize_weight_converts_supported_units_to_kg(weight, unit, expected_kg):
    result = normalize_weight(weight, unit)

    assert result.weight_kg == pytest.approx(expected_kg)
    assert result.weight_lb == pytest.approx(expected_kg * 2.20462262185)
    assert result.original_weight == weight
    assert result.original_unit == unit


def test_normalize_weight_rejects_unknown_units():
    with pytest.raises(ValueError, match="Unsupported weight unit"):
        normalize_weight(1, "stone")


def test_dimensional_weight_uses_5000_divisor():
    assert dimensional_weight_kg(Dimensions(50, 40, 30)) == 12


def test_packed_actual_weight_uses_default_multiplier():
    assert packed_actual_weight_kg(10) == pytest.approx(11.5)


def test_chargeable_weight_uses_dimensional_weight_when_larger():
    dimensions = Dimensions(50, 40, 30)

    assert chargeable_weight_kg(dimensions, item_weight_kg=5) == 12


def test_chargeable_weight_uses_packed_actual_weight_when_larger():
    dimensions = Dimensions(10, 10, 10)

    assert chargeable_weight_kg(dimensions, item_weight_kg=10) == pytest.approx(11.5)
