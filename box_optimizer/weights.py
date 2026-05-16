"""Weight calculation helpers."""

from dataclasses import dataclass

from box_optimizer.models import Dimensions


KG_TO_LB = 2.20462262185
_WEIGHT_TO_KG = {
    "kg": 1.0,
    "kilogram": 1.0,
    "kilograms": 1.0,
    "g": 0.001,
    "gram": 0.001,
    "grams": 0.001,
    "lb": 0.45359237,
    "lbs": 0.45359237,
    "pound": 0.45359237,
    "pounds": 0.45359237,
    "oz": 0.028349523125,
    "ounce": 0.028349523125,
    "ounces": 0.028349523125,
}


@dataclass(frozen=True)
class NormalizedWeight:
    """Weight converted to kilograms and pounds with audit metadata."""

    weight_kg: float
    weight_lb: float
    original_weight: float
    original_unit: str


def total_weight(weights: list[float]) -> float:
    """Return the sum of item weights."""
    return sum(weights)


def normalize_weight(weight: float, unit: str = "kg") -> NormalizedWeight:
    """Convert a weight value to kilograms and pounds."""
    unit_key = unit.strip().lower()
    if unit_key not in _WEIGHT_TO_KG:
        raise ValueError(f"Unsupported weight unit: {unit}")

    weight_kg = weight * _WEIGHT_TO_KG[unit_key]
    return NormalizedWeight(
        weight_kg=weight_kg,
        weight_lb=weight_kg * KG_TO_LB,
        original_weight=weight,
        original_unit=unit,
    )


def dimensional_weight_kg(dimensions: Dimensions) -> float:
    """Return dimensional weight in kg using L * W * H / 5000."""
    return dimensions.length * dimensions.width * dimensions.height / 5000


def packed_actual_weight_kg(item_weight_kg: float, multiplier: float = 1.15) -> float:
    """Return item weight after packing multiplier."""
    return item_weight_kg * multiplier


def chargeable_weight_kg(
    dimensions: Dimensions,
    item_weight_kg: float,
    multiplier: float = 1.15,
) -> float:
    """Return the greater of packed actual weight and dimensional weight."""
    return max(
        packed_actual_weight_kg(item_weight_kg, multiplier=multiplier),
        dimensional_weight_kg(dimensions),
    )
