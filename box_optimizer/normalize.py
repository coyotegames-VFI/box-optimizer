"""Input normalization helpers."""

from dataclasses import dataclass

from box_optimizer.models import Dimensions


_DIMENSION_TO_CM = {
    "cm": 1.0,
    "centimeter": 1.0,
    "centimeters": 1.0,
    "mm": 0.1,
    "millimeter": 0.1,
    "millimeters": 0.1,
    "m": 100.0,
    "meter": 100.0,
    "meters": 100.0,
    "in": 2.54,
    "inch": 2.54,
    "inches": 2.54,
    "ft": 30.48,
    "foot": 30.48,
    "feet": 30.48,
}


@dataclass(frozen=True)
class NormalizedDimensions:
    """Dimension values converted to centimeters with audit metadata."""

    dimensions: Dimensions
    original_dimensions: tuple[float, ...]
    original_unit: str


def normalize_sku(value: object) -> str:
    """Normalize SKU-like values into a stable uppercase string."""
    return str(value).strip().upper()


def normalize_dimensions(
    length: float,
    width: float,
    height: float | None = None,
    unit: str = "cm",
) -> NormalizedDimensions:
    """Convert dimensions to centimeters and sort them as L >= W >= H."""
    unit_key = unit.strip().lower()
    if unit_key not in _DIMENSION_TO_CM:
        raise ValueError(f"Unsupported dimension unit: {unit}")

    original = (length, width) if height is None else (length, width, height)
    factor = _DIMENSION_TO_CM[unit_key]
    converted = [value * factor for value in original]

    if height is None:
        converted.append(1.0)

    sorted_dimensions = sorted(converted, reverse=True)
    return NormalizedDimensions(
        dimensions=Dimensions(
            length=sorted_dimensions[0],
            width=sorted_dimensions[1],
            height=sorted_dimensions[2],
        ),
        original_dimensions=original,
        original_unit=unit,
    )
