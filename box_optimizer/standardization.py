"""Standardization helpers for dimensions and SKU data."""

from dataclasses import dataclass

from box_optimizer.models import Dimensions, PackedItem
from box_optimizer.packing.packer import (
    MAX_CARTON_DIMENSIONS,
    OptimizedCartonResult,
    optimize_carton_dimensions,
)


def sort_dimensions(dimensions: Dimensions) -> Dimensions:
    """Return dimensions sorted from largest to smallest."""
    length, width, height = sorted(
        [dimensions.length, dimensions.width, dimensions.height],
        reverse=True,
    )
    return Dimensions(length=length, width=width, height=height)


@dataclass(frozen=True)
class OptimizedOrderCarton:
    """An optimized carton result for one order or exact SKU combination."""

    order_id: str
    combination_key: str
    optimized_dimensions: Dimensions
    chargeable_weight_kg: float
    placements: list = None


@dataclass(frozen=True)
class StandardizedBoxAssignment:
    """Output row for campaign box standardization."""

    order_id: str
    combination_key: str
    box_type: str
    optimized_length_cm: float
    optimized_width_cm: float
    optimized_height_cm: float
    assigned_length_cm: float
    assigned_width_cm: float
    assigned_height_cm: float
    box_standardization_note: str
    placements: list = None


@dataclass
class _BoxType:
    name: str
    dimensions: Dimensions
    combination_keys: set[str]


def _within_cap(dimensions: Dimensions) -> bool:
    return (
        dimensions.length <= MAX_CARTON_DIMENSIONS.length
        and dimensions.width <= MAX_CARTON_DIMENSIONS.width
        and dimensions.height <= MAX_CARTON_DIMENSIONS.height
    )


def _can_round_up_to_box(
    optimized: Dimensions,
    assigned: Dimensions,
    tolerance_cm: float,
) -> bool:
    return (
        assigned.length >= optimized.length
        and assigned.width >= optimized.width
        and assigned.height >= optimized.height
        and assigned.length - optimized.length <= tolerance_cm
        and assigned.width - optimized.width <= tolerance_cm
        and assigned.height - optimized.height <= tolerance_cm
    )


def _merged_dimensions(left: Dimensions, right: Dimensions) -> Dimensions:
    return Dimensions(
        length=max(left.length, right.length),
        width=max(left.width, right.width),
        height=max(left.height, right.height),
    )


def _format_note(optimized: Dimensions, assigned: Dimensions) -> str:
    if optimized == assigned:
        return "Optimized dimensions used as assigned box."
    return "Rounded up to shared campaign box type."


def _box_type_name(index: int) -> str:
    return f"Box Type {index}"


def standardize_optimized_cartons(
    optimized_cartons: list[OptimizedOrderCarton],
    tolerance_cm: float = 2,
) -> list[StandardizedBoxAssignment]:
    """Group optimized cartons into a practical shared campaign box menu."""
    ordered = sorted(
        optimized_cartons,
        key=lambda carton: (
            carton.combination_key,
            carton.optimized_dimensions.length,
            carton.optimized_dimensions.width,
            carton.optimized_dimensions.height,
            carton.order_id,
        ),
    )
    box_types: list[_BoxType] = []
    assigned_by_combo: dict[str, _BoxType] = {}
    assignment_lookup: dict[str, _BoxType] = {}

    for carton in ordered:
        optimized_dimensions = carton.optimized_dimensions
        if not _within_cap(optimized_dimensions):
            raise ValueError(
                f"Optimized dimensions exceed max carton cap for {carton.order_id}"
            )

        box_type = assigned_by_combo.get(carton.combination_key)
        if box_type is not None and _can_round_up_to_box(
            optimized_dimensions,
            box_type.dimensions,
            tolerance_cm,
        ):
            assignment_lookup[carton.order_id] = box_type
            continue

        best: tuple[int, _BoxType, Dimensions] | None = None
        for candidate in box_types:
            merged = _merged_dimensions(candidate.dimensions, optimized_dimensions)
            if not _within_cap(merged):
                continue
            if not _can_round_up_to_box(optimized_dimensions, merged, tolerance_cm):
                continue
            if not all(
                _can_round_up_to_box(existing.optimized_dimensions, merged, tolerance_cm)
                for existing in ordered
                if assignment_lookup.get(existing.order_id) is candidate
            ):
                continue

            score = (
                (merged.length - candidate.dimensions.length)
                + (merged.width - candidate.dimensions.width)
                + (merged.height - candidate.dimensions.height)
            )
            if best is None or score < best[0]:
                best = (score, candidate, merged)

        if best is None:
            box_type = _BoxType(
                name=_box_type_name(len(box_types) + 1),
                dimensions=optimized_dimensions,
                combination_keys={carton.combination_key},
            )
            box_types.append(box_type)
        else:
            _, box_type, merged = best
            box_type.dimensions = merged
            box_type.combination_keys.add(carton.combination_key)

        assigned_by_combo[carton.combination_key] = box_type
        assignment_lookup[carton.order_id] = box_type

    return [
        StandardizedBoxAssignment(
            order_id=carton.order_id,
            combination_key=carton.combination_key,
            box_type=assignment_lookup[carton.order_id].name,
            optimized_length_cm=carton.optimized_dimensions.length,
            optimized_width_cm=carton.optimized_dimensions.width,
            optimized_height_cm=carton.optimized_dimensions.height,
            assigned_length_cm=assignment_lookup[carton.order_id].dimensions.length,
            assigned_width_cm=assignment_lookup[carton.order_id].dimensions.width,
            assigned_height_cm=assignment_lookup[carton.order_id].dimensions.height,
            box_standardization_note=_format_note(
                carton.optimized_dimensions,
                assignment_lookup[carton.order_id].dimensions,
            ),
            placements=carton.placements,
        )
        for carton in optimized_cartons
    ]


def optimize_and_standardize_orders(
    orders: dict[str, tuple[str, list[PackedItem]]],
    tolerance_cm: float = 2,
) -> list[StandardizedBoxAssignment]:
    """Optimize each order, then standardize the campaign box menu."""
    optimized_cartons = []
    for order_id, (combination_key, items) in orders.items():
        result: OptimizedCartonResult = optimize_carton_dimensions(items)
        if not result.success:
            raise ValueError(f"Could not optimize carton for order {order_id}")
        optimized_cartons.append(
            OptimizedOrderCarton(
                order_id=order_id,
                combination_key=combination_key,
                optimized_dimensions=Dimensions(
                    length=result.length_cm,
                    width=result.width_cm,
                    height=result.height_cm,
                ),
                chargeable_weight_kg=result.chargeable_weight_kg,
                placements=result.placements,
            )
        )
    return standardize_optimized_cartons(optimized_cartons, tolerance_cm=tolerance_cm)
