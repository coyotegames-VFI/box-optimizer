"""Deterministic heuristic 3D packer."""

from dataclasses import dataclass
from itertools import product

from box_optimizer.models import Dimensions, PackedItem
from box_optimizer.packing.geometry import (
    Coordinate,
    boxes_overlap,
    fits_within_boundaries,
    volume,
)
from box_optimizer.packing.rotations import unique_rotations
from box_optimizer.weights import chargeable_weight_kg


MAX_CARTON_DIMENSIONS = Dimensions(length=74, width=37, height=44)


@dataclass(frozen=True)
class Placement:
    """A single packed item placement."""

    canonical_sku: str
    quantity: int
    dimensions: Dimensions
    origin: Coordinate
    weight_kg: float


@dataclass(frozen=True)
class PackingResult:
    """Result from a heuristic packing attempt."""

    success: bool
    placements: list[Placement]
    unplaced_items: list[PackedItem]


@dataclass(frozen=True)
class OptimizedCartonResult:
    """Best carton dimensions found by testing practical candidates."""

    success: bool
    length_cm: float | None
    width_cm: float | None
    height_cm: float | None
    chargeable_weight_kg: float | None
    volume_cm3: float | None
    placements: list[Placement]
    unplaced_items: list[PackedItem]


def _longest_dimension(dimensions: Dimensions) -> float:
    return max(dimensions.length, dimensions.width, dimensions.height)


def _expand_items(items: list[PackedItem]) -> list[PackedItem]:
    expanded = []
    for item in items:
        expanded.extend(
            PackedItem(
                canonical_sku=item.canonical_sku,
                quantity=1,
                unpadded_dimensions=item.unpadded_dimensions,
                padded_dimensions=item.padded_dimensions,
                weight_kg=item.weight_kg,
            )
            for _ in range(item.quantity)
        )
    return expanded


def _axis_subset_sums(values_by_item: list[set[float]], limit: float) -> set[float]:
    sums = {0.0}
    for values in values_by_item:
        next_sums = set()
        for current_sum in sums:
            for value in values:
                candidate = current_sum + value
                if candidate <= limit:
                    next_sums.add(candidate)
        sums = next_sums
        if not sums:
            return set()
    return sums


def _candidate_axis_values(
    items: list[PackedItem],
    axis: str,
    limit: float,
) -> set[float]:
    values_by_item = []
    for item in items:
        values_by_item.append(
            {
                getattr(rotation, axis)
                for rotation in unique_rotations(item.padded_dimensions)
                if getattr(rotation, axis) <= limit
            }
        )

    if any(not values for values in values_by_item):
        return set()

    single_item_floor = max(min(values) for values in values_by_item)
    candidates = {single_item_floor, limit}
    candidates.update(_axis_subset_sums(values_by_item, limit))
    return {candidate for candidate in candidates if single_item_floor <= candidate <= limit}


def _generate_candidate_cartons(items: list[PackedItem]) -> list[Dimensions]:
    length_values = _candidate_axis_values(items, "length", MAX_CARTON_DIMENSIONS.length)
    width_values = _candidate_axis_values(items, "width", MAX_CARTON_DIMENSIONS.width)
    height_values = _candidate_axis_values(items, "height", MAX_CARTON_DIMENSIONS.height)

    candidates = [
        Dimensions(length=length, width=width, height=height)
        for length, width, height in product(length_values, width_values, height_values)
        if length <= MAX_CARTON_DIMENSIONS.length
        and width <= MAX_CARTON_DIMENSIONS.width
        and height <= MAX_CARTON_DIMENSIONS.height
    ]
    return sorted(candidates, key=lambda dimensions: (volume(dimensions), dimensions.length, dimensions.width, dimensions.height))


def _sort_items(items: list[PackedItem]) -> list[PackedItem]:
    return sorted(
        items,
        key=lambda item: (
            -volume(item.padded_dimensions),
            -_longest_dimension(item.padded_dimensions),
            -item.weight_kg,
            item.canonical_sku,
        ),
    )


def _candidate_points(placements: list[Placement]) -> list[Coordinate]:
    points = {(0.0, 0.0, 0.0)}
    for placement in placements:
        x, y, z = placement.origin
        dimensions = placement.dimensions
        points.add((x + dimensions.length, y, z))
        points.add((x, y + dimensions.width, z))
        points.add((x, y, z + dimensions.height))
    return sorted(points, key=lambda point: (point[2], point[1], point[0]))


def _overlaps_existing(
    origin: Coordinate,
    dimensions: Dimensions,
    placements: list[Placement],
) -> bool:
    return any(
        boxes_overlap(origin, dimensions, placement.origin, placement.dimensions)
        for placement in placements
    )


def _find_best_placement(
    item: PackedItem,
    carton_dimensions: Dimensions,
    placements: list[Placement],
) -> Placement | None:
    best: tuple[tuple[float, ...], Placement] | None = None

    for point in _candidate_points(placements):
        for rotation in unique_rotations(item.padded_dimensions):
            if not fits_within_boundaries(rotation, carton_dimensions, point):
                continue
            if _overlaps_existing(point, rotation, placements):
                continue

            x, y, z = point
            score = (
                z,
                y,
                x,
                z + rotation.height,
                y + rotation.width,
                x + rotation.length,
                rotation.length,
                rotation.width,
                rotation.height,
            )
            placement = Placement(
                canonical_sku=item.canonical_sku,
                quantity=1,
                dimensions=rotation,
                origin=point,
                weight_kg=item.weight_kg,
            )

            if best is None or score < best[0]:
                best = (score, placement)

    return None if best is None else best[1]


def pack_items(
    items: list[PackedItem],
    carton_dimensions: Dimensions,
) -> PackingResult:
    """Pack items into a carton using a deterministic extreme-points heuristic."""
    placements: list[Placement] = []
    unplaced_items: list[PackedItem] = []

    for item in _sort_items(_expand_items(items)):
        placement = _find_best_placement(item, carton_dimensions, placements)
        if placement is None:
            unplaced_items.append(item)
            continue
        placements.append(placement)

    return PackingResult(
        success=not unplaced_items,
        placements=placements,
        unplaced_items=unplaced_items,
    )


def _awkward_dimension_count(dimensions: Dimensions) -> int:
    awkward_thresholds = MAX_CARTON_DIMENSIONS
    return sum(
        [
            dimensions.length > awkward_thresholds.length * 0.85,
            dimensions.width > awkward_thresholds.width * 0.85,
            dimensions.height > awkward_thresholds.height * 0.85,
        ]
    )


def optimize_carton_dimensions(items: list[PackedItem]) -> OptimizedCartonResult:
    """Find the smallest practical carton that the real 3D packer can fill."""
    expanded_items = _expand_items(items)
    if not expanded_items:
        return OptimizedCartonResult(
            success=True,
            length_cm=0,
            width_cm=0,
            height_cm=0,
            chargeable_weight_kg=0,
            volume_cm3=0,
            placements=[],
            unplaced_items=[],
        )

    total_weight_kg = sum(item.weight_kg for item in expanded_items)
    best: tuple[tuple[float, ...], Dimensions, PackingResult] | None = None
    failed_unplaced = expanded_items

    for candidate in _generate_candidate_cartons(expanded_items):
        result = pack_items(expanded_items, candidate)
        if not result.success:
            failed_unplaced = result.unplaced_items
            continue

        candidate_chargeable_weight = chargeable_weight_kg(candidate, total_weight_kg)
        score = (
            candidate_chargeable_weight,
            volume(candidate),
            _awkward_dimension_count(candidate),
            candidate.height,
            candidate.width,
            candidate.length,
        )
        if best is None or score < best[0]:
            best = (score, candidate, result)

    if best is None:
        return OptimizedCartonResult(
            success=False,
            length_cm=None,
            width_cm=None,
            height_cm=None,
            chargeable_weight_kg=None,
            volume_cm3=None,
            placements=[],
            unplaced_items=failed_unplaced,
        )

    score, dimensions, result = best
    return OptimizedCartonResult(
        success=True,
        length_cm=dimensions.length,
        width_cm=dimensions.width,
        height_cm=dimensions.height,
        chargeable_weight_kg=score[0],
        volume_cm3=volume(dimensions),
        placements=result.placements,
        unplaced_items=[],
    )


def optimize_carton_dimensions_fast(items: list[PackedItem]) -> OptimizedCartonResult:
    """Pack into the capped carton first to avoid expensive carton searching."""
    expanded_items = _expand_items(items)
    if not expanded_items:
        return OptimizedCartonResult(
            success=True,
            length_cm=0,
            width_cm=0,
            height_cm=0,
            chargeable_weight_kg=0,
            volume_cm3=0,
            placements=[],
            unplaced_items=[],
        )

    result = pack_items(expanded_items, MAX_CARTON_DIMENSIONS)
    if not result.success:
        return OptimizedCartonResult(
            success=False,
            length_cm=None,
            width_cm=None,
            height_cm=None,
            chargeable_weight_kg=None,
            volume_cm3=None,
            placements=result.placements,
            unplaced_items=result.unplaced_items,
        )

    length = max(
        placement.origin[0] + placement.dimensions.length
        for placement in result.placements
    )
    width = max(
        placement.origin[1] + placement.dimensions.width
        for placement in result.placements
    )
    height = max(
        placement.origin[2] + placement.dimensions.height
        for placement in result.placements
    )
    dimensions = Dimensions(length=length, width=width, height=height)
    total_weight_kg = sum(item.weight_kg for item in expanded_items)

    return OptimizedCartonResult(
        success=True,
        length_cm=length,
        width_cm=width,
        height_cm=height,
        chargeable_weight_kg=chargeable_weight_kg(dimensions, total_weight_kg),
        volume_cm3=volume(dimensions),
        placements=result.placements,
        unplaced_items=[],
    )
