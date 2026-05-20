"""Deterministic heuristic 3D packer."""

from collections import Counter
from dataclasses import dataclass
from dataclasses import replace
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
            replace(item, quantity=1)
            for _ in range(item.quantity)
        )
    return expanded


def _rotations_for_item(item: PackedItem) -> list[Dimensions]:
    if item.allowed_orientations:
        rotations = list(item.allowed_orientations)
    elif not item.allow_rotation:
        rotations = [item.padded_dimensions]
    else:
        rotations = unique_rotations(item.padded_dimensions)

    if item.must_stay_flat:
        flat_height = item.padded_dimensions.height
        rotations = [
            rotation
            for rotation in rotations
            if rotation.height == flat_height
        ]
    return rotations


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
                for rotation in _rotations_for_item(item)
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
        for rotation in _rotations_for_item(item):
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



def _placement_bounds(placements: list[Placement]) -> tuple[float, float, float]:
    if not placements:
        return (0.0, 0.0, 0.0)
    return (
        max(placement.origin[0] + placement.dimensions.length for placement in placements),
        max(placement.origin[1] + placement.dimensions.width for placement in placements),
        max(placement.origin[2] + placement.dimensions.height for placement in placements),
    )


def _candidate_placements_for_item(
    item: PackedItem,
    carton_dimensions: Dimensions,
    placements: list[Placement],
) -> list[tuple[tuple[float, ...], Placement]]:
    candidates = []
    for point in _candidate_points(placements):
        for rotation in _rotations_for_item(item):
            if not fits_within_boundaries(rotation, carton_dimensions, point):
                continue
            if _overlaps_existing(point, rotation, placements):
                continue

            x, y, z = point
            candidate = Placement(
                canonical_sku=item.canonical_sku,
                quantity=1,
                dimensions=rotation,
                origin=point,
                weight_kg=item.weight_kg,
            )
            next_placements = [*placements, candidate]
            used_length, used_width, used_height = _placement_bounds(next_placements)
            score = (
                used_height,
                used_width,
                used_length,
                z,
                y,
                x,
                rotation.height,
                rotation.width,
                rotation.length,
            )
            candidates.append((score, candidate))
    return sorted(candidates, key=lambda candidate: candidate[0])


def _pack_items_with_backtracking(
    items: list[PackedItem],
    carton_dimensions: Dimensions,
    max_nodes: int = 20000,
) -> PackingResult:
    """Try a bounded small-order search when the greedy pass paints itself into a corner."""
    expanded = _sort_items(_expand_items(items))
    if len(expanded) > 8:
        return PackingResult(success=False, placements=[], unplaced_items=expanded)

    best_partial: list[Placement] = []
    best_success: list[Placement] | None = None
    nodes = 0

    def search(remaining: list[PackedItem], placements: list[Placement]) -> bool:
        nonlocal best_partial, best_success, nodes
        nodes += 1
        if nodes > max_nodes:
            return False
        if len(placements) > len(best_partial):
            best_partial = list(placements)
        if not remaining:
            best_success = list(placements)
            return True

        # Prefer the item with the fewest currently feasible placements. This avoids
        # spending the small search budget on easy pieces while a constrained flat
        # item gets boxed out by an earlier rotation.
        options_by_index = []
        for index, item in enumerate(remaining):
            options = _candidate_placements_for_item(item, carton_dimensions, placements)
            if not options:
                return False
            options_by_index.append((len(options), -volume(item.padded_dimensions), index, options))
        _count, _volume_score, index, options = min(options_by_index, key=lambda value: value[:3])
        item = remaining[index]
        next_remaining = [*remaining[:index], *remaining[index + 1 :]]

        for _score, placement in options:
            if search(next_remaining, [*placements, placement]):
                return True
        return False

    success = search(expanded, [])
    if success and best_success is not None:
        return PackingResult(success=True, placements=best_success, unplaced_items=[])

    placed_keys = Counter((p.canonical_sku, p.weight_kg, p.dimensions) for p in best_partial)
    unplaced = []
    for item in expanded:
        key_options = [
            (item.canonical_sku, item.weight_kg, rotation)
            for rotation in _rotations_for_item(item)
        ]
        matched_key = next((key for key in key_options if placed_keys.get(key, 0) > 0), None)
        if matched_key:
            placed_keys[matched_key] -= 1
        else:
            unplaced.append(item)
    return PackingResult(success=False, placements=best_partial, unplaced_items=unplaced or expanded[len(best_partial):])

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

    greedy_result = PackingResult(
        success=not unplaced_items,
        placements=placements,
        unplaced_items=unplaced_items,
    )
    if greedy_result.success:
        return greedy_result

    fallback_result = _pack_items_with_backtracking(items, carton_dimensions)
    if fallback_result.success or len(fallback_result.placements) > len(greedy_result.placements):
        return fallback_result
    return greedy_result


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
