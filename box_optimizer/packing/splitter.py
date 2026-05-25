"""Shipment splitting helpers."""

from dataclasses import dataclass
import math
from itertools import combinations

from box_optimizer.models import Dimensions, PackedItem
from box_optimizer.padding import add_final_exterior_padding
from box_optimizer.packing.packer import (
    MAX_CARTON_DIMENSIONS,
    OptimizedCartonResult,
    _expand_items,
    _sort_items,
    optimize_carton_dimensions,
    optimize_carton_dimensions_fast,
    pack_items,
)
from box_optimizer.standardization import (
    OptimizedOrderCarton,
    VENDOR_BOXES,
    _vendor_candidates,
)
from box_optimizer.weights import chargeable_weight_kg


@dataclass(frozen=True)
class SplitCarton:
    """One carton produced by an order split."""

    box_number: int
    result: OptimizedCartonResult
    box_type: str | None = None
    rule_applied: str = ""
    warning: str = ""
    dimensions_are_final: bool = False


@dataclass(frozen=True)
class SplitResult:
    """Result from splitting an order across practical cartons."""

    success: bool
    box_qty: int
    cartons: list[SplitCarton]
    unplaced_items: list[PackedItem]


def split_items(items: list[object], max_items: int) -> list[list[object]]:
    """Split items into simple chunks."""
    if max_items <= 0:
        raise ValueError("max_items must be greater than zero")
    return [items[index : index + max_items] for index in range(0, len(items), max_items)]


def _canonical_assignments(item_count: int, box_count: int) -> list[tuple[int, ...]]:
    if item_count == 0:
        return [()]
    if box_count <= 0 or box_count > item_count:
        return []

    assignments = []

    def build(current: list[int], max_label: int) -> None:
        if len(current) == item_count:
            if len(set(current)) == box_count:
                assignments.append(tuple(current))
            return

        remaining = item_count - len(current)
        missing_labels = box_count - len(set(current))
        if missing_labels > remaining:
            return

        for label in range(min(max_label + 2, box_count)):
            current.append(label)
            build(current, max(max_label, label))
            current.pop()

    build([0], 0)
    return assignments


def _items_for_assignment(
    items: list[PackedItem],
    assignment: tuple[int, ...],
    box_count: int,
) -> list[list[PackedItem]]:
    grouped = [[] for _ in range(box_count)]
    for item, box_index in zip(items, assignment, strict=True):
        grouped[box_index].append(item)
    return grouped


def _score_cartons(cartons: list[OptimizedCartonResult]) -> tuple[float, float]:
    return (
        sum(carton.chargeable_weight_kg or 0 for carton in cartons),
        sum(carton.volume_cm3 or 0 for carton in cartons),
    )


def _greedy_groups(items: list[PackedItem]) -> tuple[list[list[PackedItem]], list[PackedItem]]:
    groups: list[list[PackedItem]] = []
    best_failure: list[PackedItem] = []

    for item in items:
        placed = False
        best_group_index: int | None = None
        best_group_result: OptimizedCartonResult | None = None
        best_score: tuple[float, float] | None = None

        for index, group in enumerate(groups):
            candidate_group = [*group, item]
            candidate_result = optimize_carton_dimensions_fast(candidate_group)
            if not candidate_result.success:
                best_failure = candidate_result.unplaced_items
                continue
            score = (
                candidate_result.chargeable_weight_kg or 0,
                candidate_result.volume_cm3 or 0,
            )
            if best_score is None or score < best_score:
                best_score = score
                best_group_index = index
                best_group_result = candidate_result

        if best_group_index is not None and best_group_result is not None:
            groups[best_group_index].append(item)
            placed = True

        if not placed:
            single_result = optimize_carton_dimensions_fast([item])
            if not single_result.success:
                return [], single_result.unplaced_items or [item]
            groups.append([item])

    return groups, best_failure


def _split_groups(groups: list[list[PackedItem]], refine: bool = True) -> SplitResult:
    cartons = []
    best_failure: list[PackedItem] = []
    for index, group in enumerate(groups, start=1):
        fast_result = optimize_carton_dimensions_fast(group)
        if not fast_result.success:
            best_failure = fast_result.unplaced_items
            continue
        result = optimize_carton_dimensions(group) if refine else fast_result
        if not result.success:
            result = fast_result
        cartons.append(
            SplitCarton(
                box_number=index,
                result=result,
                warning="Split created because items could not fit existing capped cartons." if len(groups) > 1 else "",
            )
        )

    if len(cartons) == len(groups):
        return SplitResult(
            success=True,
            box_qty=len(cartons),
            cartons=cartons,
            unplaced_items=[],
        )
    return SplitResult(False, 0, [], best_failure)


def _fast_greedy_split(items: list[PackedItem]) -> SplitResult:
    """Deterministically split into as few capped cartons as practical."""
    groups, best_failure = _greedy_groups(items)
    if not groups:
        return SplitResult(False, 0, [], best_failure)
    return _split_groups(groups)


def _dimension_tuple(result: OptimizedCartonResult) -> tuple[float, float, float]:
    return (
        result.length_cm or 0,
        result.width_cm or 0,
        result.height_cm or 0,
    )


def _display_candidate_dimensions(dimensions: Dimensions) -> Dimensions:
    padded = add_final_exterior_padding(dimensions)
    return Dimensions(
        length=min(math.ceil(padded.length), MAX_CARTON_DIMENSIONS.length),
        width=min(math.ceil(padded.width), MAX_CARTON_DIMENSIONS.width),
        height=min(math.ceil(padded.height), MAX_CARTON_DIMENSIONS.height),
    )


def _vendor_score(result: OptimizedCartonResult) -> tuple[float, float, float, str]:
    if result.length_cm is None or result.width_cm is None or result.height_cm is None:
        return (float("inf"), float("inf"), float("inf"), "")
    dimensions = _display_candidate_dimensions(Dimensions(result.length_cm, result.width_cm, result.height_cm))
    actual_weight = sum(getattr(placement, "weight_kg", 0) for placement in result.placements)
    carton = OptimizedOrderCarton(
        order_id="candidate",
        combination_key="candidate",
        optimized_dimensions=dimensions,
        chargeable_weight_kg=chargeable_weight_kg(dimensions, actual_weight),
        placements=result.placements,
    )
    candidates = _vendor_candidates(carton, VENDOR_BOXES, band_size_kg=1.0, same_band_only=True)
    if not candidates:
        candidates = _vendor_candidates(carton, VENDOR_BOXES, band_size_kg=1.0, same_band_only=False)
    if candidates:
        billed, chargeable, box_volume, vendor_box, _assigned_dimensions = candidates[0]
        return (billed, chargeable, box_volume, vendor_box.vendor_id)
    return (
        result.chargeable_weight_kg or 0,
        result.chargeable_weight_kg or 0,
        result.volume_cm3 or 0,
        "CUSTOM",
    )


def _split_score(split_result: SplitResult) -> tuple[float, float, float, int, float]:
    if not split_result.success:
        return (float("inf"), float("inf"), float("inf"), 999999, float("inf"))
    vendor_scores = [_vendor_score(carton.result) for carton in split_result.cartons]
    billed = sum(score[0] for score in vendor_scores)
    chargeable = sum(score[1] for score in vendor_scores)
    vendor_volume = sum(score[2] for score in vendor_scores)
    local_box_types = len({score[3] for score in vendor_scores})
    return (split_result.box_qty, billed, chargeable, local_box_types, vendor_volume)


def _placement_dimensions(placements) -> Dimensions:
    return Dimensions(
        length=max(placement.origin[0] + placement.dimensions.length for placement in placements),
        width=max(placement.origin[1] + placement.dimensions.width for placement in placements),
        height=max(placement.origin[2] + placement.dimensions.height for placement in placements),
    )


def _result_from_vendor_shaped_pack(items: list[PackedItem], vendor_dimensions: Dimensions) -> OptimizedCartonResult | None:
    # Vendor boxes are outside dimensions; keep the existing +2 cm final exterior allowance
    # by packing into the usable interior candidate for this quick warehouse-style pass.
    interior = Dimensions(
        length=vendor_dimensions.length - 2,
        width=vendor_dimensions.width - 2,
        height=vendor_dimensions.height - 2,
    )
    if interior.length <= 0 or interior.width <= 0 or interior.height <= 0:
        return None
    packed = pack_items(items, interior)
    if not packed.success:
        return None
    dimensions = _placement_dimensions(packed.placements)
    total_weight_kg = sum(item.weight_kg for item in items)
    return OptimizedCartonResult(
        success=True,
        length_cm=dimensions.length,
        width_cm=dimensions.width,
        height_cm=dimensions.height,
        chargeable_weight_kg=chargeable_weight_kg(dimensions, total_weight_kg),
        volume_cm3=dimensions.length * dimensions.width * dimensions.height,
        placements=packed.placements,
        unplaced_items=[],
    )


def _vendor_shaped_fast_result(items: list[PackedItem], baseline: OptimizedCartonResult) -> OptimizedCartonResult:
    expanded_items = _sort_items(_expand_items(items))
    if len(expanded_items) < 2 or len(expanded_items) > 8:
        return baseline

    candidates = [SplitResult(True, 1, [SplitCarton(box_number=1, result=baseline)], [])]
    for vendor_box in sorted(
        VENDOR_BOXES,
        key=lambda box: (
            box.dimensions.length * box.dimensions.width * box.dimensions.height,
            box.dimensions.length,
            box.dimensions.width,
            box.dimensions.height,
            box.vendor_id,
        ),
    ):
        candidate = _result_from_vendor_shaped_pack(expanded_items, vendor_box.dimensions)
        if candidate is None:
            continue
        candidates.append(SplitResult(True, 1, [SplitCarton(box_number=1, result=candidate)], []))

    best = min(candidates, key=_split_score)
    return best.cartons[0].result


def _item_volume(item: PackedItem) -> float:
    dimensions = item.padded_dimensions
    return dimensions.length * dimensions.width * dimensions.height


def _item_longest_dimension(item: PackedItem) -> float:
    dimensions = item.padded_dimensions
    return max(dimensions.length, dimensions.width, dimensions.height)


def _balanced_orderings(items: list[PackedItem]) -> list[list[PackedItem]]:
    orderings = [
        list(items),
        sorted(items, key=lambda item: (_item_volume(item), _item_longest_dimension(item), item.weight_kg), reverse=True),
        sorted(items, key=lambda item: (_item_longest_dimension(item), _item_volume(item), item.weight_kg), reverse=True),
        sorted(items, key=lambda item: (item.padded_dimensions.height, -_item_volume(item), item.canonical_sku)),
        sorted(items, key=lambda item: (item.weight_kg, _item_volume(item)), reverse=True),
    ]
    unique = []
    seen = set()
    for ordering in orderings:
        signature = tuple((item.canonical_sku, item.padded_dimensions, item.weight_kg) for item in ordering)
        if signature not in seen:
            seen.add(signature)
            unique.append(ordering)
    return unique


def _try_recombine_groups(
    groups: list[list[PackedItem]],
    remaining_budget_seconds: float | None = None,
    min_remaining_seconds: float = 3,
) -> list[list[PackedItem]]:
    improved = [list(group) for group in groups]
    changed = True
    while changed and len(improved) > 1:
        changed = False
        current_result = _split_groups(improved)
        current_score = _split_score(current_result)
        best_groups = improved
        best_score = current_score

        if remaining_budget_seconds is not None and remaining_budget_seconds <= min_remaining_seconds:
            break
        for left, right in combinations(range(len(improved)), 2):
            if remaining_budget_seconds is not None and remaining_budget_seconds <= min_remaining_seconds:
                break
            merged = [*improved[left], *improved[right]]
            if not optimize_carton_dimensions_fast(merged).success:
                continue
            candidate_groups = [
                group
                for index, group in enumerate(improved)
                if index not in {left, right}
            ]
            candidate_groups.append(merged)
            candidate_result = _split_groups(candidate_groups)
            candidate_score = _split_score(candidate_result)
            if candidate_score <= best_score:
                best_score = candidate_score
                best_groups = candidate_groups

        if best_groups is not improved:
            improved = best_groups
            changed = True
    return improved


def _balanced_split(
    items: list[PackedItem],
    max_items_for_deep_search: int = 18,
    max_item_quantity_for_recombine: int = 10,
    min_remaining_seconds: float = 3,
    remaining_budget_seconds: float | None = None,
) -> SplitResult:
    candidates: list[SplitResult] = []
    best_failure: list[PackedItem] = []
    expanded_count = len(items)
    budget_is_low = remaining_budget_seconds is not None and remaining_budget_seconds <= min_remaining_seconds
    should_deep_search = expanded_count <= max_items_for_deep_search and not budget_is_low
    should_refine = should_deep_search and expanded_count <= 12
    should_recombine = should_deep_search and expanded_count <= max_item_quantity_for_recombine
    orderings = _balanced_orderings(items) if should_deep_search else _balanced_orderings(items)[:1]

    fast_baseline = _fast_greedy_split(items)
    if fast_baseline.success:
        candidates.append(fast_baseline)
        if not should_deep_search:
            return fast_baseline
    else:
        best_failure = fast_baseline.unplaced_items

    for ordering in orderings:
        groups, failure = _greedy_groups(ordering)
        if not groups:
            best_failure = failure or best_failure
            continue
        candidates.append(_split_groups(groups, refine=should_refine))
        if should_recombine:
            recombined_groups = _try_recombine_groups(
                groups,
                remaining_budget_seconds=remaining_budget_seconds,
                min_remaining_seconds=min_remaining_seconds,
            )
            if recombined_groups != groups:
                candidates.append(_split_groups(recombined_groups, refine=should_refine))

    if should_deep_search and expanded_count <= 8:
        for box_count in range(2, len(items) + 1):
            best_for_count: tuple[tuple[float, float], list[OptimizedCartonResult]] | None = None
            for assignment in _canonical_assignments(len(items), box_count):
                grouped_items = _items_for_assignment(items, assignment, box_count)
                carton_results = [optimize_carton_dimensions(group) for group in grouped_items]
                failures = [
                    item
                    for carton_result in carton_results
                    if not carton_result.success
                    for item in carton_result.unplaced_items
                ]
                if failures:
                    best_failure = failures
                    continue
                score = _score_cartons(carton_results)
                if best_for_count is None or score < best_for_count[0]:
                    best_for_count = (score, carton_results)
            if best_for_count is not None:
                candidates.append(
                    SplitResult(
                        True,
                        box_count,
                        [
                            SplitCarton(box_number=index + 1, result=carton)
                            for index, carton in enumerate(best_for_count[1])
                        ],
                        [],
                    )
                )
                break

    successful = [candidate for candidate in candidates if candidate.success]
    if not successful:
        return SplitResult(False, 0, [], best_failure)
    best = min(successful, key=_split_score)
    if fast_baseline.success and best.box_qty >= fast_baseline.box_qty:
        return fast_baseline
    return best


def split_order_into_cartons(
    items: list[PackedItem],
    packing_mode: str = "normal",
    force_simple_split: bool = False,
    balanced_max_items_for_deep_search: int = 18,
    balanced_max_item_quantity_for_recombine: int = 10,
    balanced_min_remaining_seconds: float = 3,
    remaining_budget_seconds: float | None = None,
) -> SplitResult:
    """Split an order into the fewest practical optimized cartons."""
    expanded_items = _sort_items(_expand_items(items))
    if not expanded_items:
        return SplitResult(success=True, box_qty=0, cartons=[], unplaced_items=[])

    optimizer = (
        optimize_carton_dimensions_fast
        if packing_mode in {"fast", "balanced"}
        else optimize_carton_dimensions
    )

    if force_simple_split:
        cartons = []
        unplaced = []
        for index, item in enumerate(expanded_items, start=1):
            carton = optimizer([item])
            if carton.success:
                cartons.append(
                    SplitCarton(
                        box_number=index,
                        result=carton,
                        box_type=item.box_type,
                        rule_applied=item.rule_applied,
                        warning=item.warning_note,
                        dimensions_are_final=False,
                    )
                )
            else:
                unplaced.extend(carton.unplaced_items or [item])
        if cartons and not unplaced:
            return SplitResult(
                success=True,
                box_qty=len(cartons),
                cartons=cartons,
                unplaced_items=[],
            )
        return SplitResult(
            success=False,
            box_qty=0,
            cartons=[],
            unplaced_items=unplaced,
        )

    single_box = optimizer(expanded_items)
    budget_is_low = remaining_budget_seconds is not None and remaining_budget_seconds <= balanced_min_remaining_seconds

    if packing_mode == "balanced" and single_box.success:
        single_box = _vendor_shaped_fast_result(expanded_items, single_box)
        fast_single = SplitResult(
            success=True,
            box_qty=1,
            cartons=[SplitCarton(box_number=1, result=single_box)],
            unplaced_items=[],
        )
        if budget_is_low or len(expanded_items) > balanced_max_items_for_deep_search:
            return fast_single
        refined_single = optimize_carton_dimensions(expanded_items)
        if refined_single.success:
            normal_single = SplitResult(
                success=True,
                box_qty=1,
                cartons=[SplitCarton(box_number=1, result=refined_single)],
                unplaced_items=[],
            )
            return min([fast_single, normal_single], key=_split_score)
        return fast_single
    if single_box.success:
        if packing_mode == "fast":
            single_box = _vendor_shaped_fast_result(expanded_items, single_box)
        return SplitResult(
            success=True,
            box_qty=1,
            cartons=[SplitCarton(box_number=1, result=single_box)],
            unplaced_items=[],
        )

    best_failure = single_box.unplaced_items
    if packing_mode == "fast":
        result = _fast_greedy_split(expanded_items)
        return result if result.success else SplitResult(False, 0, [], result.unplaced_items or best_failure)
    if packing_mode == "balanced":
        result = _balanced_split(
            expanded_items,
            max_items_for_deep_search=balanced_max_items_for_deep_search,
            max_item_quantity_for_recombine=balanced_max_item_quantity_for_recombine,
            min_remaining_seconds=balanced_min_remaining_seconds,
            remaining_budget_seconds=remaining_budget_seconds,
        )
        return result if result.success else SplitResult(False, 0, [], result.unplaced_items or best_failure)

    if len(expanded_items) > 8:
        result = _fast_greedy_split(expanded_items)
        return result if result.success else SplitResult(False, 0, [], result.unplaced_items or best_failure)

    for box_count in range(2, len(expanded_items) + 1):
        best: tuple[tuple[float, float], list[OptimizedCartonResult]] | None = None

        for assignment in _canonical_assignments(len(expanded_items), box_count):
            grouped_items = _items_for_assignment(expanded_items, assignment, box_count)
            carton_results = [optimizer(group) for group in grouped_items]
            failures = [
                item
                for carton_result in carton_results
                if not carton_result.success
                for item in carton_result.unplaced_items
            ]
            if failures:
                best_failure = failures
                continue

            score = _score_cartons(carton_results)
            if best is None or score < best[0]:
                best = (score, carton_results)

        if best is not None:
            cartons = [
                SplitCarton(box_number=index + 1, result=carton)
                for index, carton in enumerate(best[1])
            ]
            return SplitResult(
                success=True,
                box_qty=box_count,
                cartons=cartons,
                unplaced_items=[],
            )

    return SplitResult(
        success=False,
        box_qty=0,
        cartons=[],
        unplaced_items=best_failure,
    )
