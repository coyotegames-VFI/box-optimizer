"""Shipment splitting helpers."""

from dataclasses import dataclass

from box_optimizer.models import PackedItem
from box_optimizer.packing.packer import (
    OptimizedCartonResult,
    _expand_items,
    _sort_items,
    optimize_carton_dimensions,
    optimize_carton_dimensions_fast,
)


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


def split_order_into_cartons(
    items: list[PackedItem],
    packing_mode: str = "normal",
    force_simple_split: bool = False,
) -> SplitResult:
    """Split an order into the fewest practical optimized cartons."""
    expanded_items = _sort_items(_expand_items(items))
    if not expanded_items:
        return SplitResult(success=True, box_qty=0, cartons=[], unplaced_items=[])

    optimizer = (
        optimize_carton_dimensions_fast
        if packing_mode == "fast"
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
    if single_box.success:
        return SplitResult(
            success=True,
            box_qty=1,
            cartons=[SplitCarton(box_number=1, result=single_box)],
            unplaced_items=[],
        )

    best_failure = single_box.unplaced_items
    if packing_mode == "fast":
        cartons = []
        unplaced = []
        for index, item in enumerate(expanded_items, start=1):
            carton = optimize_carton_dimensions_fast([item])
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
            unplaced_items=unplaced or best_failure,
        )

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
