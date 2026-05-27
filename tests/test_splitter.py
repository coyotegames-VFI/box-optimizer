from collections import Counter

from box_optimizer.models import Dimensions, PackedItem
from box_optimizer.packing.packer import MAX_CARTON_DIMENSIONS, OptimizedCartonResult
from box_optimizer.packing.splitter import _display_candidate_dimensions, _vendor_score, split_items, split_order_into_cartons


def test_split_items_chunks_items():
    assert split_items([1, 2, 3, 4, 5], max_items=2) == [[1, 2], [3, 4], [5]]


def _item(
    sku: str,
    dimensions: Dimensions,
    weight_kg: float,
    quantity: int = 1,
) -> PackedItem:
    return PackedItem(
        canonical_sku=sku,
        quantity=quantity,
        unpadded_dimensions=dimensions,
        padded_dimensions=dimensions,
        weight_kg=weight_kg,
    )


def _packed_sku_counts(result):
    counts = Counter()
    for carton in result.cartons:
        for placement in carton.result.placements:
            counts[placement.canonical_sku] += placement.quantity
    return counts


def test_single_oversized_order_splits_into_multiple_boxes():
    result = split_order_into_cartons(
        [_item("large-panel", Dimensions(45, 35, 30), 2, quantity=3)]
    )

    assert result.success is True
    assert result.box_qty > 1


def test_split_order_box_qty_is_correct():
    result = split_order_into_cartons(
        [_item("large-panel", Dimensions(45, 35, 30), 2, quantity=3)]
    )

    assert result.success is True
    assert result.box_qty == 3
    assert len(result.cartons) == 3


def test_split_order_no_resulting_box_exceeds_cap():
    result = split_order_into_cartons(
        [_item("large-panel", Dimensions(45, 35, 30), 2, quantity=3)]
    )

    assert result.success is True
    for carton in result.cartons:
        assert carton.result.length_cm <= MAX_CARTON_DIMENSIONS.length
        assert carton.result.width_cm <= MAX_CARTON_DIMENSIONS.width
        assert carton.result.height_cm <= MAX_CARTON_DIMENSIONS.height


def test_split_order_total_packed_skus_equal_original_ordered_skus():
    ordered_items = [
        _item("large-panel", Dimensions(45, 35, 30), 2, quantity=3),
        _item("small-cube", Dimensions(5, 5, 5), 1, quantity=2),
    ]

    result = split_order_into_cartons(ordered_items)

    expected = Counter(
        {
            "large-panel": 3,
            "small-cube": 2,
        }
    )
    assert result.success is True
    assert _packed_sku_counts(result) == expected


def test_large_normal_mode_order_falls_back_without_combinatorial_search(monkeypatch):
    def fail_if_called(*args, **kwargs):
        raise AssertionError("large normal mode should not run combinatorial assignment search")

    monkeypatch.setattr("box_optimizer.packing.splitter._canonical_assignments", fail_if_called)

    result = split_order_into_cartons(
        [_item("large-panel", Dimensions(45, 35, 30), 2, quantity=9)],
        packing_mode="normal",
    )

    assert result.success is True
    assert result.box_qty == 9

def test_fast_mode_uses_simple_split_without_combinatorial_search(monkeypatch):
    def fail_if_called(*args, **kwargs):
        raise AssertionError("fast mode should not run combinatorial assignment search")

    monkeypatch.setattr("box_optimizer.packing.splitter._canonical_assignments", fail_if_called)

    result = split_order_into_cartons(
        [_item("large-panel", Dimensions(45, 35, 30), 2, quantity=3)],
        packing_mode="fast",
    )

    assert result.success is True
    assert result.box_qty == 3


def test_fast_split_refines_final_group_dimensions(monkeypatch):
    def result_for(items, dimensions, success=True):
        return OptimizedCartonResult(
            success=success,
            length_cm=dimensions.length if success else None,
            width_cm=dimensions.width if success else None,
            height_cm=dimensions.height if success else None,
            chargeable_weight_kg=1 if success else None,
            volume_cm3=dimensions.length * dimensions.width * dimensions.height if success else None,
            placements=[] if not success else [
                type(
                    "PlacementStub",
                    (),
                    {
                        "canonical_sku": item.canonical_sku,
                        "quantity": 1,
                        "dimensions": dimensions,
                        "origin": (0, 0, 0),
                    },
                )()
                for item in items
            ],
            unplaced_items=[] if success else list(items),
        )

    def fast_optimizer(items):
        skus = {item.canonical_sku for item in items}
        if skus == {"large", "small"}:
            return result_for(items, MAX_CARTON_DIMENSIONS, success=False)
        return result_for(items, MAX_CARTON_DIMENSIONS)

    def tight_optimizer(items):
        sku = items[0].canonical_sku
        dimensions = Dimensions(70, 35, 40) if sku == "large" else Dimensions(12, 10, 5)
        return result_for(items, dimensions)

    monkeypatch.setattr("box_optimizer.packing.splitter.optimize_carton_dimensions_fast", fast_optimizer)
    monkeypatch.setattr("box_optimizer.packing.splitter.optimize_carton_dimensions", tight_optimizer)

    result = split_order_into_cartons(
        [
            _item("large", Dimensions(70, 35, 40), 2),
            _item("small", Dimensions(12, 10, 5), 0.2),
        ],
        packing_mode="fast",
    )

    assert result.success is True
    assert result.box_qty == 2
    small_carton = next(
        carton for carton in result.cartons
        if any(placement.canonical_sku == "small" for placement in carton.result.placements)
    )
    assert small_carton.result.length_cm == 12
    assert small_carton.result.width_cm == 10
    assert small_carton.result.height_cm == 5


def test_balanced_mode_tries_alternate_orderings_without_changing_fast(monkeypatch):
    def result_for(items, success=True):
        dimensions = Dimensions(10 * len(items), 10, 5)
        return OptimizedCartonResult(
            success=success,
            length_cm=dimensions.length if success else None,
            width_cm=dimensions.width if success else None,
            height_cm=dimensions.height if success else None,
            chargeable_weight_kg=len(items) if success else None,
            volume_cm3=dimensions.length * dimensions.width * dimensions.height if success else None,
            placements=[] if not success else [
                type(
                    "PlacementStub",
                    (),
                    {
                        "canonical_sku": item.canonical_sku,
                        "quantity": 1,
                        "dimensions": dimensions,
                        "origin": (0, 0, 0),
                    },
                )()
                for item in items
            ],
            unplaced_items=[] if success else list(items),
        )

    def normal_optimizer(items):
        return result_for(items, success=len(items) < 3)

    def greedy_groups(items):
        by_sku = {item.canonical_sku: item for item in items}
        if items[0].canonical_sku == "A":
            return [[by_sku["A"]], [by_sku["B"]], [by_sku["C"]]], []
        return [[by_sku["B"], by_sku["C"]], [by_sku["A"]]], []

    monkeypatch.setattr("box_optimizer.packing.splitter.optimize_carton_dimensions", normal_optimizer)
    monkeypatch.setattr("box_optimizer.packing.splitter.optimize_carton_dimensions_fast", lambda items: result_for(items, success=len(items) < 3))
    monkeypatch.setattr("box_optimizer.packing.splitter._greedy_groups", greedy_groups)
    monkeypatch.setattr(
        "box_optimizer.packing.splitter._balanced_orderings",
        lambda items: [items, [items[1], items[2], items[0]]],
    )

    items = [
        _item("A", Dimensions(10, 10, 5), 1),
        _item("B", Dimensions(10, 10, 5), 1),
        _item("C", Dimensions(10, 10, 5), 1),
    ]

    fast_result = split_order_into_cartons(items, packing_mode="fast")
    balanced_result = split_order_into_cartons(items, packing_mode="balanced")

    assert fast_result.success is True
    assert fast_result.box_qty == 3
    assert balanced_result.success is True
    assert balanced_result.box_qty == 2


def test_balanced_high_complexity_combo_uses_fast_baseline_without_deep_search(monkeypatch):
    normal_calls = []

    def fast_optimizer(items):
        return OptimizedCartonResult(
            success=True,
            length_cm=20,
            width_cm=10,
            height_cm=5,
            chargeable_weight_kg=1,
            volume_cm3=1000,
            placements=[
                type(
                    "PlacementStub",
                    (),
                    {
                        "canonical_sku": item.canonical_sku,
                        "quantity": 1,
                        "dimensions": item.padded_dimensions,
                        "origin": (0, 0, 0),
                        "weight_kg": item.weight_kg,
                    },
                )()
                for item in items
            ],
            unplaced_items=[],
        )

    def normal_optimizer(items):
        normal_calls.append(len(items))
        raise AssertionError("high-complexity balanced run should not start deep search")

    monkeypatch.setattr("box_optimizer.packing.splitter.optimize_carton_dimensions_fast", fast_optimizer)
    monkeypatch.setattr("box_optimizer.packing.splitter.optimize_carton_dimensions", normal_optimizer)

    result = split_order_into_cartons(
        [_item(f"SKU-{index}", Dimensions(5, 5, 5), 0.1) for index in range(20)],
        packing_mode="balanced",
        balanced_max_items_for_deep_search=10,
    )

    assert result.success is True
    assert result.box_qty == 1
    assert normal_calls == []

def test_fast_mode_uses_vendor_shaped_cube_candidate_for_medium_rectangles():
    result = split_order_into_cartons(
        [
            _item("A", Dimensions(30, 25, 6.5), 0),
            _item("B", Dimensions(25, 25, 9.5), 0),
        ],
        packing_mode="fast",
    )

    assert result.success is True
    assert result.box_qty == 1
    carton = result.cartons[0].result
    assert (carton.length_cm, carton.width_cm, carton.height_cm) == (30, 25, 16.0)
    assert _display_candidate_dimensions(Dimensions(carton.length_cm, carton.width_cm, carton.height_cm)) == Dimensions(32, 27, 18)
    assert _vendor_score(carton)[3] == "52"


def test_balanced_single_box_baseline_uses_vendor_shaped_fast_candidate():
    result = split_order_into_cartons(
        [
            _item("A", Dimensions(30, 25, 6.5), 0),
            _item("B", Dimensions(25, 25, 9.5), 0),
        ],
        packing_mode="balanced",
    )

    assert result.success is True
    assert result.box_qty == 1
    assert _vendor_score(result.cartons[0].result)[3] == "52"


def test_fast_mode_preserves_long_box_pattern_for_large_anchor_items():
    result = split_order_into_cartons(
        [_item(f"large-{index}", Dimensions(37, 35, 12), 0) for index in range(3)],
        packing_mode="fast",
    )

    assert result.success is True
    assert result.box_qty == 1
    carton = result.cartons[0].result
    assert (carton.length_cm, carton.width_cm, carton.height_cm) == (74, 35, 24)
    assert _vendor_score(carton)[3] == "36"

