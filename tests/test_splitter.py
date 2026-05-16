from collections import Counter

from box_optimizer.models import Dimensions, PackedItem
from box_optimizer.packing.packer import MAX_CARTON_DIMENSIONS
from box_optimizer.packing.splitter import split_items, split_order_into_cartons


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
