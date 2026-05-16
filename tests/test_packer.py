from box_optimizer.models import Dimensions, PackedItem
from box_optimizer.packing.geometry import boxes_overlap, fits_within_boundaries, volume
from box_optimizer.packing.packer import (
    MAX_CARTON_DIMENSIONS,
    optimize_carton_dimensions,
    pack_items,
)


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


def test_pack_items_returns_success_and_placement_detail():
    result = pack_items(
        items=[_item("sku-1", Dimensions(5, 5, 5), 1)],
        carton_dimensions=Dimensions(10, 10, 10),
    )

    assert result.success is True
    assert result.unplaced_items == []
    assert len(result.placements) == 1
    assert result.placements[0].canonical_sku == "sku-1"
    assert result.placements[0].origin == (0.0, 0.0, 0.0)


def test_pack_items_prevents_overlaps():
    result = pack_items(
        items=[
            _item("a", Dimensions(5, 5, 5), 1),
            _item("b", Dimensions(5, 5, 5), 1),
        ],
        carton_dimensions=Dimensions(10, 5, 5),
    )

    assert result.success is True
    first, second = result.placements
    assert not boxes_overlap(
        first.origin,
        first.dimensions,
        second.origin,
        second.dimensions,
    )


def test_pack_items_respects_carton_boundaries():
    carton = Dimensions(10, 10, 10)
    result = pack_items(
        items=[
            _item("a", Dimensions(6, 5, 5), 1),
            _item("b", Dimensions(4, 5, 5), 1),
        ],
        carton_dimensions=carton,
    )

    assert result.success is True
    assert all(
        fits_within_boundaries(placement.dimensions, carton, placement.origin)
        for placement in result.placements
    )


def test_pack_items_returns_failure_with_unplaced_detail():
    result = pack_items(
        items=[_item("too-big", Dimensions(11, 10, 10), 1)],
        carton_dimensions=Dimensions(10, 10, 10),
    )

    assert result.success is False
    assert result.placements == []
    assert result.unplaced_items[0].canonical_sku == "too-big"


def test_pack_items_places_larger_heavier_items_first():
    result = pack_items(
        items=[
            _item("small", Dimensions(2, 2, 2), 10),
            _item("large", Dimensions(5, 5, 5), 1),
            _item("heavy", Dimensions(5, 5, 5), 9),
        ],
        carton_dimensions=Dimensions(20, 20, 20),
    )

    assert [placement.canonical_sku for placement in result.placements[:2]] == [
        "heavy",
        "large",
    ]


def test_pack_items_expands_quantities_into_individual_placements():
    result = pack_items(
        items=[_item("small", Dimensions(2, 2, 2), 1, quantity=3)],
        carton_dimensions=Dimensions(6, 2, 2),
    )

    assert result.success is True
    assert len(result.placements) == 3
    assert [placement.quantity for placement in result.placements] == [1, 1, 1]


def test_pack_items_prefers_lower_z_positions():
    result = pack_items(
        items=[
            _item("base-a", Dimensions(5, 5, 5), 1),
            _item("base-b", Dimensions(5, 5, 5), 1),
        ],
        carton_dimensions=Dimensions(10, 5, 10),
    )

    assert [placement.origin for placement in result.placements] == [
        (0.0, 0.0, 0.0),
        (5.0, 0.0, 0.0),
    ]


def test_pack_items_uses_rotation_when_needed():
    result = pack_items(
        items=[_item("rotated", Dimensions(10, 4, 3), 1)],
        carton_dimensions=Dimensions(4, 10, 3),
    )

    assert result.success is True
    assert result.placements[0].dimensions == Dimensions(4, 10, 3)


def test_pack_items_is_not_volume_only():
    item_volume = volume(Dimensions(8, 8, 2))
    carton_volume = volume(Dimensions(7, 7, 3))

    assert item_volume < carton_volume

    result = pack_items(
        items=[_item("wrong-shape", Dimensions(8, 8, 2), 1)],
        carton_dimensions=Dimensions(7, 7, 3),
    )

    assert result.success is False


def test_requested_1_one_item_fits_in_one_box():
    result = pack_items(
        items=[_item("one", Dimensions(4, 4, 4), 1)],
        carton_dimensions=Dimensions(5, 5, 5),
    )

    assert result.success is True
    assert len(result.placements) == 1


def test_requested_2_two_identical_items_fit_side_by_side():
    result = pack_items(
        items=[_item("two-pack", Dimensions(5, 5, 5), 1, quantity=2)],
        carton_dimensions=Dimensions(10, 5, 5),
    )

    assert result.success is True
    assert [placement.origin for placement in result.placements] == [
        (0.0, 0.0, 0.0),
        (5.0, 0.0, 0.0),
    ]


def test_requested_3_item_only_fits_when_rotated():
    result = pack_items(
        items=[_item("rotate-me", Dimensions(9, 4, 3), 1)],
        carton_dimensions=Dimensions(4, 9, 3),
    )

    assert result.success is True
    assert result.placements[0].dimensions == Dimensions(4, 9, 3)


def test_requested_4_lower_total_volume_can_still_fail_due_to_dimensions():
    item_dimensions = Dimensions(8, 8, 2)
    carton_dimensions = Dimensions(7, 7, 3)

    assert volume(item_dimensions) < volume(carton_dimensions)

    result = pack_items(
        items=[_item("low-volume-bad-shape", item_dimensions, 1)],
        carton_dimensions=carton_dimensions,
    )

    assert result.success is False
    assert result.unplaced_items[0].canonical_sku == "low-volume-bad-shape"


def test_requested_5_no_placed_items_overlap():
    result = pack_items(
        items=[
            _item("a", Dimensions(5, 5, 5), 1),
            _item("b", Dimensions(5, 5, 5), 1),
            _item("c", Dimensions(5, 5, 5), 1),
            _item("d", Dimensions(5, 5, 5), 1),
        ],
        carton_dimensions=Dimensions(10, 10, 5),
    )

    assert result.success is True
    for index, first in enumerate(result.placements):
        for second in result.placements[index + 1 :]:
            assert not boxes_overlap(
                first.origin,
                first.dimensions,
                second.origin,
                second.dimensions,
            )


def test_optimize_carton_dimensions_finds_smallest_side_by_side_carton():
    result = optimize_carton_dimensions(
        [_item("cube", Dimensions(5, 5, 5), 1, quantity=2)]
    )

    assert result.success is True
    assert (result.length_cm, result.width_cm, result.height_cm) == (10, 5, 5)
    assert len(result.placements) == 2


def test_optimize_carton_dimensions_uses_actual_packer_not_volume_only():
    result = optimize_carton_dimensions(
        [_item("long-flat", Dimensions(8, 8, 2), 1)]
    )

    assert result.success is True
    assert result.length_cm <= MAX_CARTON_DIMENSIONS.length
    assert result.width_cm <= MAX_CARTON_DIMENSIONS.width
    assert result.height_cm <= MAX_CARTON_DIMENSIONS.height
    assert result.volume_cm3 == volume(Dimensions(8, 8, 2))


def test_optimize_carton_dimensions_returns_failure_when_no_candidate_can_pack():
    result = optimize_carton_dimensions(
        [_item("too-tall-any-way", Dimensions(75, 75, 75), 1)]
    )

    assert result.success is False
    assert result.placements == []
    assert result.unplaced_items[0].canonical_sku == "too-tall-any-way"


def test_optimize_carton_dimensions_respects_maximum_carton_limits():
    result = optimize_carton_dimensions(
        [_item("near-limit", Dimensions(74, 37, 44), 1)]
    )

    assert result.success is True
    assert result.length_cm <= 74
    assert result.width_cm <= 37
    assert result.height_cm <= 44


def test_optimize_carton_dimensions_returns_non_overlapping_placement_detail():
    result = optimize_carton_dimensions(
        [
            _item("a", Dimensions(5, 5, 5), 1),
            _item("b", Dimensions(5, 5, 5), 1),
            _item("c", Dimensions(5, 5, 5), 1),
        ]
    )

    assert result.success is True
    for index, first in enumerate(result.placements):
        for second in result.placements[index + 1 :]:
            assert not boxes_overlap(
                first.origin,
                first.dimensions,
                second.origin,
                second.dimensions,
            )
