import pytest

from box_optimizer.models import Dimensions
from box_optimizer.standardization import (
    OptimizedOrderCarton,
    optimize_and_standardize_orders,
    sort_dimensions,
    standardize_optimized_cartons,
)


def test_sort_dimensions_largest_to_smallest():
    assert sort_dimensions(Dimensions(2, 5, 3)) == Dimensions(5, 3, 2)


def _optimized(order_id: str, combo: str, dimensions: Dimensions, chargeable_weight_kg: float = 1):
    return OptimizedOrderCarton(
        order_id=order_id,
        combination_key=combo,
        optimized_dimensions=dimensions,
        chargeable_weight_kg=chargeable_weight_kg,
        placements=[],
    )


def test_standardization_rounds_up_never_down_within_tolerance():
    assignments = standardize_optimized_cartons(
        [
            _optimized("order-1", "Core Game x1", Dimensions(10, 8, 4)),
            _optimized("order-2", "Expansion A x1", Dimensions(11, 9, 5)),
        ]
    )

    assigned = {
        assignment.order_id: Dimensions(
            assignment.assigned_length_cm,
            assignment.assigned_width_cm,
            assignment.assigned_height_cm,
        )
        for assignment in assignments
    }

    assert assigned["order-1"] == Dimensions(11, 9, 5)
    assert assigned["order-2"] == Dimensions(11, 9, 5)
    for assignment in assignments:
        assert assignment.assigned_length_cm >= assignment.optimized_length_cm
        assert assignment.assigned_width_cm >= assignment.optimized_width_cm
        assert assignment.assigned_height_cm >= assignment.optimized_height_cm


def test_identical_exact_sku_combinations_receive_same_box_when_possible():
    assignments = standardize_optimized_cartons(
        [
            _optimized("order-1", "Core Game x1 | Dice Pack x1", Dimensions(10, 8, 4)),
            _optimized("order-2", "Core Game x1 | Dice Pack x1", Dimensions(10, 8, 4)),
        ]
    )

    assert assignments[0].box_type == assignments[1].box_type
    assert (
        assignments[0].assigned_length_cm,
        assignments[0].assigned_width_cm,
        assignments[0].assigned_height_cm,
    ) == (
        assignments[1].assigned_length_cm,
        assignments[1].assigned_width_cm,
        assignments[1].assigned_height_cm,
    )


def test_standardization_prefers_fewer_practical_box_types_within_tolerance():
    assignments = standardize_optimized_cartons(
        [
            _optimized("order-1", "A x1", Dimensions(10, 8, 4)),
            _optimized("order-2", "B x1", Dimensions(12, 10, 6)),
            _optimized("order-3", "C x1", Dimensions(30, 20, 10)),
        ],
        tolerance_cm=2,
    )

    box_types = {assignment.box_type for assignment in assignments}

    assert len(box_types) == 2
    assert assignments[0].box_type == assignments[1].box_type
    assert assignments[2].box_type != assignments[0].box_type


def test_standardization_does_not_merge_beyond_tolerance():
    assignments = standardize_optimized_cartons(
        [
            _optimized("order-1", "A x1", Dimensions(10, 8, 4)),
            _optimized("order-2", "B x1", Dimensions(13, 8, 4)),
        ],
        tolerance_cm=2,
    )

    assert assignments[0].box_type != assignments[1].box_type


def test_standardization_does_not_cross_next_dimensional_billing_band():
    assignments = standardize_optimized_cartons(
        [
            _optimized("order-1", "Small x1", Dimensions(50, 35, 14), chargeable_weight_kg=0.49),
            _optimized("order-2", "Shared x1", Dimensions(51, 35, 15), chargeable_weight_kg=0.5355),
        ],
        tolerance_cm=4,
    )

    assigned = {
        assignment.order_id: Dimensions(
            assignment.assigned_length_cm,
            assignment.assigned_width_cm,
            assignment.assigned_height_cm,
        )
        for assignment in assignments
    }
    assert assigned["order-1"] == Dimensions(50, 35, 14)
    assert assigned["order-2"] == Dimensions(51, 35, 15)
    assert assignments[0].box_type != assignments[1].box_type
    assert "no safe standardization candidate within billing band" in assignments[0].box_standardization_note


def test_standardization_can_round_up_within_same_dimensional_billing_band():
    assignments = standardize_optimized_cartons(
        [
            _optimized("order-1", "Small x1", Dimensions(10, 8, 4), chargeable_weight_kg=0.12),
            _optimized("order-2", "Shared x1", Dimensions(11, 9, 5), chargeable_weight_kg=0.12),
        ],
        tolerance_cm=4,
    )

    assert {
        (
            assignment.assigned_length_cm,
            assignment.assigned_width_cm,
            assignment.assigned_height_cm,
        )
        for assignment in assignments
    } == {(11, 9, 5)}
    assert len({assignment.box_type for assignment in assignments}) == 1
    assert any("without increasing billing band" in assignment.box_standardization_note for assignment in assignments)


def test_standardization_can_increase_dimensional_weight_inside_same_billing_band():
    assignments = standardize_optimized_cartons(
        [
            _optimized("order-1", "A x1", Dimensions(25.1, 20, 10), chargeable_weight_kg=1.01),
            _optimized("order-2", "B x1", Dimensions(29.1, 20, 10), chargeable_weight_kg=1.5),
        ],
        tolerance_cm=4,
    )

    first = next(assignment for assignment in assignments if assignment.order_id == "order-1")
    assert (
        first.assigned_length_cm,
        first.assigned_width_cm,
        first.assigned_height_cm,
    ) == (29.1, 20, 10)
    assert len({assignment.box_type for assignment in assignments}) == 1


def test_fewer_box_types_cannot_override_billing_band_priority():
    assignments = standardize_optimized_cartons(
        [
            _optimized("order-1", "Light x1", Dimensions(20, 20, 5), chargeable_weight_kg=0.2),
            _optimized("order-2", "HeavyDim x1", Dimensions(24, 24, 8), chargeable_weight_kg=0.9216),
        ],
        tolerance_cm=4,
    )

    first = next(assignment for assignment in assignments if assignment.order_id == "order-1")
    assert (
        first.assigned_length_cm,
        first.assigned_width_cm,
        first.assigned_height_cm,
    ) == (20, 20, 5)
    assert "no safe standardization candidate within billing band" in first.box_standardization_note


def test_standardization_never_assigns_box_larger_than_cap():
    with pytest.raises(ValueError, match="exceed max carton cap"):
        standardize_optimized_cartons(
            [_optimized("order-1", "Huge x1", Dimensions(75, 37, 44))]
        )


def test_vendor_box_menu_prefers_highlighted_box_with_one_kg_band():
    assignments = standardize_optimized_cartons(
        [_optimized("order-1", "Large x1", Dimensions(74, 36, 31), chargeable_weight_kg=17)],
        use_vendor_box_menu=True,
        billing_band_kg=1.0,
    )

    assert assignments[0].box_type == "Vendor Box 36"
    assert assignments[0].vendor_box_id == "36"
    assert assignments[0].assigned_length_cm == 74
    assert assignments[0].assigned_width_cm == 36
    assert assignments[0].assigned_height_cm == 42
    assert assignments[0].selection_decision == "preferred_fallback_higher_band"


def test_vendor_box_menu_can_use_full_list_beyond_old_cap():
    assignments = standardize_optimized_cartons(
        [_optimized("order-1", "Huge x1", Dimensions(74, 37, 38), chargeable_weight_kg=21)],
        use_vendor_box_menu=True,
        billing_band_kg=1.0,
    )

    assert assignments[0].box_type == "Vendor Box 41"
    assert assignments[0].vendor_box_id == "41"
    assert assignments[0].assigned_length_cm == 90
    assert assignments[0].assigned_width_cm == 45
    assert assignments[0].assigned_height_cm == 38


def test_vendor_box_menu_uses_custom_box_when_400_carton_minimum_is_met():
    cartons = [
        _optimized(f"order-{index}", "Same x1", Dimensions(74, 37, 38), chargeable_weight_kg=21)
        for index in range(400)
    ]

    assignments = standardize_optimized_cartons(
        cartons,
        use_vendor_box_menu=True,
        billing_band_kg=1.0,
        custom_box_min_units=400,
    )

    assert {assignment.box_type for assignment in assignments} == {"Custom Box 1"}
    assert {assignment.selection_decision for assignment in assignments} == {"custom_minimum_met"}
    assert {assignment.assigned_length_cm for assignment in assignments} == {74}


def test_optimize_and_standardize_orders_returns_required_output_fields():
    from box_optimizer.models import PackedItem

    assignments = optimize_and_standardize_orders(
        {
            "order-1": (
                "Cube x1",
                [
                    PackedItem(
                        canonical_sku="Cube",
                        quantity=1,
                        unpadded_dimensions=Dimensions(5, 5, 5),
                        padded_dimensions=Dimensions(5, 5, 5),
                        weight_kg=1,
                    )
                ],
            )
        }
    )

    assert assignments[0].box_type == "Box Type 1"
    assert assignments[0].optimized_length_cm == 5
    assert assignments[0].assigned_length_cm == 5
    assert assignments[0].box_standardization_note
