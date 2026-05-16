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


def _optimized(order_id: str, combo: str, dimensions: Dimensions):
    return OptimizedOrderCarton(
        order_id=order_id,
        combination_key=combo,
        optimized_dimensions=dimensions,
        chargeable_weight_kg=1,
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


def test_standardization_never_assigns_box_larger_than_cap():
    with pytest.raises(ValueError, match="exceed max carton cap"):
        standardize_optimized_cartons(
            [_optimized("order-1", "Huge x1", Dimensions(75, 37, 44))]
        )


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
