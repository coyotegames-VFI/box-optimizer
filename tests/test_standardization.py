import pytest

from box_optimizer.models import Dimensions
from box_optimizer.packing.packer import Placement
from box_optimizer.standardization import (
    OptimizedOrderCarton,
    VENDOR_BOXES,
    VendorBox,
    _guarded_vendor_fit_candidates,
    _vendor_candidates,
    optimize_and_standardize_orders,
    sort_dimensions,
    standardize_optimized_cartons,
)


def test_sort_dimensions_largest_to_smallest():
    assert sort_dimensions(Dimensions(2, 5, 3)) == Dimensions(5, 3, 2)


def _optimized(
    order_id: str,
    combo: str,
    dimensions: Dimensions,
    chargeable_weight_kg: float = 1,
    allow_vendor_box_fit_tolerance: bool = False,
):
    return OptimizedOrderCarton(
        order_id=order_id,
        combination_key=combo,
        optimized_dimensions=dimensions,
        chargeable_weight_kg=chargeable_weight_kg,
        placements=[],
        allow_vendor_box_fit_tolerance=allow_vendor_box_fit_tolerance,
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


def test_vendor_box_menu_uses_smallest_safe_vendor_box_without_preference_override():
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
    assert assignments[0].selection_decision == "vendor_smallest_safe_fit"


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




def test_vendor_box_menu_uses_smallest_safe_full_list_box_even_below_non_preferred_threshold():
    assignments = standardize_optimized_cartons(
        [_optimized("order-1", "Combo x1", Dimensions(54.3, 36.6, 10.3), chargeable_weight_kg=5)],
        use_vendor_box_menu=True,
        billing_band_kg=1.0,
        non_preferred_box_min_units=100,
    )

    assert assignments[0].box_type == "Vendor Box 42"
    assert assignments[0].vendor_box_id == "42"
    assert assignments[0].selection_decision == "vendor_smallest_safe_fit"
    assert "smallest safe vendor fit" in assignments[0].box_standardization_note


def test_vendor_box_menu_uses_non_preferred_box_when_threshold_is_met():
    cartons = [
        _optimized(f"order-{index}", "Combo x1", Dimensions(54.3, 36.6, 10.3), chargeable_weight_kg=5)
        for index in range(100)
    ]

    assignments = standardize_optimized_cartons(
        cartons,
        use_vendor_box_menu=True,
        billing_band_kg=1.0,
        non_preferred_box_min_units=100,
    )

    assert {assignment.box_type for assignment in assignments} == {"Vendor Box 42"}
    assert {assignment.vendor_box_id for assignment in assignments} == {"42"}
    assert {assignment.selection_decision for assignment in assignments} == {"vendor_smallest_safe_fit"}
    assert all("smallest safe vendor fit" in assignment.box_standardization_note for assignment in assignments)



def test_vendor_box_menu_assigns_small_addon_carton_to_smallest_safe_box_not_vb36():
    assignments = standardize_optimized_cartons(
        [_optimized("order-1", "Dice Tray x1 | Sleeve Pack 1 x1 | Sleeve Pack 2 x1", Dimensions(26, 36, 14), chargeable_weight_kg=3)],
        use_vendor_box_menu=True,
        billing_band_kg=1.0,
    )

    assert assignments[0].vendor_box_id != "36"
    assert assignments[0].vendor_box_id == "16"
    assert assignments[0].box_type == "Vendor Box 16"
    assert assignments[0].selection_decision == "vendor_smallest_safe_fit"


def test_vendor_box_candidates_score_cut_down_height_before_billing_band_filter():
    carton = OptimizedOrderCarton(
        order_id="order-1",
        combination_key="Flat x1",
        optimized_dimensions=Dimensions(30, 26, 13),
        chargeable_weight_kg=3.4,
        placements=[
            Placement(
                canonical_sku="Flat",
                quantity=1,
                dimensions=Dimensions(30, 26, 8),
                origin=(0, 0, 0),
                weight_kg=1,
            )
        ],
    )

    candidates = _vendor_candidates(
        carton,
        VENDOR_BOXES,
        band_size_kg=1.0,
        same_band_only=True,
    )

    assert candidates
    billed, chargeable, volume_cm3, vendor_box, assigned_dimensions = candidates[0]
    assert vendor_box.vendor_id == "52"
    assert billed == 4
    assert chargeable == 3.4
    assert volume_cm3 == vendor_box.dimensions.length * vendor_box.dimensions.width * 10
    assert assigned_dimensions.height == 10


def test_vendor_box_candidates_reject_cut_down_that_cannot_fit_rigid_packed_bounds():
    carton = OptimizedOrderCarton(
        order_id="order-1",
        combination_key="Rigid Square x1",
        optimized_dimensions=Dimensions(34, 34, 12),
        chargeable_weight_kg=2.8,
        placements=[
            Placement(
                canonical_sku="Rigid Square",
                quantity=1,
                dimensions=Dimensions(31.4, 31.4, 9.8),
                origin=(0, 0, 0),
                weight_kg=1.95,
            )
        ],
    )

    candidates = _vendor_candidates(
        carton,
        VENDOR_BOXES,
        band_size_kg=1.0,
        same_band_only=True,
    )

    assert candidates
    assert all(candidate[3].vendor_id != "39" for candidate in candidates)
    assert candidates[0][3].vendor_id == "15"
    assert candidates[0][4] == Dimensions(35.4, 32.4, 12)


def test_rigid_packed_bounds_fit_vb15_before_cut_down():
    carton = OptimizedOrderCarton(
        order_id="order-1",
        combination_key="Rigid Square x1",
        optimized_dimensions=Dimensions(34, 34, 12),
        chargeable_weight_kg=2.8,
        placements=[
            Placement(
                canonical_sku="Rigid Square",
                quantity=1,
                dimensions=Dimensions(31.4, 31.4, 9.8),
                origin=(0, 0, 0),
                weight_kg=1.95,
            )
        ],
    )
    vb15 = tuple(box for box in VENDOR_BOXES if box.vendor_id == "15")

    candidates = _vendor_candidates(
        carton,
        vb15,
        band_size_kg=1.0,
        same_band_only=False,
    )

    assert len(candidates) == 1
    assert candidates[0][3].vendor_id == "15"
    assert candidates[0][4] == Dimensions(35.4, 32.4, 12)


def test_no_padding_rigid_item_still_cannot_fit_cut_down_box_with_too_small_side():
    carton = OptimizedOrderCarton(
        order_id="order-1",
        combination_key="Rigid Square x1",
        optimized_dimensions=Dimensions(32, 32, 10),
        chargeable_weight_kg=2.1,
        placements=[
            Placement(
                canonical_sku="Rigid Square",
                quantity=1,
                dimensions=Dimensions(29.4, 29.4, 7.8),
                origin=(0, 0, 0),
                weight_kg=1.95,
            )
        ],
    )
    vb39 = tuple(box for box in VENDOR_BOXES if box.vendor_id == "39")

    candidates = _vendor_candidates(
        carton,
        vb39,
        band_size_kg=1.0,
        same_band_only=False,
    )

    assert candidates == []


def test_final_vendor_assignment_dimensions_fit_rigid_placement_bounds():
    assignments = standardize_optimized_cartons(
        [
            OptimizedOrderCarton(
                order_id="order-1",
                combination_key="Rigid Square x1",
                optimized_dimensions=Dimensions(34, 34, 12),
                chargeable_weight_kg=2.8,
                placements=[
                    Placement(
                        canonical_sku="Rigid Square",
                        quantity=1,
                        dimensions=Dimensions(31.4, 31.4, 9.8),
                        origin=(0, 0, 0),
                        weight_kg=1.95,
                    )
                ],
            )
        ],
        use_vendor_box_menu=True,
        billing_band_kg=1.0,
    )

    assignment = assignments[0]
    assert assignment.vendor_box_id == "15"
    assert (
        assignment.assigned_length_cm,
        assignment.assigned_width_cm,
        assignment.assigned_height_cm,
    ) == (35.4, 32.4, 12)


def test_vendor_assignment_records_next_valid_backup_box_for_summary():
    assignments = standardize_optimized_cartons(
        [
            OptimizedOrderCarton(
                order_id="order-1",
                combination_key="Rigid Square x1",
                optimized_dimensions=Dimensions(34, 34, 12),
                chargeable_weight_kg=2.8,
                placements=[
                    Placement(
                        canonical_sku="Rigid Square",
                        quantity=1,
                        dimensions=Dimensions(31.4, 31.4, 9.8),
                        origin=(0, 0, 0),
                        weight_kg=1.95,
                    )
                ],
            )
        ],
        use_vendor_box_menu=True,
        billing_band_kg=1.0,
    )

    assignment = assignments[0]
    assert assignment.vendor_box_id == "15"
    assert assignment.backup_vendor_box_id
    assert assignment.backup_vendor_box_id != assignment.vendor_box_id
    assert assignment.backup_vendor_box_id != "39"
    assert (
        assignment.backup_assigned_length_cm,
        assignment.backup_assigned_width_cm,
        assignment.backup_assigned_height_cm,
    ) == (34, 34, 12)


def test_vendor_box_fit_tolerance_allows_small_real_world_carton_flex():
    assignments = standardize_optimized_cartons(
        [
            _optimized(
                "order-1",
                "Soft Combo x1",
                Dimensions(36.2, 34.8, 20),
                chargeable_weight_kg=5,
                allow_vendor_box_fit_tolerance=True,
            )
        ],
        use_vendor_box_menu=True,
        billing_band_kg=1.0,
        vendor_box_fit_tolerance_cm=1.5,
    )

    assert assignments[0].vendor_box_id == "21"
    assert assignments[0].assigned_length_cm == 36.2
    assert assignments[0].assigned_width_cm == 35
    assert assignments[0].assigned_height_cm == 21
    assert "fit tolerance" in assignments[0].box_standardization_note


def test_vendor_box_fit_tolerance_requires_carton_eligibility():
    assignments = standardize_optimized_cartons(
        [_optimized("order-1", "Rigid Combo x1", Dimensions(36.2, 34.8, 20), chargeable_weight_kg=5)],
        use_vendor_box_menu=True,
        billing_band_kg=1.0,
        vendor_box_fit_tolerance_cm=1.5,
    )

    assert assignments[0].vendor_box_id != "21"
    assert "fit tolerance" not in assignments[0].box_standardization_note


def test_vendor_box_fit_tolerance_is_capped_at_two_cm():
    assignments = standardize_optimized_cartons(
        [_optimized("order-1", "Too Big x1", Dimensions(37.2, 35, 20), chargeable_weight_kg=5)],
        use_vendor_box_menu=True,
        billing_band_kg=1.0,
        vendor_box_fit_tolerance_cm=3,
    )

    assert assignments[0].vendor_box_id != "21"


def test_vendor_box_fit_guardrail_rejects_flex_when_chargeable_increase_is_too_high():
    flex_box = VendorBox("FLEX", Dimensions(34, 34, 21))
    baseline_box = VendorBox("BASE", Dimensions(40, 27, 21))
    flex_candidate = (
        6.0,
        5.6,
        35.5 * 34 * 21,
        flex_box,
        Dimensions(35.5, 34, 21),
    )
    baseline_candidate = (
        6.0,
        4.6,
        40 * 27 * 21,
        baseline_box,
        Dimensions(40, 27, 21),
    )

    guarded = _guarded_vendor_fit_candidates([flex_candidate], [baseline_candidate], 0.25)

    assert guarded == []


def test_vendor_box_fit_guardrail_allows_flex_when_it_keeps_billed_weight_lower():
    flex_box = VendorBox("FLEX", Dimensions(34, 34, 21))
    baseline_box = VendorBox("BASE", Dimensions(40, 27, 21))
    flex_candidate = (
        5.0,
        5.6,
        35.5 * 34 * 21,
        flex_box,
        Dimensions(35.5, 34, 21),
    )
    baseline_candidate = (
        6.0,
        4.6,
        40 * 27 * 21,
        baseline_box,
        Dimensions(40, 27, 21),
    )

    guarded = _guarded_vendor_fit_candidates([flex_candidate], [baseline_candidate], 0.25)

    assert guarded == [flex_candidate]


def test_vendor_box_fit_guardrail_allows_preferred_flex_one_band_above_baseline():
    preferred_flex_box = VendorBox("34", Dimensions(48, 36, 27))
    baseline_box = VendorBox("39", Dimensions(40, 26.5, 44))
    preferred_candidate = (
        10.0,
        9.5904,
        48 * 37 * 27,
        preferred_flex_box,
        Dimensions(48, 37, 27),
    )
    baseline_candidate = (
        9.0,
        8.268,
        40 * 26.5 * 39,
        baseline_box,
        Dimensions(40, 26.5, 39),
    )

    guarded = _guarded_vendor_fit_candidates(
        [preferred_candidate],
        [baseline_candidate],
        1.0,
        band_size_kg=1.0,
    )

    assert guarded == [preferred_candidate]


def test_vendor_box_fit_guardrail_keeps_nonpreferred_flex_strict():
    nonpreferred_flex_box = VendorBox("NON", Dimensions(48, 36, 27))
    baseline_box = VendorBox("39", Dimensions(40, 26.5, 44))
    nonpreferred_candidate = (
        10.0,
        9.5904,
        48 * 37 * 27,
        nonpreferred_flex_box,
        Dimensions(48, 37, 27),
    )
    baseline_candidate = (
        9.0,
        8.268,
        40 * 26.5 * 39,
        baseline_box,
        Dimensions(40, 26.5, 39),
    )

    guarded = _guarded_vendor_fit_candidates(
        [nonpreferred_candidate],
        [baseline_candidate],
        1.0,
        band_size_kg=1.0,
    )

    assert guarded == []


def test_vendor_box_fit_guardrail_rejects_oversized_preferred_flex_candidate():
    preferred_flex_box = VendorBox("34", Dimensions(48, 36, 27))
    baseline_box = VendorBox("39", Dimensions(40, 26.5, 44))
    oversized_candidate = (
        11.0,
        10.1,
        48 * 37 * 30,
        preferred_flex_box,
        Dimensions(48, 37, 30),
    )
    baseline_candidate = (
        9.0,
        8.268,
        40 * 26.5 * 39,
        baseline_box,
        Dimensions(40, 26.5, 39),
    )

    guarded = _guarded_vendor_fit_candidates(
        [oversized_candidate],
        [baseline_candidate],
        1.0,
        band_size_kg=1.0,
    )

    assert guarded == []


def test_preferred_tolerance_vb34_style_candidate_beats_vb47_when_allowed():
    assignments = standardize_optimized_cartons(
        [
            OptimizedOrderCarton(
                order_id="order-1",
                combination_key="All In x1",
                optimized_dimensions=Dimensions(39, 37, 27),
                chargeable_weight_kg=7.7922,
                placements=[
                    Placement("Game Bundle", 1, Dimensions(31.4, 14.8, 31.4), (0, 0, 0), 3.25),
                    Placement("Insert", 1, Dimensions(37, 5.7, 37), (0, 14.8, 0), 1.0),
                    Placement("Playmat", 1, Dimensions(35, 4, 35), (0, 20.5, 0), 0.88),
                    Placement("Small Items", 1, Dimensions(3, 7, 3), (31.4, 0, 0), 0.495),
                ],
                allow_vendor_box_fit_tolerance=True,
            )
        ],
        use_vendor_box_menu=True,
        billing_band_kg=1.0,
        vendor_box_fit_tolerance_cm=1.5,
        vendor_box_fit_tolerance_guardrail=True,
        vendor_box_fit_tolerance_max_chargeable_increase_kg=1.0,
    )

    assignment = assignments[0]
    assert assignment.vendor_box_id == "34"
    assert (
        assignment.assigned_length_cm,
        assignment.assigned_width_cm,
        assignment.assigned_height_cm,
    ) == (48, 37, 27)
    assert "fit tolerance" in assignment.box_standardization_note


def test_preferred_flexible_allowance_does_not_apply_to_rigid_only_carton():
    assignments = standardize_optimized_cartons(
        [
            OptimizedOrderCarton(
                order_id="order-1",
                combination_key="All In x1",
                optimized_dimensions=Dimensions(39, 37, 27),
                chargeable_weight_kg=7.7922,
                placements=[
                    Placement("Game Bundle", 1, Dimensions(31.4, 14.8, 31.4), (0, 0, 0), 3.25),
                    Placement("Insert", 1, Dimensions(37, 5.7, 37), (0, 14.8, 0), 1.0),
                ],
                allow_vendor_box_fit_tolerance=False,
            )
        ],
        use_vendor_box_menu=True,
        billing_band_kg=1.0,
        vendor_box_fit_tolerance_cm=1.5,
        vendor_box_fit_tolerance_guardrail=True,
        vendor_box_fit_tolerance_max_chargeable_increase_kg=1.0,
    )

    assert assignments[0].vendor_box_id != "34"
    assert "fit tolerance" not in assignments[0].box_standardization_note


def test_vendor_box_menu_uses_vendor_box_even_when_400_custom_minimum_is_met():
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

    assert {assignment.box_type for assignment in assignments} == {"Vendor Box 41"}
    assert {assignment.vendor_box_id for assignment in assignments} == {"41"}
    assert {assignment.selection_decision for assignment in assignments} == {"vendor_smallest_safe_fit"}


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
