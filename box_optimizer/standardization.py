"""Standardization helpers for dimensions and SKU data."""

import itertools
from collections import Counter
from dataclasses import dataclass

from box_optimizer.models import Dimensions, PackedItem
from box_optimizer.packing.packer import (
    MAX_CARTON_DIMENSIONS,
    OptimizedCartonResult,
    optimize_carton_dimensions,
)
from box_optimizer.weights import dimensional_weight_kg


def sort_dimensions(dimensions: Dimensions) -> Dimensions:
    """Return dimensions sorted from largest to smallest."""
    length, width, height = sorted(
        [dimensions.length, dimensions.width, dimensions.height],
        reverse=True,
    )
    return Dimensions(length=length, width=width, height=height)


@dataclass(frozen=True)
class OptimizedOrderCarton:
    """An optimized carton result for one order or exact SKU combination."""

    order_id: str
    combination_key: str
    optimized_dimensions: Dimensions
    chargeable_weight_kg: float
    placements: list = None
    allow_vendor_box_fit_tolerance: bool = False


@dataclass(frozen=True)
class StandardizedBoxAssignment:
    """Output row for campaign box standardization."""

    order_id: str
    combination_key: str
    box_type: str
    optimized_length_cm: float
    optimized_width_cm: float
    optimized_height_cm: float
    assigned_length_cm: float
    assigned_width_cm: float
    assigned_height_cm: float
    box_standardization_note: str
    placements: list = None
    vendor_box_id: str | None = None
    selection_decision: str = ""


@dataclass(frozen=True)
class VendorBox:
    """A vendor carton option in centimeters."""

    vendor_id: str
    dimensions: Dimensions


@dataclass
class _BoxType:
    name: str
    dimensions: Dimensions
    combination_keys: set[str]
    order_ids: set[str]


PREFERRED_VENDOR_BOX_IDS = frozenset(
    {"3", "6", "7", "12", "15", "18", "20", "33", "34", "35", "36", "47", "48", "53"}
)

VENDOR_BOXES = (
    VendorBox("1", Dimensions(15, 13.5, 15)),
    VendorBox("2", Dimensions(8, 8, 72.5)),
    VendorBox("3", Dimensions(26.5, 18.5, 10)),
    VendorBox("4", Dimensions(20, 20, 21)),
    VendorBox("5", Dimensions(23.5, 23.5, 9)),
    VendorBox("6", Dimensions(23.5, 23.5, 15)),
    VendorBox("7", Dimensions(35.5, 24.5, 12.5)),
    VendorBox("8", Dimensions(39.5, 18.5, 21.5)),
    VendorBox("9", Dimensions(34, 34, 16)),
    VendorBox("10", Dimensions(35, 27, 13)),
    VendorBox("11", Dimensions(35, 26.5, 13.5)),
    VendorBox("12", Dimensions(31, 31, 15.5)),
    VendorBox("13", Dimensions(35, 30.5, 13.5)),
    VendorBox("14", Dimensions(36.5, 20, 22)),
    VendorBox("15", Dimensions(35.4, 32.4, 13.8)),
    VendorBox("16", Dimensions(43.8, 30.7, 14.7)),
    VendorBox("17", Dimensions(35, 27, 21)),
    VendorBox("18", Dimensions(35, 26, 22)),
    VendorBox("18-1", Dimensions(35, 26, 18.5)),
    VendorBox("19", Dimensions(55, 20, 22)),
    VendorBox("20", Dimensions(34, 34, 21)),
    VendorBox("20-1", Dimensions(34, 34, 22)),
    VendorBox("21", Dimensions(35, 35, 21)),
    VendorBox("22", Dimensions(36.5, 36.5, 21)),
    VendorBox("23", Dimensions(37, 37, 21)),
    VendorBox("23-1", Dimensions(37, 38, 20.5)),
    VendorBox("24", Dimensions(34, 34, 28)),
    VendorBox("25", Dimensions(35, 35, 30)),
    VendorBox("26", Dimensions(31, 16, 16)),
    VendorBox("27", Dimensions(30, 23.5, 12)),
    VendorBox("28", Dimensions(32, 23, 11.5)),
    VendorBox("29", Dimensions(40, 16, 16)),
    VendorBox("30", Dimensions(50, 25, 9)),
    VendorBox("31", Dimensions(44.5, 34, 12.5)),
    VendorBox("32", Dimensions(41.5, 40.5, 24)),
    VendorBox("33", Dimensions(48, 31, 31)),
    VendorBox("34", Dimensions(48, 36, 27)),
    VendorBox("35", Dimensions(53, 38, 36)),
    VendorBox("36", Dimensions(74, 36, 42)),
    VendorBox("37", Dimensions(91, 44.5, 36)),
    VendorBox("38", Dimensions(38, 27.5, 12.5)),
    VendorBox("39", Dimensions(40, 26.5, 44)),
    VendorBox("40", Dimensions(28, 28, 9)),
    VendorBox("41", Dimensions(90, 45, 38)),
    VendorBox("42", Dimensions(55, 38, 23)),
    VendorBox("43", Dimensions(30, 23.5, 20)),
    VendorBox("44", Dimensions(39, 24, 13.5)),
    VendorBox("45", Dimensions(39, 26, 13.5)),
    VendorBox("46", Dimensions(45, 37, 14.5)),
    VendorBox("47", Dimensions(45, 39, 35)),
    VendorBox("48", Dimensions(45, 35, 39)),
    VendorBox("49", Dimensions(49, 34, 14)),
    VendorBox("50", Dimensions(42, 31.5, 12.5)),
    VendorBox("51", Dimensions(42, 31.5, 20)),
    VendorBox("52", Dimensions(31, 28, 16)),
    VendorBox("53", Dimensions(28, 28, 17)),
    VendorBox("54", Dimensions(31, 31, 17)),
    VendorBox("55", Dimensions(61, 16.8, 13)),
    VendorBox("56", Dimensions(28, 23, 18.5)),
    VendorBox("57", Dimensions(35, 26, 18)),
    VendorBox("58", Dimensions(41.5, 25.5, 11)),
    VendorBox("59", Dimensions(21.5, 21.5, 11)),
)


def _within_cap(dimensions: Dimensions) -> bool:
    return (
        dimensions.length <= MAX_CARTON_DIMENSIONS.length
        and dimensions.width <= MAX_CARTON_DIMENSIONS.width
        and dimensions.height <= MAX_CARTON_DIMENSIONS.height
    )


def _can_round_up_to_box(
    optimized: Dimensions,
    assigned: Dimensions,
    tolerance_cm: float,
) -> bool:
    return (
        assigned.length >= optimized.length
        and assigned.width >= optimized.width
        and assigned.height >= optimized.height
        and assigned.length - optimized.length <= tolerance_cm
        and assigned.width - optimized.width <= tolerance_cm
        and assigned.height - optimized.height <= tolerance_cm
    )


def _merged_dimensions(left: Dimensions, right: Dimensions) -> Dimensions:
    return Dimensions(
        length=max(left.length, right.length),
        width=max(left.width, right.width),
        height=max(left.height, right.height),
    )


def _format_note(optimized: Dimensions, assigned: Dimensions) -> str:
    if optimized == assigned:
        return "Optimized dimensions used as assigned box."
    return "Standardized upward to shared campaign box type without increasing billing band."


def _box_type_name(index: int) -> str:
    return f"Box Type {index}"


def _billing_band_kg(weight_kg: float, band_size_kg: float = 0.5) -> int:
    if weight_kg <= 0:
        return 0
    return int((weight_kg + band_size_kg - 1e-9) // band_size_kg)


def _is_billing_safe(
    carton: OptimizedOrderCarton,
    assigned_dimensions: Dimensions,
) -> bool:
    optimized_dim_weight = dimensional_weight_kg(carton.optimized_dimensions)
    assigned_dim_weight = dimensional_weight_kg(assigned_dimensions)
    return _billing_band_kg(assigned_dim_weight) <= _billing_band_kg(optimized_dim_weight)


def _billed_weight_kg(weight_kg: float, band_size_kg: float) -> float:
    return _billing_band_kg(weight_kg, band_size_kg) * band_size_kg


def _dimensions_key(dimensions: Dimensions) -> tuple[float, float, float]:
    return (dimensions.length, dimensions.width, dimensions.height)


def _box_volume(dimensions: Dimensions) -> float:
    return dimensions.length * dimensions.width * dimensions.height


def _ceil_cm(value: float) -> int:
    return int(value) if float(value).is_integer() else int(value) + 1


def _placement_top_height(placements: list | None) -> float:
    if not placements:
        return 0.0
    return max(
        (
            (placement.origin[2] if placement.origin else 0)
            + placement.dimensions.height
            for placement in placements
        ),
        default=0.0,
    )


def _placement_bounds(placements: list | None) -> Dimensions | None:
    if not placements:
        return None
    return Dimensions(
        length=max((placement.origin[0] if placement.origin else 0) + placement.dimensions.length for placement in placements),
        width=max((placement.origin[1] if placement.origin else 0) + placement.dimensions.width for placement in placements),
        height=max((placement.origin[2] if placement.origin else 0) + placement.dimensions.height for placement in placements),
    )


def _required_fit_dimensions(carton: OptimizedOrderCarton) -> Dimensions:
    return _placement_bounds(carton.placements) or carton.optimized_dimensions


def _vendor_score_dimensions(
    carton: OptimizedOrderCarton,
    vendor_box: VendorBox,
    tolerance_cm: float = 0,
) -> Dimensions:
    packed_height = _placement_top_height(carton.placements)
    if not packed_height:
        score_dimensions = vendor_box.dimensions
    else:
        cut_height = min(vendor_box.dimensions.height, _ceil_cm(packed_height + 2))
        score_dimensions = Dimensions(
            length=vendor_box.dimensions.length,
            width=vendor_box.dimensions.width,
            height=cut_height,
        )
    if tolerance_cm <= 0:
        return score_dimensions
    return Dimensions(
        length=max(score_dimensions.length, carton.optimized_dimensions.length),
        width=max(score_dimensions.width, carton.optimized_dimensions.width),
        height=max(score_dimensions.height, carton.optimized_dimensions.height),
    )


def _uses_vendor_fit_tolerance(assigned_dimensions: Dimensions, vendor_dimensions: Dimensions) -> bool:
    return (
        assigned_dimensions.length > vendor_dimensions.length
        or assigned_dimensions.width > vendor_dimensions.width
        or assigned_dimensions.height > vendor_dimensions.height
    )


def _vendor_fit_tolerance_note(assigned_dimensions: Dimensions, vendor_box: VendorBox) -> str:
    if not _uses_vendor_fit_tolerance(assigned_dimensions, vendor_box.dimensions):
        return ""
    overages = []
    for label, assigned, nominal in [
        ("L", assigned_dimensions.length, vendor_box.dimensions.length),
        ("W", assigned_dimensions.width, vendor_box.dimensions.width),
        ("H", assigned_dimensions.height, vendor_box.dimensions.height),
    ]:
        if assigned > nominal:
            overages.append(f"{label}+{assigned - nominal:.1f} cm")
    return " Uses measured fit tolerance for real-world carton flex: " + ", ".join(overages) + "."


def _vendor_height_cutdown_note(assigned_dimensions: Dimensions, vendor_box: VendorBox) -> str:
    if assigned_dimensions.height >= vendor_box.dimensions.height:
        return ""
    return f" Vendor box height cut down to {int(assigned_dimensions.height)} cm."


def _fits_with_rotation(item: Dimensions, carton: Dimensions, tolerance_cm: float = 0) -> bool:
    return any(
        rotation.length <= carton.length + tolerance_cm
        and rotation.width <= carton.width + tolerance_cm
        and rotation.height <= carton.height + tolerance_cm
        for rotation in {
            Dimensions(*values)
            for values in itertools.permutations(
                (item.length, item.width, item.height),
                3,
            )
        }
    )


def _fits_effective_vendor_dimensions(
    item: Dimensions,
    vendor_dimensions: Dimensions,
    assigned_dimensions: Dimensions,
    tolerance_cm: float,
) -> bool:
    if not _fits_with_rotation(item, vendor_dimensions, tolerance_cm=tolerance_cm):
        return False
    return (
        assigned_dimensions.length <= vendor_dimensions.length + tolerance_cm + 1e-9
        and assigned_dimensions.width <= vendor_dimensions.width + tolerance_cm + 1e-9
        and assigned_dimensions.height <= vendor_dimensions.height + tolerance_cm + 1e-9
    )


def _fits_final_assigned_dimensions(
    required_dimensions: Dimensions,
    assigned_dimensions: Dimensions,
    tolerance_cm: float,
) -> bool:
    return _fits_with_rotation(
        required_dimensions,
        assigned_dimensions,
        tolerance_cm=tolerance_cm,
    )


def _vendor_candidates(
    carton: OptimizedOrderCarton,
    vendor_boxes: tuple[VendorBox, ...],
    band_size_kg: float,
    same_band_only: bool,
    vendor_box_fit_tolerance_cm: float = 0,
) -> list[tuple[float, float, float, VendorBox, Dimensions]]:
    optimized_billed = _billed_weight_kg(carton.chargeable_weight_kg, band_size_kg)
    tolerance_cm = max(0.0, min(float(vendor_box_fit_tolerance_cm or 0), 2.0))
    candidates = []
    required_dimensions = _required_fit_dimensions(carton)
    for vendor_box in vendor_boxes:
        if not _fits_with_rotation(required_dimensions, vendor_box.dimensions, tolerance_cm=tolerance_cm):
            continue
        score_dimensions = _vendor_score_dimensions(carton, vendor_box, tolerance_cm=tolerance_cm)
        if not _fits_effective_vendor_dimensions(
            required_dimensions,
            vendor_box.dimensions,
            score_dimensions,
            tolerance_cm,
        ):
            continue
        if not _fits_final_assigned_dimensions(
            required_dimensions,
            score_dimensions,
            tolerance_cm,
        ):
            continue
        vendor_dimensional_weight = dimensional_weight_kg(score_dimensions)
        vendor_chargeable_weight = max(carton.chargeable_weight_kg, vendor_dimensional_weight)
        vendor_billed_weight = _billed_weight_kg(vendor_chargeable_weight, band_size_kg)
        if same_band_only and vendor_billed_weight > optimized_billed:
            continue
        candidates.append(
            (
                vendor_billed_weight,
                vendor_chargeable_weight,
                _box_volume(score_dimensions),
                vendor_box,
                score_dimensions,
            )
        )
    return sorted(candidates, key=lambda candidate: candidate[:3])


def _guarded_vendor_fit_candidates(
    candidates: list[tuple[float, float, float, VendorBox, Dimensions]],
    baseline_candidates: list[tuple[float, float, float, VendorBox, Dimensions]],
    max_chargeable_increase_kg: float,
) -> list[tuple[float, float, float, VendorBox, Dimensions]]:
    if not candidates or not baseline_candidates:
        return candidates
    baseline_billed, baseline_chargeable, _baseline_volume, _baseline_box, _baseline_dimensions = baseline_candidates[0]
    guarded = []
    for candidate in candidates:
        billed, chargeable, _volume, vendor_box, assigned_dimensions = candidate
        if _uses_vendor_fit_tolerance(assigned_dimensions, vendor_box.dimensions):
            if billed > baseline_billed + 1e-9:
                continue
            if (
                billed >= baseline_billed - 1e-9
                and chargeable > baseline_chargeable + max_chargeable_increase_kg + 1e-9
            ):
                continue
        guarded.append(candidate)
    return guarded


def _vendor_assignment(
    carton: OptimizedOrderCarton,
    box_type: str,
    vendor_box: VendorBox,
    note: str,
    decision: str,
    assigned_dimensions: Dimensions | None = None,
) -> StandardizedBoxAssignment:
    assigned = assigned_dimensions or vendor_box.dimensions
    return StandardizedBoxAssignment(
        order_id=carton.order_id,
        combination_key=carton.combination_key,
        box_type=box_type,
        optimized_length_cm=carton.optimized_dimensions.length,
        optimized_width_cm=carton.optimized_dimensions.width,
        optimized_height_cm=carton.optimized_dimensions.height,
        assigned_length_cm=assigned.length,
        assigned_width_cm=assigned.width,
        assigned_height_cm=assigned.height,
        box_standardization_note=note,
        placements=carton.placements,
        vendor_box_id=vendor_box.vendor_id,
        selection_decision=decision,
    )


def _custom_assignment(carton: OptimizedOrderCarton, index: int, demand: int) -> StandardizedBoxAssignment:
    return StandardizedBoxAssignment(
        order_id=carton.order_id,
        combination_key=carton.combination_key,
        box_type=f"Custom Box {index}",
        optimized_length_cm=carton.optimized_dimensions.length,
        optimized_width_cm=carton.optimized_dimensions.width,
        optimized_height_cm=carton.optimized_dimensions.height,
        assigned_length_cm=carton.optimized_dimensions.length,
        assigned_width_cm=carton.optimized_dimensions.width,
        assigned_height_cm=carton.optimized_dimensions.height,
        box_standardization_note=(
            f"Custom optimized carton used; demand {demand} meets 400 carton minimum."
        ),
        placements=carton.placements,
        vendor_box_id=None,
        selection_decision="custom_minimum_met",
    )


def _standardize_to_vendor_boxes(
    optimized_cartons: list[OptimizedOrderCarton],
    band_size_kg: float,
    custom_box_min_units: int,
    non_preferred_box_min_units: int,
    vendor_box_fit_tolerance_cm: float = 0,
    vendor_box_fit_tolerance_guardrail: bool = True,
    vendor_box_fit_tolerance_max_chargeable_increase_kg: float = 1.0,
) -> list[StandardizedBoxAssignment]:
    demand_by_dimensions = Counter(
        _dimensions_key(carton.optimized_dimensions)
        for carton in optimized_cartons
    )
    custom_names: dict[tuple[float, float, float], int] = {}
    assignments = []
    for carton in optimized_cartons:
        effective_vendor_box_fit_tolerance_cm = (
            vendor_box_fit_tolerance_cm if carton.allow_vendor_box_fit_tolerance else 0
        )
        baseline_vendor_candidates = _vendor_candidates(
            carton,
            VENDOR_BOXES,
            band_size_kg=band_size_kg,
            same_band_only=False,
            vendor_box_fit_tolerance_cm=0,
        )
        same_band_candidates = _vendor_candidates(
            carton,
            VENDOR_BOXES,
            band_size_kg=band_size_kg,
            same_band_only=True,
            vendor_box_fit_tolerance_cm=effective_vendor_box_fit_tolerance_cm,
        )
        if vendor_box_fit_tolerance_guardrail:
            same_band_candidates = _guarded_vendor_fit_candidates(
                same_band_candidates,
                baseline_vendor_candidates,
                vendor_box_fit_tolerance_max_chargeable_increase_kg,
            )
        if same_band_candidates:
            _billed, _chargeable, _volume, vendor_box, assigned_dimensions = same_band_candidates[0]
            assignments.append(
                _vendor_assignment(
                    carton,
                    f"Vendor Box {vendor_box.vendor_id}",
                    vendor_box,
                    f"Assigned vendor Box {vendor_box.vendor_id}; smallest safe vendor fit within billing band."
                    + _vendor_height_cutdown_note(assigned_dimensions, vendor_box)
                    + _vendor_fit_tolerance_note(assigned_dimensions, vendor_box),
                    "vendor_smallest_safe_same_band",
                    assigned_dimensions=assigned_dimensions,
                )
            )
            continue

        all_vendor_candidates = _vendor_candidates(
            carton,
            VENDOR_BOXES,
            band_size_kg=band_size_kg,
            same_band_only=False,
            vendor_box_fit_tolerance_cm=effective_vendor_box_fit_tolerance_cm,
        )
        if vendor_box_fit_tolerance_guardrail:
            all_vendor_candidates = _guarded_vendor_fit_candidates(
                all_vendor_candidates,
                baseline_vendor_candidates,
                vendor_box_fit_tolerance_max_chargeable_increase_kg,
            ) or baseline_vendor_candidates
        if all_vendor_candidates:
            _billed, _chargeable, _volume, vendor_box, assigned_dimensions = all_vendor_candidates[0]
            assignments.append(
                _vendor_assignment(
                    carton,
                    f"Vendor Box {vendor_box.vendor_id}",
                    vendor_box,
                    f"Assigned vendor Box {vendor_box.vendor_id}; smallest safe vendor fit."
                    + _vendor_height_cutdown_note(assigned_dimensions, vendor_box)
                    + _vendor_fit_tolerance_note(assigned_dimensions, vendor_box),
                    "vendor_smallest_safe_fit",
                    assigned_dimensions=assigned_dimensions,
                )
            )
            continue

        dimensions_key = _dimensions_key(carton.optimized_dimensions)
        demand = demand_by_dimensions[dimensions_key]
        if demand >= custom_box_min_units:
            custom_names.setdefault(dimensions_key, len(custom_names) + 1)
            assignments.append(
                _custom_assignment(
                    carton,
                    custom_names[dimensions_key],
                    demand,
                )
            )
            continue

        assignments.append(
            StandardizedBoxAssignment(
                order_id=carton.order_id,
                combination_key=carton.combination_key,
                box_type="NO-VENDOR-BOX",
                optimized_length_cm=carton.optimized_dimensions.length,
                optimized_width_cm=carton.optimized_dimensions.width,
                optimized_height_cm=carton.optimized_dimensions.height,
                assigned_length_cm=carton.optimized_dimensions.length,
                assigned_width_cm=carton.optimized_dimensions.width,
                assigned_height_cm=carton.optimized_dimensions.height,
                box_standardization_note="No vendor box can fit optimized carton dimensions.",
                placements=carton.placements,
                selection_decision="no_vendor_fit",
            )
        )
    return assignments

def _box_type_members(
    ordered: list[OptimizedOrderCarton],
    assignment_lookup: dict[str, _BoxType],
    candidate: _BoxType,
) -> list[OptimizedOrderCarton]:
    return [
        existing
        for existing in ordered
        if assignment_lookup.get(existing.order_id) is candidate
    ]


def standardize_optimized_cartons(
    optimized_cartons: list[OptimizedOrderCarton],
    tolerance_cm: float = 4,
    use_vendor_box_menu: bool = False,
    billing_band_kg: float = 0.5,
    custom_box_min_units: int = 400,
    non_preferred_box_min_units: int = 100,
    vendor_box_fit_tolerance_cm: float = 0,
    vendor_box_fit_tolerance_guardrail: bool = True,
    vendor_box_fit_tolerance_max_chargeable_increase_kg: float = 1.0,
) -> list[StandardizedBoxAssignment]:
    """Group optimized cartons into a practical shared campaign box menu."""
    if use_vendor_box_menu:
        return _standardize_to_vendor_boxes(
            optimized_cartons,
            band_size_kg=billing_band_kg,
            custom_box_min_units=custom_box_min_units,
            non_preferred_box_min_units=non_preferred_box_min_units,
            vendor_box_fit_tolerance_cm=vendor_box_fit_tolerance_cm,
            vendor_box_fit_tolerance_guardrail=vendor_box_fit_tolerance_guardrail,
            vendor_box_fit_tolerance_max_chargeable_increase_kg=vendor_box_fit_tolerance_max_chargeable_increase_kg,
        )

    ordered = sorted(
        optimized_cartons,
        key=lambda carton: (
            carton.combination_key,
            carton.optimized_dimensions.length,
            carton.optimized_dimensions.width,
            carton.optimized_dimensions.height,
            carton.order_id,
        ),
    )
    box_types: list[_BoxType] = []
    assigned_by_combo: dict[str, _BoxType] = {}
    assignment_lookup: dict[str, _BoxType] = {}
    notes_by_order: dict[str, str] = {}

    for carton in ordered:
        optimized_dimensions = carton.optimized_dimensions
        if not _within_cap(optimized_dimensions):
            raise ValueError(
                f"Optimized dimensions exceed max carton cap for {carton.order_id}"
            )

        box_type = assigned_by_combo.get(carton.combination_key)
        if box_type is not None and _can_round_up_to_box(
            optimized_dimensions,
            box_type.dimensions,
            tolerance_cm,
        ) and _is_billing_safe(carton, box_type.dimensions):
            assignment_lookup[carton.order_id] = box_type
            box_type.order_ids.add(carton.order_id)
            notes_by_order[carton.order_id] = _format_note(
                optimized_dimensions,
                box_type.dimensions,
            )
            continue

        skipped_unsafe_candidate = False
        best: tuple[tuple[float, float], _BoxType, Dimensions] | None = None
        for candidate in box_types:
            merged = _merged_dimensions(candidate.dimensions, optimized_dimensions)
            if not _within_cap(merged):
                continue
            if not _can_round_up_to_box(optimized_dimensions, merged, tolerance_cm):
                continue
            members = _box_type_members(ordered, assignment_lookup, candidate)
            if not all(_can_round_up_to_box(existing.optimized_dimensions, merged, tolerance_cm) for existing in members):
                continue
            if not _is_billing_safe(carton, merged) or not all(
                _is_billing_safe(existing, merged)
                for existing in members
            ):
                skipped_unsafe_candidate = True
                continue

            score = (
                dimensional_weight_kg(merged),
                (merged.length - candidate.dimensions.length)
                + (merged.width - candidate.dimensions.width)
                + (merged.height - candidate.dimensions.height)
            )
            if best is None or score < best[0]:
                best = (score, candidate, merged)

        if best is None:
            box_type = _BoxType(
                name=_box_type_name(len(box_types) + 1),
                dimensions=optimized_dimensions,
                combination_keys={carton.combination_key},
                order_ids={carton.order_id},
            )
            box_types.append(box_type)
            notes_by_order[carton.order_id] = (
                "Optimized dimensions used; no safe standardization candidate within billing band."
                if skipped_unsafe_candidate
                else _format_note(optimized_dimensions, optimized_dimensions)
            )
        else:
            _, box_type, merged = best
            box_type.dimensions = merged
            box_type.combination_keys.add(carton.combination_key)
            box_type.order_ids.add(carton.order_id)
            for member in _box_type_members(ordered, assignment_lookup, box_type):
                notes_by_order[member.order_id] = _format_note(
                    member.optimized_dimensions,
                    merged,
                )
            notes_by_order[carton.order_id] = _format_note(
                optimized_dimensions,
                merged,
            )

        assigned_by_combo[carton.combination_key] = box_type
        assignment_lookup[carton.order_id] = box_type

    return [
        StandardizedBoxAssignment(
            order_id=carton.order_id,
            combination_key=carton.combination_key,
            box_type=assignment_lookup[carton.order_id].name,
            optimized_length_cm=carton.optimized_dimensions.length,
            optimized_width_cm=carton.optimized_dimensions.width,
            optimized_height_cm=carton.optimized_dimensions.height,
            assigned_length_cm=assignment_lookup[carton.order_id].dimensions.length,
            assigned_width_cm=assignment_lookup[carton.order_id].dimensions.width,
            assigned_height_cm=assignment_lookup[carton.order_id].dimensions.height,
            box_standardization_note=notes_by_order.get(carton.order_id)
            or _format_note(
                carton.optimized_dimensions,
                assignment_lookup[carton.order_id].dimensions,
            ),
            placements=carton.placements,
        )
        for carton in optimized_cartons
    ]


def optimize_and_standardize_orders(
    orders: dict[str, tuple[str, list[PackedItem]]],
    tolerance_cm: float = 2,
) -> list[StandardizedBoxAssignment]:
    """Optimize each order, then standardize the campaign box menu."""
    optimized_cartons = []
    for order_id, (combination_key, items) in orders.items():
        result: OptimizedCartonResult = optimize_carton_dimensions(items)
        if not result.success:
            raise ValueError(f"Could not optimize carton for order {order_id}")
        optimized_cartons.append(
            OptimizedOrderCarton(
                order_id=order_id,
                combination_key=combination_key,
                optimized_dimensions=Dimensions(
                    length=result.length_cm,
                    width=result.width_cm,
                    height=result.height_cm,
                ),
                chargeable_weight_kg=result.chargeable_weight_kg,
                placements=result.placements,
            )
        )
    return standardize_optimized_cartons(optimized_cartons, tolerance_cm=tolerance_cm)
