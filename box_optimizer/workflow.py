"""Top-level workbook optimization workflow."""

from collections import defaultdict
from pathlib import Path

from box_optimizer.bundling import sku_combination_key
from box_optimizer.io.excel_reader import read_intake
from box_optimizer.io.excel_writer import write_workbook
from box_optimizer.models import Dimensions, OrderLine, PackedItem, SKUItem
from box_optimizer.packing.geometry import volume
from box_optimizer.packing.splitter import SplitResult, split_order_into_cartons
from box_optimizer.padding import add_final_exterior_padding, add_padding
from box_optimizer.standardization import (
    OptimizedOrderCarton,
    StandardizedBoxAssignment,
    standardize_optimized_cartons,
)
from box_optimizer.weights import KG_TO_LB, dimensional_weight_kg, packed_actual_weight_kg


DEFAULT_CONFIG = {
    "max_carton_cm": [74, 37, 44],
    "dimensional_divisor": 5000,
    "packing_weight_uplift": 1.15,
    "standardization_tolerance_cm": 2,
    "preserve_region_sheets": True,
}


_US_STATE_ABBREVIATIONS = {
    "alabama": "AL",
    "alaska": "AK",
    "arizona": "AZ",
    "arkansas": "AR",
    "california": "CA",
    "colorado": "CO",
    "connecticut": "CT",
    "delaware": "DE",
    "district of columbia": "DC",
    "florida": "FL",
    "georgia": "GA",
    "hawaii": "HI",
    "idaho": "ID",
    "illinois": "IL",
    "indiana": "IN",
    "iowa": "IA",
    "kansas": "KS",
    "kentucky": "KY",
    "louisiana": "LA",
    "maine": "ME",
    "maryland": "MD",
    "massachusetts": "MA",
    "michigan": "MI",
    "minnesota": "MN",
    "mississippi": "MS",
    "missouri": "MO",
    "montana": "MT",
    "nebraska": "NE",
    "nevada": "NV",
    "new hampshire": "NH",
    "new jersey": "NJ",
    "new mexico": "NM",
    "new york": "NY",
    "north carolina": "NC",
    "north dakota": "ND",
    "ohio": "OH",
    "oklahoma": "OK",
    "oregon": "OR",
    "pennsylvania": "PA",
    "rhode island": "RI",
    "south carolina": "SC",
    "south dakota": "SD",
    "tennessee": "TN",
    "texas": "TX",
    "utah": "UT",
    "vermont": "VT",
    "virginia": "VA",
    "washington": "WA",
    "west virginia": "WV",
    "wisconsin": "WI",
    "wyoming": "WY",
}


def _config(config: dict | None) -> dict:
    merged = dict(DEFAULT_CONFIG)
    if config:
        merged.update(config)
    return merged


def _sku_lookup(sku_items: list[SKUItem]) -> dict[str, SKUItem]:
    lookup = {}
    for item in sku_items:
        lookup[item.canonical_sku] = item
        for alias in item.aliases:
            lookup[str(alias).strip().upper()] = item
    return lookup


def _diagnostic_warnings(debug: dict) -> list[str]:
    warnings = []
    if debug.get("sku_items_parsed", 0) == 0:
        warnings.append("No SKU records parsed.")
    if debug.get("order_lines_created", 0) == 0:
        warnings.append("No order lines parsed.")
    if (
        debug.get("order_rows_read", 0) > 0
        and debug.get("order_lines_created", 0) == 0
        and debug.get("wide_product_columns_detected", 0) == 0
    ):
        warnings.append("No product quantity columns detected.")
    if debug.get("order_lines_created", 0) > 0 and debug.get("matched", 0) == 0:
        warnings.append("No matched SKUs found.")
    return warnings


def _group_order_lines(lines: list[OrderLine]) -> dict[str, list[OrderLine]]:
    grouped = defaultdict(list)
    for line in lines:
        grouped[line.order_id].append(line)
    return dict(grouped)


def _packed_items_for_order(
    lines: list[OrderLine],
    sku_lookup: dict[str, SKUItem],
) -> list[PackedItem]:
    items = []
    for line in lines:
        sku_item = sku_lookup[line.canonical_sku]
        unpadded = Dimensions(
            length=sku_item.length_cm,
            width=sku_item.width_cm,
            height=sku_item.height_cm,
        )
        items.append(
            PackedItem(
                canonical_sku=line.canonical_sku,
                quantity=line.quantity,
                unpadded_dimensions=unpadded,
                padded_dimensions=add_padding(unpadded),
                weight_kg=sku_item.weight_kg,
            )
        )
    return items


def _state_abbreviation(country: str | None, state_province: str | None) -> str:
    if not country or not state_province or country.strip().upper() not in {"US", "USA", "UNITED STATES"}:
        return ""
    normalized = state_province.strip()
    if len(normalized) == 2:
        return normalized.upper()
    return _US_STATE_ABBREVIATIONS.get(normalized.lower(), "")


def _metadata_for_order(lines: list[OrderLine]) -> dict:
    metadata = {}
    for line in lines:
        for key, value in line.metadata.items():
            metadata.setdefault(key, value)
    return metadata


def _append_metadata(row: dict, metadata: dict) -> dict:
    output = dict(row)
    for key, value in metadata.items():
        if key not in output:
            output[key] = value
        elif output[key] != value:
            metadata_key = f"Input {key}"
            suffix = 2
            while metadata_key in output:
                metadata_key = f"Input {key} {suffix}"
                suffix += 1
            output[metadata_key] = value
    return output


def _sku_breakdown(lines: list[OrderLine]) -> str:
    return sku_combination_key(lines)


def _distinct_skus(lines: list[OrderLine]) -> int:
    return len({line.canonical_sku for line in lines})


def _total_units(lines: list[OrderLine]) -> int:
    return sum(line.quantity for line in lines)


def _actual_weight_kg(items: list[PackedItem]) -> float:
    return sum(item.weight_kg * item.quantity for item in items)


def _padded_volume_cm3(items: list[PackedItem]) -> float:
    return sum(volume(item.padded_dimensions) * item.quantity for item in items)


def _carton_dimensions(split_result: SplitResult, box_index: int) -> Dimensions:
    carton = split_result.cartons[box_index].result
    return add_final_exterior_padding(
        Dimensions(
            length=carton.length_cm or 0,
            width=carton.width_cm or 0,
            height=carton.height_cm or 0,
        )
    )


def _build_standardization_inputs(
    split_results: dict[str, SplitResult],
    combo_by_order: dict[str, str],
) -> list[OptimizedOrderCarton]:
    optimized = []
    for order_id, split_result in split_results.items():
        for index, carton in enumerate(split_result.cartons):
            dimensions = _carton_dimensions(split_result, index)
            optimized.append(
                OptimizedOrderCarton(
                    order_id=f"{order_id}#{index + 1}",
                    combination_key=combo_by_order[order_id],
                    optimized_dimensions=dimensions,
                    chargeable_weight_kg=carton.result.chargeable_weight_kg or 0,
                    placements=carton.result.placements,
                )
            )
    return optimized


def _assignment_lookup(
    assignments: list[StandardizedBoxAssignment],
) -> dict[str, StandardizedBoxAssignment]:
    return {assignment.order_id: assignment for assignment in assignments}


def _summary_rows(result: dict) -> list[dict]:
    return [
        {"Metric": "Orders Processed", "Value": result["orders_processed"]},
        {"Metric": "Boxes Created", "Value": result["boxes_created"]},
        {"Metric": "Box Types", "Value": result["box_types"]},
        {"Metric": "Unmatched SKUs", "Value": result["unmatched_skus"]},
    ]


def _box_size_summary(assignments: list[StandardizedBoxAssignment]) -> list[dict]:
    by_type = {}
    for assignment in assignments:
        by_type[assignment.box_type] = {
            "Box Type": assignment.box_type,
            "Assigned Box Length cm": assignment.assigned_length_cm,
            "Assigned Box Width cm": assignment.assigned_width_cm,
            "Assigned Box Height cm": assignment.assigned_height_cm,
        }
    return list(by_type.values())


def _unmatched_rows(unmatched_skus) -> list[dict]:
    rows = []
    for unmatched in unmatched_skus:
        row = {
            "Order ID": unmatched.order_line.order_id,
            "Raw SKU": unmatched.order_line.raw_sku,
            "Canonical SKU": unmatched.order_line.canonical_sku,
            "Reason": unmatched.reason,
        }
        row.update(unmatched.metadata)
        rows.append(row)
    return rows


def _packing_detail_rows(split_results: dict[str, SplitResult]) -> list[dict]:
    rows = []
    for order_id, split_result in split_results.items():
        for carton in split_result.cartons:
            for placement in carton.result.placements:
                x, y, z = placement.origin
                rows.append(
                    {
                        "Order ID": order_id,
                        "Box Number": carton.box_number,
                        "Canonical SKU": placement.canonical_sku,
                        "Quantity": placement.quantity,
                        "X cm": x,
                        "Y cm": y,
                        "Z cm": z,
                        "Length cm": placement.dimensions.length,
                        "Width cm": placement.dimensions.width,
                        "Height cm": placement.dimensions.height,
                    }
                )
    return rows


def _multi_box_rows(split_results: dict[str, SplitResult]) -> list[dict]:
    return [
        {"Order ID": order_id, "Box Qty": split_result.box_qty}
        for order_id, split_result in split_results.items()
        if split_result.box_qty > 1
    ]


def optimize_workbook(
    sku_master_path: str,
    orders_path: str,
    output_path: str,
    config: dict | None = None,
) -> dict:
    """Optimize a SKU/order workbook pair and write an output workbook."""
    cfg = _config(config)
    warnings = []
    if cfg["max_carton_cm"] != DEFAULT_CONFIG["max_carton_cm"]:
        warnings.append("Custom max_carton_cm is not yet supported; using 74 x 37 x 44 cm.")
    if cfg["dimensional_divisor"] != DEFAULT_CONFIG["dimensional_divisor"]:
        warnings.append("Custom dimensional_divisor is not yet supported; using 5000.")
    if cfg["packing_weight_uplift"] != DEFAULT_CONFIG["packing_weight_uplift"]:
        warnings.append("Custom packing_weight_uplift is not yet supported; using 1.15.")

    intake = read_intake(sku_master_path, orders_path)
    if intake.unmatched_skus:
        warnings.append(f"{len(intake.unmatched_skus)} unmatched SKU rows were preserved.")
    warnings.extend(_diagnostic_warnings(intake.debug))

    sku_items = _sku_lookup(intake.sku_items)
    grouped_orders = _group_order_lines(intake.matched_order_lines)
    split_results = {}
    combo_by_order = {}
    items_by_order = {}
    failed_orders = []

    for order_id, lines in grouped_orders.items():
        items = _packed_items_for_order(lines, sku_items)
        split_result = split_order_into_cartons(items)
        if split_result.success:
            split_results[order_id] = split_result
            combo_by_order[order_id] = _sku_breakdown(lines)
            items_by_order[order_id] = items
        else:
            failed_orders.append(order_id)

    if failed_orders:
        warnings.append(f"{len(failed_orders)} orders could not be packed.")

    assignments = standardize_optimized_cartons(
        _build_standardization_inputs(split_results, combo_by_order),
        tolerance_cm=cfg["standardization_tolerance_cm"],
    )
    assignments_by_key = _assignment_lookup(assignments)
    order_rows = []

    for order_id, lines in grouped_orders.items():
        if order_id not in split_results:
            continue
        split_result = split_results[order_id]
        first_line = lines[0]
        actual_weight_kg = _actual_weight_kg(items_by_order[order_id])
        packed_weight_kg = packed_actual_weight_kg(actual_weight_kg)
        for index, carton in enumerate(split_result.cartons):
            assignment = assignments_by_key[f"{order_id}#{index + 1}"]
            optimized_dimensions = _carton_dimensions(split_result, index)
            assigned_dimensions = Dimensions(
                assignment.assigned_length_cm,
                assignment.assigned_width_cm,
                assignment.assigned_height_cm,
            )
            dim_weight_kg = dimensional_weight_kg(assigned_dimensions)
            chargeable_kg = max(packed_weight_kg, dim_weight_kg)
            row = {
                "Region": first_line.region or "",
                "Order ID": order_id,
                "Country": first_line.country or "",
                "State/Province": first_line.state_province or "",
                "US State Abbreviation": _state_abbreviation(first_line.country, first_line.state_province),
                "Packed Actual Weight kg": packed_weight_kg,
                "Dimensional Weight kg (/5000)": dim_weight_kg,
                "Chargeable Weight kg": chargeable_kg,
                "Chargeable Weight g": chargeable_kg * 1000,
                "Total Units": _total_units(lines),
                "Box Qty": split_result.box_qty,
                "Box Type": assignment.box_type,
                "Length cm": assigned_dimensions.length,
                "Width cm": assigned_dimensions.width,
                "Height cm": assigned_dimensions.height,
                "Optimized Length cm": optimized_dimensions.length,
                "Optimized Width cm": optimized_dimensions.width,
                "Optimized Height cm": optimized_dimensions.height,
                "Assigned Box Length cm": assignment.assigned_length_cm,
                "Assigned Box Width cm": assignment.assigned_width_cm,
                "Assigned Box Height cm": assignment.assigned_height_cm,
                "Box Standardization Note": assignment.box_standardization_note,
                "Actual Item Weight lb": actual_weight_kg * KG_TO_LB,
                "Packed Actual Weight lb (+15%)": packed_weight_kg * KG_TO_LB,
                "Bundled/Padded Volume cm³": _padded_volume_cm3(items_by_order[order_id]),
                "Dimensional Weight lb": dim_weight_kg * KG_TO_LB,
                "Chargeable Weight lb": chargeable_kg * KG_TO_LB,
                "Distinct SKUs": _distinct_skus(lines),
                "SKU Breakdown": combo_by_order[order_id],
            }
            order_rows.append(_append_metadata(row, _metadata_for_order(lines)))

    result = {
        "output_path": str(Path(output_path)),
        "orders_processed": len(split_results),
        "boxes_created": sum(split_result.box_qty for split_result in split_results.values()),
        "box_types": len({assignment.box_type for assignment in assignments}),
        "unmatched_skus": len(intake.unmatched_skus),
        "warnings": warnings,
    }

    region_sheets = {}
    if cfg["preserve_region_sheets"]:
        rows_by_region = defaultdict(list)
        for row in order_rows:
            if row.get("Region"):
                rows_by_region[f"Region - {row['Region']}"].append(row)
        region_sheets = dict(rows_by_region)

    write_workbook(
        output_path,
        summary_rows=_summary_rows(result),
        order_volume_weights_rows=order_rows,
        box_size_summary_rows=_box_size_summary(assignments),
        unmatched_skus_rows=_unmatched_rows(intake.unmatched_skus),
        packing_detail_rows=_packing_detail_rows(split_results),
        multi_box_detail_rows=_multi_box_rows(split_results),
        input_column_mapping_rows=intake.column_mappings,
        errors_and_warnings_rows=[{"Warning": warning} for warning in warnings],
        sheets=region_sheets,
    )

    return result
