"""Top-level workbook optimization workflow."""

import hashlib
import json
import logging
import math
import re
import time
import zipfile
from collections import Counter, defaultdict
from xml.etree import ElementTree
from copy import deepcopy
from dataclasses import dataclass, replace
from itertools import permutations
from pathlib import Path

from box_optimizer.bundling import sku_combination_key
from box_optimizer.io.excel_reader import IntakeResult, read_intake, read_workbook
from box_optimizer.io.excel_writer import write_workbook
from box_optimizer.models import Dimensions, OrderLine, PackedItem, SKUItem
from box_optimizer.normalize import normalize_sku
from box_optimizer.packing.geometry import volume
from box_optimizer.packing.packer import (
    MAX_CARTON_DIMENSIONS,
    OptimizedCartonResult,
    Placement,
    _expand_items,
    pack_items,
)
from box_optimizer.packing.splitter import SplitCarton, SplitResult, split_order_into_cartons
from box_optimizer.padding import add_final_exterior_padding, add_padding
from box_optimizer.standardization import (
    OptimizedOrderCarton,
    PREFERRED_VENDOR_BOX_IDS,
    StandardizedBoxAssignment,
    standardize_optimized_cartons,
)
from box_optimizer.weights import KG_TO_LB, dimensional_weight_kg, packed_actual_weight_kg


DEFAULT_CONFIG = {
    "max_carton_cm": [74, 37, 44],
    "dimensional_divisor": 5000,
    "packing_weight_uplift": 1.15,
    "standardization_tolerance_cm": 4,
    "use_vendor_box_menu": True,
    "billing_band_kg": 1.0,
    "custom_box_min_units": 400,
    "non_preferred_box_min_units": 100,
    "preserve_region_sheets": True,
    "debug": False,
    "max_orders": None,
    "packing_mode": "normal",
    "max_optimization_seconds": 180,
    "balanced_max_items_for_deep_search": 18,
    "balanced_max_item_quantity_for_recombine": 10,
    "balanced_min_remaining_seconds": 3,
    "bundle_footprint_tolerance_cm": 5,
    "chargeable_weight_split_savings_threshold_kg": 0.5,
    "chargeable_weight_split_savings_threshold_pct": 0.05,
    "chargeable_weight_split_two_extra_box_threshold_kg": 2.0,
    "max_extra_boxes_per_order": 1,
    "oversized_vendor_box_ids": ["36", "37", "41"],
    "oversized_vendor_box_chargeable_threshold_kg": 20.0,
    "oversized_max_extra_boxes_per_order": 2,
    "non_preferred_extra_box_savings_threshold_kg": 1.0,
    "non_preferred_extra_box_savings_threshold_pct": 0.075,
    "non_preferred_two_extra_box_savings_threshold_kg": 3.0,
    "non_preferred_two_extra_box_savings_threshold_pct": 0.10,
    "company_protection_extra_box_guardrail": True,
    "company_protection_rate_bands": None,
    "company_protection_zone": None,
    "company_protection_zone_markups": {"Zone USA": 1.25, "default": 1.30},
    "company_protection_max_rate_weight_kg": 49.0,
    "rate_sheet_path": "ShippingRatesHub2026.xlsx",
    "company_protection_min_margin_delta": 0.01,
    "repeat_retail_batch_planning_enabled": True,
    "repeat_retail_min_repeated_units": 24,
    "repeat_retail_batch_sizes": [8, 10, 12, 16, 20],
    "repeat_retail_max_extra_boxes_per_order": 16,
    "repeat_retail_max_candidate_boxes": 24,
    "repeat_retail_min_optimization_seconds": 450,
    "repeat_retail_min_savings_threshold_kg": 1.0,
    "repeat_retail_min_savings_threshold_pct": 0.03,
    "repeat_retail_max_margin_giveback": 5.0,
    "repeat_retail_min_customer_savings": 5.0,
    "vendor_box_fit_mode": "auto",
    "vendor_box_fit_tolerance_cm": 1.5,
    "vendor_box_fit_tolerance_max_cm": 2.0,
    "vendor_box_fit_tolerance_guardrail": True,
    "vendor_box_fit_tolerance_max_chargeable_increase_kg": 1.0,
    "output_granularity": "order_summary",
}


logger = logging.getLogger("box_optimizer")


@dataclass(frozen=True)
class WorkflowWarning:
    """Structured warning row for the workbook diagnostics tab."""

    order_id: str
    stage: str
    error_type: str
    message: str
    sku_breakdown: str = ""
    sku: str = ""
    rule_applied: str = ""


@dataclass(frozen=True)
class RuleSplitGroup:
    """Order lines grouped for either rule-required separation or normal packing."""

    lines: list[OrderLine]
    reason: str = ""
    rule_keys: tuple[str, ...] = ()


@dataclass(frozen=True)
class CachedPackingPlan:
    """Order-neutral packing output reused for matching pledge combinations."""

    split_result: SplitResult
    group_warnings: tuple[WorkflowWarning, ...] = ()


@dataclass(frozen=True)
class CandidatePlanScore:
    """Comparable score for a fully assigned candidate packing plan."""

    total_chargeable_weight_kg: float
    box_qty: int
    box_type_count: int
    total_assigned_volume_cm3: float
    oversized_box_count: int = 0
    non_preferred_box_count: int = 0
    package_chargeable_weights_kg: tuple[float, ...] = ()


@dataclass(frozen=True)
class PackingOrderContext:
    """Precomputed packing inputs for one order in the workflow cache loop."""

    order_id: str
    lines: list[OrderLine]
    combo: str
    items: list[PackedItem]
    groups: list[RuleSplitGroup]
    cache_key: str
    first_index: int


@dataclass(frozen=True)
class SKUCampaignRule:
    """Runtime campaign rule matched to a SKU or product name."""

    key: str
    no_padding: bool = False
    prepacked: bool = False
    ships_alone: bool = False
    can_mix_with_other_items: bool = True
    forced_box_cm: Dimensions | None = None
    must_stay_flat: bool = False
    allow_rotation: bool = True
    allowed_orientations: tuple[Dimensions, ...] | None = None
    extra_padding_cm: Dimensions | None = None
    wrap_around_largest_item: bool = False
    wrapped_height_cm: float = 4
    compressible: bool = False
    compressed_height_ratio: float = 0.6
    compressed_volume_ratio: float = 0.75
    exclude_from_standardization: bool = False
    box_type: str | None = None
    warning_note: str = ""


_US_STATE_ABBREVIATIONS = {
    'aa': 'AA',
    'ae': 'AE',
    'ak': 'AK',
    'al': 'AL',
    'alabama': 'AL',
    'alaska': 'AK',
    'american samoa': 'AS',
    'ap': 'AP',
    'ar': 'AR',
    'arizona': 'AZ',
    'arkansas': 'AR',
    'armed forces': 'AP',
    'armed forces aa': 'AA',
    'armed forces ae': 'AE',
    'armed forces ap': 'AP',
    'as': 'AS',
    'az': 'AZ',
    'ca': 'CA',
    'california': 'CA',
    'co': 'CO',
    'colorado': 'CO',
    'connecticut': 'CT',
    'ct': 'CT',
    'dc': 'DC',
    'de': 'DE',
    'delaware': 'DE',
    'district of columbia': 'DC',
    'fl': 'FL',
    'florida': 'FL',
    'ga': 'GA',
    'georgia': 'GA',
    'gu': 'GU',
    'guam': 'GU',
    'hawaii': 'HI',
    'hi': 'HI',
    'ia': 'IA',
    'id': 'ID',
    'idaho': 'ID',
    'il': 'IL',
    'illinois': 'IL',
    'in': 'IN',
    'indiana': 'IN',
    'iowa': 'IA',
    'kansas': 'KS',
    'kentucky': 'KY',
    'ks': 'KS',
    'ky': 'KY',
    'la': 'LA',
    'louisiana': 'LA',
    'ma': 'MA',
    'maine': 'ME',
    'maryland': 'MD',
    'massachusetts': 'MA',
    'md': 'MD',
    'me': 'ME',
    'mi': 'MI',
    'michigan': 'MI',
    'minnesota': 'MN',
    'mississippi': 'MS',
    'missouri': 'MO',
    'mn': 'MN',
    'mo': 'MO',
    'montana': 'MT',
    'mp': 'MP',
    'ms': 'MS',
    'mt': 'MT',
    'nc': 'NC',
    'nd': 'ND',
    'ne': 'NE',
    'nebraska': 'NE',
    'nevada': 'NV',
    'new hampshire': 'NH',
    'new jersey': 'NJ',
    'new mexico': 'NM',
    'new york': 'NY',
    'nh': 'NH',
    'nj': 'NJ',
    'nm': 'NM',
    'north carolina': 'NC',
    'north dakota': 'ND',
    'northern mariana islands': 'MP',
    'nv': 'NV',
    'ny': 'NY',
    'oh': 'OH',
    'ohio': 'OH',
    'ok': 'OK',
    'oklahoma': 'OK',
    'or': 'OR',
    'oregon': 'OR',
    'pa': 'PA',
    'pennsylvania': 'PA',
    'pr': 'PR',
    'puerto rico': 'PR',
    'rhode island': 'RI',
    'ri': 'RI',
    'sc': 'SC',
    'sd': 'SD',
    'south carolina': 'SC',
    'south dakota': 'SD',
    'tennessee': 'TN',
    'texas': 'TX',
    'tn': 'TN',
    'tx': 'TX',
    'u s virgin islands': 'VI',
    'ut': 'UT',
    'utah': 'UT',
    'va': 'VA',
    'vermont': 'VT',
    'vi': 'VI',
    'virginia': 'VA',
    'vt': 'VT',
    'wa': 'WA',
    'washington': 'WA',
    'west virginia': 'WV',
    'wi': 'WI',
    'wisconsin': 'WI',
    'wv': 'WV',
    'wy': 'WY',
    'wyoming': 'WY',
}

_US_STATE_CODES = {
    'AA', 'AE', 'AK', 'AL', 'AP', 'AR', 'AS', 'AZ', 'CA', 'CO', 'CT', 'DC', 'DE', 'FL', 'GA', 'GU', 'HI', 'IA', 'ID', 'IL', 'IN', 'KS', 'KY', 'LA', 'MA', 'MD', 'ME', 'MI', 'MN', 'MO', 'MP', 'MS', 'MT', 'NC', 'ND', 'NE', 'NH', 'NJ', 'NM', 'NV', 'NY', 'OH', 'OK', 'OR', 'PA', 'PR', 'RI', 'SC', 'SD', 'TN', 'TX', 'UT', 'VA', 'VI', 'VT', 'WA', 'WI', 'WV', 'WY',
}

_USA_PLUS_ZONE_STATES = {"AA", "AE", "AK", "AP", "AS", "GU", "HI", "MP", "PR", "VI"}

_COUNTRY_NAME_BY_CODE = {
    'AD': 'Andorra',
    'AE': 'United Arab Emirates',
    'AF': 'Afghanistan',
    'AG': 'Antigua and Barbuda',
    'AI': 'Anguilla',
    'AL': 'Albania',
    'AM': 'Armenia',
    'AN': 'Netherlands Antilles',
    'AO': 'Cabinda',
    'AR': 'Argentina',
    'AS': 'American Samoa',
    'AT': 'Austria',
    'AU': 'Australia',
    'AW': 'Aruba',
    'AZ': 'Azerbaijan',
    'BA': 'Bosnia and Herzegovina',
    'BB': 'Barbados',
    'BD': 'Bangladesh',
    'BE': 'Belgium',
    'BF': 'Burkina Faso',
    'BG': 'Bulgaria',
    'BH': 'Bahrain',
    'BI': 'Burundi',
    'BJ': 'Benin',
    'BM': 'Bermuda',
    'BN': 'Brunei',
    'BO': 'Bolivia',
    'BR': 'Brazil',
    'BS': 'Bahamas',
    'BT': 'Bhutan',
    'BW': 'Botswana',
    'BY': 'Belarus',
    'BZ': 'Belize',
    'CA': 'Canada',
    'CC': 'Cocos Island',
    'CD': 'Zaire (DRC)',
    'CF': 'Central African Republic',
    'CG': 'Congo',
    'CH': 'Switzerland',
    'CI': 'Ivory coast',
    'CK': 'Cook islands',
    'CL': 'Easter Island',
    'CM': 'Cameroon',
    'CN': 'China',
    'CO': 'Colombia',
    'CR': 'Costa Rica',
    'CU': 'Cuba',
    'CV': 'Cape Verde',
    'CX': 'Christmas Island',
    'CY': 'Cyprus',
    'CZ': 'Czech Republic',
    'DE': 'Germany',
    'DJ': 'Djibouti',
    'DK': 'Denmark',
    'DM': 'Dominica',
    'DZ': 'Algeria',
    'EC': 'Ecuador',
    'EE': 'Estonia',
    'EG': 'Egypt',
    'EH': 'Western Sahara',
    'ER': 'Eritrea',
    'ES': 'Spain',
    'ET': 'Ethiopia',
    'FI': 'Finland',
    'FJ': 'Fiji',
    'FK': 'Falkland Islands',
    'FM': 'Micronesia',
    'FO': 'Faroe Islands',
    'FR': 'France',
    'GA': 'Gabon',
    'GB': 'United Kingdom',
    'GD': 'Grenada',
    'GE': 'Georgia',
    'GF': 'French Guiana',
    'GI': 'Gibraltar',
    'GL': 'Greenland',
    'GM': 'Gambia',
    'GN': 'Guinea',
    'GR': 'Greece',
    'GT': 'Guatemala',
    'GU': 'Guam',
    'GY': 'Guyana',
    'HK': 'Hong Kong',
    'HN': 'Honduras',
    'HR': 'Croatia',
    'HT': 'Haiti',
    'HU': 'Hungary',
    'IC': 'Canary Islands',
    'ID': 'Indonesia',
    'IE': 'Ireland',
    'IL': 'Israel',
    'IN': 'India',
    'IO': 'British Indian Ocean Territory',
    'IQ': 'Iraq',
    'IR': 'Iran',
    'IS': 'Iceland',
    'IT': 'Italy',
    'JM': 'Jamaica',
    'JO': 'Jordan',
    'JP': 'Japan',
    'KE': 'Kenya',
    'KG': 'Kyrgyzstan',
    'KH': 'Cambodia',
    'KI': 'Kiribati',
    'KM': 'Comoros',
    'KN': 'St. Kitts and Nevis',
    'KP': 'Korea, North',
    'KR': 'South Korea',
    'KW': 'Kuwait',
    'KY': 'Cayman Islands',
    'KZ': 'Kazakhstan',
    'LA': 'Laos',
    'LB': 'Lebanon',
    'LC': 'Saint Lucia',
    'LI': 'Liechtenstein',
    'LK': 'Sri Lanka',
    'LS': 'Lesotho',
    'LT': 'Lithuania',
    'LU': 'Luxembourg',
    'LV': 'Latvia',
    'LY': 'Libya',
    'MA': 'Morocco',
    'MC': 'Monaco',
    'MD': 'Moldova',
    'ME': 'Montenegro',
    'MG': 'Madagascar',
    'MH': 'Marshall Islands',
    'MK': 'Macedonia',
    'ML': 'Mali',
    'MM': 'Myanmar',
    'MN': 'Mongolia',
    'MP': 'Northern Mariana Islands',
    'MQ': 'Martinique',
    'MR': 'Mauritania',
    'MS': 'Montserrat',
    'MT': 'Malta',
    'MU': 'Mauritius',
    'MV': 'Maldives',
    'MW': 'Malawi',
    'MX': 'Mexico',
    'MY': 'Malaysia',
    'MZ': 'Mozambique',
    'NA': 'Namibia',
    'NC': 'New Caledonia',
    'NE': 'Niger',
    'NF': 'Norfolk Island',
    'NG': 'Nigeria',
    'NI': 'Nicaragua',
    'NL': 'Netherlands',
    'NO': 'Norway',
    'NP': 'Nepal',
    'NR': 'Nauru',
    'NU': 'Niue',
    'NZ': 'New Zealand',
    'OM': 'Oman',
    'PA': 'Panama',
    'PF': 'French Polynesia',
    'PG': 'Papua New Guinea',
    'PH': 'Philippines',
    'PK': 'Pakistan',
    'PL': 'Poland',
    'PM': 'Saint Pierre and Miquelon',
    'PR': 'Puerto Rico',
    'PS': 'Palestine',
    'PT': 'Portugal',
    'PW': 'Palau',
    'PY': 'Paraguay',
    'QA': 'Qatar',
    'RE': 'Reunion',
    'RO': 'Romania',
    'RS': 'Serbia',
    'RU': 'Russia',
    'RW': 'Rwanda',
    'SA': 'Saudi Arabia',
    'SB': 'Solomon islands',
    'SC': 'Seychelles',
    'SE': 'Sweden',
    'SG': 'Singapore',
    'SH': 'St. Helena',
    'SI': 'Slovenia',
    'SJ': 'Svalbard and Jan Mayen',
    'SK': 'Slovakia',
    'SL': 'Sierra Leone',
    'SN': 'Senegal',
    'SO': 'Somalia',
    'SR': 'Suriname',
    'ST': 'Sao Tome and Principe',
    'SY': 'Syria',
    'SZ': 'Swaziland (Eswatini)',
    'TC': 'Turks and Caicos Islands',
    'TD': 'Chad',
    'TF': 'French Southern Territories',
    'TG': 'Togo',
    'TH': 'Thailand',
    'TJ': 'Tajikistan',
    'TL': 'East Timor',
    'TM': 'Turkmenistan',
    'TN': 'Tunisia',
    'TO': 'Tonga',
    'TR': 'Turkey',
    'TT': 'Trinidad and Tobago',
    'TV': 'Tuvalu',
    'TW': 'Taiwan',
    'TZ': 'Tanzania',
    'UA': 'Ukraine',
    'UG': 'Uganda',
    'UM': 'United States Minor Outlying Islands',
    'US': 'United States',
    'UY': 'Uruguay',
    'UZ': 'Uzbekistan',
    'VA': 'Vatican City',
    'VC': 'Saint Vincent and the Grenadines',
    'VE': 'Venezuela',
    'VG': 'British Virgin Islands',
    'VI': 'US Virgin Islands',
    'VN': 'Vietnam',
    'VU': 'Vanuatu',
    'WF': 'Wallis and Futuna',
    'YE': 'Yemen',
    'YT': 'Mayotte',
    'ZA': 'South Africa',
    'ZM': 'Zambia',
    'ZW': 'Zimbabwe',
}

_COUNTRY_ALIASES = {
    'america': 'US',
    'bolivia plurinational state of': 'BO',
    'britain': 'GB',
    'brunei darussalam': 'BN',
    'congo democratic republic': 'CD',
    'cote d ivoire': 'CI',
    'czechia': 'CZ',
    'democratic republic of congo': 'CD',
    'democratic republic of the congo': 'CD',
    'drc': 'CD',
    'england': 'GB',
    'great britain': 'GB',
    'ivory coast': 'CI',
    'korea': 'KR',
    'korea north': 'KP',
    'korea republic of': 'KR',
    'korea south': 'KR',
    'macao': 'MO',
    'moldova republic of': 'MD',
    'north korea': 'KP',
    'republic of korea': 'KR',
    'russian federation': 'RU',
    's korea': 'KR',
    'saint kitts and nevis': 'KN',
    'saint lucia': 'LC',
    'saint vincent and the grenadines': 'VC',
    'south korea': 'KR',
    'st kitts and nevis': 'KN',
    'st lucia': 'LC',
    'st vincent and the grenadines': 'VC',
    'taiwan province of china': 'TW',
    'tanzania united republic of': 'TZ',
    'u k': 'GB',
    'u s': 'US',
    'u s a': 'US',
    'uk': 'GB',
    'united states': 'US',
    'united states of america': 'US',
    'us': 'US',
    'usa': 'US',
    'venezuela bolivarian republic of': 'VE',
    'viet nam': 'VN',
    'zaire': 'CD',
}

def _normalize_lookup_key(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(value or "").strip().lower()).strip()


def _config(config: dict | None) -> dict:
    merged = dict(DEFAULT_CONFIG)
    if config:
        merged.update(config)
    return merged


def format_kg_display(value: float, decimals: int = 1, mode: str = "truncate") -> float:
    """Format kg display values without changing internal calculations."""
    factor = 10**decimals
    if mode == "truncate":
        return math.trunc(float(value) * factor) / factor
    return round(float(value), decimals)


def _format_weight_display(value: float) -> float:
    return format_kg_display(value)


def _format_lb_display(value: float) -> float:
    return format_kg_display(value)


def _ceil_cm(value: float) -> int:
    return int(math.ceil(float(value)))


def _display_dimensions(dimensions: Dimensions, cap: bool = True) -> Dimensions:
    if not cap:
        return Dimensions(
            length=_ceil_cm(dimensions.length),
            width=_ceil_cm(dimensions.width),
            height=_ceil_cm(dimensions.height),
        )
    return Dimensions(
        length=min(_ceil_cm(dimensions.length), int(MAX_CARTON_DIMENSIONS.length)),
        width=min(_ceil_cm(dimensions.width), int(MAX_CARTON_DIMENSIONS.width)),
        height=min(_ceil_cm(dimensions.height), int(MAX_CARTON_DIMENSIONS.height)),
    )


def _placement_top_height(placements: list) -> float:
    top_height = 0.0
    for placement in placements:
        if placement.origin is None:
            top_height = max(top_height, placement.dimensions.height)
            continue
        z = placement.origin.z if hasattr(placement.origin, "z") else placement.origin[2]
        top_height = max(top_height, z + placement.dimensions.height)
    return top_height


def _vendor_height_cutdown(
    *,
    dimensions: Dimensions,
    optimized_dimensions: Dimensions,
    placements: list,
    vendor_box_id: str,
) -> tuple[Dimensions, str]:
    if not vendor_box_id or not placements:
        return dimensions, ""
    packed_height = _placement_top_height(placements)
    cut_height = _ceil_cm(packed_height + 2) if packed_height else optimized_dimensions.height
    cut_height = min(dimensions.height, cut_height)
    if cut_height >= dimensions.height:
        return dimensions, ""
    cut_dimensions = Dimensions(
        length=dimensions.length,
        width=dimensions.width,
        height=cut_height,
    )
    return cut_dimensions, f"Vendor box height cut down to {cut_height:g} cm."


def _rule_key(value: object) -> str:
    text = str(value or "").strip().lower()
    text = text.replace("â€œ", "\"").replace("â€", "\"").replace("â€™", "'").replace("Ã—", "x")
    return re.sub(r"[\W_]+", "", text)


def _bool_from_config(value: object, default: bool = False) -> bool:
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        normalized = value.strip().lower()
        if normalized in {"true", "1", "yes", "y", "on"}:
            return True
        if normalized in {"false", "0", "no", "n", "off"}:
            return False
    return bool(value)


def _dimensions_from_config(value: object) -> Dimensions | None:
    if value is None:
        return None
    if not isinstance(value, list | tuple) or len(value) != 3:
        return None
    try:
        length, width, height = [float(part) for part in value]
    except (TypeError, ValueError):
        return None
    sorted_values = sorted([length, width, height], reverse=True)
    return Dimensions(*sorted_values)


def _parse_sku_rules(config: dict) -> dict[str, SKUCampaignRule]:
    parsed = {}
    for key, rule in (config.get("sku_rules") or {}).items():
        if not isinstance(rule, dict):
            continue
        orientations = rule.get("allowed_orientations")
        allowed_orientations = None
        if orientations:
            parsed_orientations = [
                dimensions
                for dimensions in (_dimensions_from_config(value) for value in orientations)
                if dimensions is not None
            ]
            allowed_orientations = tuple(parsed_orientations) or None
        parsed[key] = SKUCampaignRule(
            key=str(key),
            no_padding=_bool_from_config(rule.get("no_padding"), False),
            prepacked=_bool_from_config(rule.get("prepacked"), False),
            ships_alone=_bool_from_config(rule.get("ships_alone"), False),
            can_mix_with_other_items=_bool_from_config(rule.get("can_mix_with_other_items"), True),
            forced_box_cm=_dimensions_from_config(rule.get("forced_box_cm")),
            must_stay_flat=_bool_from_config(rule.get("must_stay_flat"), False),
            allow_rotation=_bool_from_config(rule.get("allow_rotation"), True),
            allowed_orientations=allowed_orientations,
            extra_padding_cm=_dimensions_from_config(rule.get("extra_padding_cm")),
            wrap_around_largest_item=_bool_from_config(rule.get("wrap_around_largest_item"), False),
            wrapped_height_cm=float(rule.get("wrapped_height_cm", 4) or 4),
            compressible=_bool_from_config(rule.get("compressible"), False),
            compressed_height_ratio=float(rule.get("compressed_height_ratio", 0.6) or 0.6),
            compressed_volume_ratio=float(rule.get("compressed_volume_ratio", 0.75) or 0.75),
            exclude_from_standardization=_bool_from_config(rule.get("exclude_from_standardization"), False),
            box_type=rule.get("box_type"),
            warning_note=str(rule.get("warning_note", "") or ""),
        )
    return parsed


def _rule_candidates_for_sku(item: SKUItem) -> set[str]:
    return {
        normalize_sku(item.raw_sku),
        normalize_sku(item.canonical_sku),
        normalize_sku(item.product_name),
        _rule_key(item.raw_sku),
        _rule_key(item.canonical_sku),
        _rule_key(item.product_name),
    }


def _rule_candidates_for_line(line: OrderLine) -> set[str]:
    return {
        normalize_sku(line.raw_sku),
        normalize_sku(line.canonical_sku),
        _rule_key(line.raw_sku),
        _rule_key(line.canonical_sku),
    }


def _match_sku_rules(
    config: dict,
    sku_items: list[SKUItem],
    order_lines: list[OrderLine],
) -> tuple[dict[str, SKUCampaignRule], list[str], list[str]]:
    rules = _parse_sku_rules(config)
    normalized_rules = {
        candidate: rule
        for rule in rules.values()
        for candidate in {normalize_sku(rule.key), _rule_key(rule.key)}
    }
    matches = {}
    matched_keys = set()
    for item in sku_items:
        rule = next(
            (
                normalized_rules[candidate]
                for candidate in _rule_candidates_for_sku(item)
                if candidate in normalized_rules
            ),
            None,
        )
        if rule:
            matches[item.canonical_sku] = rule
            matched_keys.add(rule.key)
    for line in order_lines:
        rule = next(
            (
                normalized_rules[candidate]
                for candidate in _rule_candidates_for_line(line)
                if candidate in normalized_rules
            ),
            None,
        )
        if rule:
            matches[line.canonical_sku] = rule
            matched_keys.add(rule.key)
    return matches, sorted(matched_keys), sorted(set(rules) - matched_keys)


def _log_event(event: str, **fields) -> None:
    details = " ".join(f"{key}={value}" for key, value in fields.items() if value is not None)
    message = f"{event} {details}" if details else event
    logger.info(message, extra={"box_optimizer": {"event": event, **fields}})


def _limited_order_ids(lines: list[OrderLine], max_orders: int | None) -> set[str] | None:
    if max_orders is None:
        return None
    order_ids = []
    seen = set()
    for line in lines:
        if line.order_id in seen:
            continue
        seen.add(line.order_id)
        order_ids.append(line.order_id)
        if len(order_ids) >= max_orders:
            break
    return set(order_ids)


def _limit_intake(intake: IntakeResult, max_orders: int | None) -> IntakeResult:
    allowed_order_ids = _limited_order_ids(intake.order_lines, max_orders)
    if allowed_order_ids is None:
        return intake

    order_lines = [
        line for line in intake.order_lines if line.order_id in allowed_order_ids
    ]
    matched_order_lines = [
        line for line in intake.matched_order_lines if line.order_id in allowed_order_ids
    ]
    unmatched_skus = [
        record
        for record in intake.unmatched_skus
        if record.order_line.order_id in allowed_order_ids
    ]
    debug = dict(intake.debug)
    debug["max_orders"] = max_orders
    debug["inspected_order_count"] = len(allowed_order_ids)
    debug["order_lines_created"] = len(order_lines)
    debug["matched"] = len(matched_order_lines)
    debug["unmatched"] = len(unmatched_skus)
    return IntakeResult(
        sku_items=intake.sku_items,
        order_lines=order_lines,
        matched_order_lines=matched_order_lines,
        unmatched_skus=unmatched_skus,
        column_mappings=intake.column_mappings,
        debug=debug,
    )


def _detected_columns(column_mappings: list[dict], key: str) -> list[str]:
    columns = []
    for mapping in column_mappings:
        value = mapping.get(key, "")
        if value:
            columns.extend(
                column.strip()
                for column in str(value).split(" | ")
                if column.strip()
            )
    return sorted(set(columns))


def inspect_workbook(
    sku_master_path: str,
    orders_path: str,
    config: dict | None = None,
) -> dict:
    """Parse and match intake files without packing or writing an XLSX file."""
    started = time.perf_counter()
    cfg = _config(config)

    _log_event("sku_parsing_started")
    _log_event("order_parsing_started")
    intake = read_intake(sku_master_path, orders_path)
    intake = _limit_intake(intake, cfg.get("max_orders"))
    _log_event(
        "matching_finished",
        sku_items=len(intake.sku_items),
        order_lines=len(intake.order_lines),
        matched=len(intake.matched_order_lines),
        unmatched=len(intake.unmatched_skus),
    )

    sheets = sorted(
        {
            mapping.get("sheet", "")
            for mapping in intake.column_mappings
            if mapping.get("sheet")
        }
    )
    sku_columns = _detected_columns(intake.column_mappings, "detected dimension column")
    weight_columns = _detected_columns(intake.column_mappings, "detected weight column")
    metadata_columns = _detected_columns(
        intake.column_mappings,
        "detected metadata columns",
    )
    product_columns = _detected_columns(
        intake.column_mappings,
        "detected product quantity columns",
    )
    warnings = _diagnostic_warnings(intake.debug)
    _rule_matches, matched_rule_keys, unmatched_rule_keys = _match_sku_rules(
        cfg,
        intake.sku_items,
        intake.order_lines,
    )

    return {
        "sku_items": len(intake.sku_items),
        "order_rows": intake.debug.get("order_rows_read", 0),
        "wide_product_columns": intake.debug.get("wide_product_columns_detected", 0),
        "order_lines": len(intake.order_lines),
        "matched": len(intake.matched_order_lines),
        "unmatched": len(intake.unmatched_skus),
        "sheets_read": sheets,
        "detected_sku_columns": sku_columns + weight_columns,
        "detected_order_columns": metadata_columns,
        "detected_product_quantity_columns": product_columns,
        "matched_rule_keys": matched_rule_keys,
        "unmatched_rule_keys": unmatched_rule_keys,
        "warnings": warnings,
        "elapsed_seconds": round(time.perf_counter() - started, 3),
    }


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


def _line_dimensions(sku_item: SKUItem) -> Dimensions:
    return Dimensions(
        length=sku_item.length_cm,
        width=sku_item.width_cm,
        height=sku_item.height_cm,
    )


def _line_can_bundle(rule: SKUCampaignRule | None) -> bool:
    if not rule:
        return True
    return not (
        rule.no_padding
        or rule.prepacked
        or rule.ships_alone
        or not rule.can_mix_with_other_items
        or rule.forced_box_cm
        or rule.must_stay_flat
        or rule.allowed_orientations
        or rule.extra_padding_cm
        or rule.wrap_around_largest_item
        or rule.compressible
    )


def _similar_footprint(left: Dimensions, right: Dimensions, tolerance_cm: float) -> bool:
    return (
        abs(left.length - right.length) <= tolerance_cm
        and abs(left.width - right.width) <= tolerance_cm
    )


def _bundle_label(members: list[tuple[OrderLine, SKUItem]]) -> str:
    counts = Counter()
    for line, _sku_item in members:
        counts[line.canonical_sku] += line.quantity
    return "BUNDLE[" + " | ".join(f"{sku} x{qty}" for sku, qty in sorted(counts.items())) + "]"


def _fits_carton_cap(dimensions: Dimensions) -> bool:
    return any(
        length <= MAX_CARTON_DIMENSIONS.length
        and width <= MAX_CARTON_DIMENSIONS.width
        and height <= MAX_CARTON_DIMENSIONS.height
        for length, width, height in set(permutations([dimensions.length, dimensions.width, dimensions.height]))
    )


def _bundled_packed_item(members: list[tuple[OrderLine, SKUItem]]) -> PackedItem | None:
    label = _bundle_label(members)
    dimensions_by_unit = []
    total_weight = 0.0
    for line, sku_item in members:
        dimensions = _line_dimensions(sku_item)
        dimensions_by_unit.extend([dimensions] * line.quantity)
        total_weight += sku_item.weight_kg * line.quantity
    bundle_dimensions = Dimensions(
        length=max(dimensions.length for dimensions in dimensions_by_unit),
        width=max(dimensions.width for dimensions in dimensions_by_unit),
        height=sum(dimensions.height for dimensions in dimensions_by_unit),
    )
    padded_dimensions = add_padding(bundle_dimensions)
    if not _fits_carton_cap(padded_dimensions):
        return None
    return PackedItem(
        canonical_sku=label,
        quantity=1,
        unpadded_dimensions=bundle_dimensions,
        padded_dimensions=padded_dimensions,
        weight_kg=total_weight,
        raw_sku=label,
        product_name=label,
        rule_applied="similar-footprint bundle padded once",
    )


def _largest_other_item_footprint(
    *,
    current_index: int,
    lines: list[OrderLine],
    sku_lookup: dict[str, SKUItem],
) -> Dimensions | None:
    candidates = [
        _line_dimensions(sku_lookup[line.canonical_sku])
        for index, line in enumerate(lines)
        if index != current_index
    ]
    if not candidates:
        return None
    return max(candidates, key=lambda dimensions: (dimensions.length * dimensions.width, volume(dimensions)))


def _wrapped_dimensions(
    *,
    current_index: int,
    sku_item: SKUItem,
    rule: SKUCampaignRule | None,
    lines: list[OrderLine],
    sku_lookup: dict[str, SKUItem],
) -> Dimensions:
    base = _line_dimensions(sku_item)
    if not rule or not rule.wrap_around_largest_item:
        return base
    footprint = _largest_other_item_footprint(
        current_index=current_index,
        lines=lines,
        sku_lookup=sku_lookup,
    )
    length = footprint.length if footprint else base.length
    width = footprint.width if footprint else base.width
    return Dimensions(length=length, width=width, height=float(rule.wrapped_height_cm))


def _clamped_ratio(value: float, default: float, minimum: float = 0.1) -> float:
    try:
        ratio = float(value)
    except (TypeError, ValueError):
        ratio = default
    return max(minimum, min(1.0, ratio))


def _compressible_dimensions(dimensions: Dimensions, rule: SKUCampaignRule) -> Dimensions:
    height_ratio = _clamped_ratio(rule.compressed_height_ratio, 0.6)
    volume_ratio = _clamped_ratio(rule.compressed_volume_ratio, 0.75)
    compressed_height = max(1.0, dimensions.height * height_ratio)
    footprint_ratio = 1.0
    if dimensions.length > 0 and dimensions.width > 0 and compressed_height > 0:
        original_volume = dimensions.length * dimensions.width * dimensions.height
        target_volume = original_volume * volume_ratio
        footprint_area = dimensions.length * dimensions.width
        footprint_ratio = math.sqrt(max(target_volume / (footprint_area * compressed_height), 0.0))
        footprint_ratio = max(0.85, min(1.0, footprint_ratio))
    return Dimensions(
        length=max(1.0, dimensions.length * footprint_ratio),
        width=max(1.0, dimensions.width * footprint_ratio),
        height=compressed_height,
    )


def _single_packed_item(
    line: OrderLine,
    sku_item: SKUItem,
    rule: SKUCampaignRule | None,
    dimensions: Dimensions | None = None,
) -> PackedItem:
    unpadded = dimensions or _line_dimensions(sku_item)
    if rule and rule.compressible:
        padded = _compressible_dimensions(unpadded, rule)
    elif rule and (rule.no_padding or rule.prepacked or rule.wrap_around_largest_item):
        padded = unpadded
    else:
        padded = add_padding(unpadded)
    if rule and rule.extra_padding_cm:
        padded = Dimensions(
            length=padded.length + rule.extra_padding_cm.length,
            width=padded.width + rule.extra_padding_cm.width,
            height=padded.height + rule.extra_padding_cm.height,
        )
    rule_applied = ""
    if rule:
        rule_applied = rule.key
        if rule.wrap_around_largest_item:
            rule_applied = f"{rule.key} wrap around largest item"
        if rule.compressible:
            rule_applied = f"{rule_applied} compressible"
    return PackedItem(
        canonical_sku=line.canonical_sku,
        quantity=line.quantity,
        unpadded_dimensions=unpadded,
        padded_dimensions=padded,
        weight_kg=sku_item.weight_kg,
        raw_sku=sku_item.raw_sku,
        product_name=sku_item.product_name,
        rule_key=rule.key if rule else None,
        rule_applied=rule_applied,
        box_type=rule.box_type if rule else None,
        warning_note=rule.warning_note if rule else "",
        exclude_from_standardization=rule.exclude_from_standardization if rule else False,
        allow_rotation=rule.allow_rotation if rule else True,
        must_stay_flat=rule.must_stay_flat if rule else False,
        allowed_orientations=rule.allowed_orientations if rule else None,
    )


def _packed_items_for_order(
    lines: list[OrderLine],
    sku_lookup: dict[str, SKUItem],
    sku_rules: dict[str, SKUCampaignRule] | None = None,
    bundle_footprint_tolerance_cm: float = 5,
) -> list[PackedItem]:
    items = []
    bundle_groups: list[list[tuple[OrderLine, SKUItem]]] = []
    rules = sku_rules or {}

    for index, line in enumerate(lines):
        sku_item = sku_lookup[line.canonical_sku]
        rule = rules.get(line.canonical_sku)
        dimensions = _wrapped_dimensions(
            current_index=index,
            sku_item=sku_item,
            rule=rule,
            lines=lines,
            sku_lookup=sku_lookup,
        )
        if not _line_can_bundle(rule):
            items.append(_single_packed_item(line, sku_item, rule, dimensions))
            continue

        for group in bundle_groups:
            group_dimensions = [_line_dimensions(member_sku) for _member_line, member_sku in group]
            footprint = Dimensions(
                length=max(item_dimensions.length for item_dimensions in group_dimensions),
                width=max(item_dimensions.width for item_dimensions in group_dimensions),
                height=0,
            )
            if _similar_footprint(footprint, dimensions, bundle_footprint_tolerance_cm):
                group.append((line, sku_item))
                break
        else:
            bundle_groups.append([(line, sku_item)])

    for group in bundle_groups:
        if len(group) > 1 or group[0][0].quantity > 1:
            bundled = _bundled_packed_item(group)
            if bundled is not None:
                items.append(bundled)
                continue
        for line, sku_item in group:
            items.append(_single_packed_item(line, sku_item, rules.get(line.canonical_sku)))
    return items


def _normalize_country(country: str | None) -> str:
    raw = str(country or "").strip()
    if not raw:
        return ""
    upper = raw.upper().replace(".", "")
    if upper in _COUNTRY_NAME_BY_CODE:
        return _COUNTRY_NAME_BY_CODE[upper]
    key = _normalize_lookup_key(raw)
    alias_code = _COUNTRY_ALIASES.get(key)
    if alias_code:
        return _COUNTRY_NAME_BY_CODE[alias_code]
    for name in _COUNTRY_NAME_BY_CODE.values():
        if _normalize_lookup_key(name) == key:
            return name
    return raw


def _state_abbreviation(country: str | None, state_province: str | None) -> str:
    if _normalize_country(country) != "United States" or not state_province:
        return ""
    normalized = str(state_province or "").strip()
    upper = normalized.upper().replace(".", "")
    if upper in _US_STATE_CODES:
        return upper
    return _US_STATE_ABBREVIATIONS.get(_normalize_lookup_key(normalized), "")


def _compact_box_type(box_type: str | None, dimensions: Dimensions | None = None) -> str:
    text = str(box_type or "").strip()
    if not text:
        if dimensions is None:
            return "Custom"
        return f"Custom {int(dimensions.length)}x{int(dimensions.width)}x{int(dimensions.height)}"
    vendor_match = re.fullmatch(r"Vendor Box\s+(.+)", text, flags=re.IGNORECASE)
    if vendor_match:
        return f"VB {vendor_match.group(1).strip()}"
    if re.fullmatch(r"(?:Box Type|Custom Box)\s+\d+", text, flags=re.IGNORECASE) or text == "NO-VENDOR-BOX":
        if dimensions is None:
            return text
        return f"Custom {int(dimensions.length)}x{int(dimensions.width)}x{int(dimensions.height)}"
    return text


def _per_box_chargeable_weight_summary(box_rows: list[dict]) -> str:
    return "; ".join(
        f"{row['Box Type']}: {row['Chargeable Weight kg']} kg"
        for row in sorted(box_rows, key=lambda row: row["Box Number"])
    )


def _metadata_for_order(lines: list[OrderLine]) -> dict:
    metadata = {}
    for line in lines:
        for key, value in line.metadata.items():
            if key in {"_source_file", "_source_sheet"}:
                continue
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


def _within_carton_cap(dimensions: Dimensions) -> bool:
    return (
        dimensions.length <= MAX_CARTON_DIMENSIONS.length
        and dimensions.width <= MAX_CARTON_DIMENSIONS.width
        and dimensions.height <= MAX_CARTON_DIMENSIONS.height
    )


def _capped_dimensions(dimensions: Dimensions) -> Dimensions:
    return Dimensions(
        length=min(dimensions.length, MAX_CARTON_DIMENSIONS.length),
        width=min(dimensions.width, MAX_CARTON_DIMENSIONS.width),
        height=min(dimensions.height, MAX_CARTON_DIMENSIONS.height),
    )


def _raw_carton_dimensions(split_result: SplitResult, box_index: int) -> Dimensions:
    carton = split_result.cartons[box_index].result
    return Dimensions(
        length=carton.length_cm or 0,
        width=carton.width_cm or 0,
        height=carton.height_cm or 0,
    )


def _carton_dimensions(split_result: SplitResult, box_index: int) -> Dimensions:
    if split_result.cartons[box_index].dimensions_are_final:
        return _display_dimensions(_raw_carton_dimensions(split_result, box_index), cap=False)
    return _display_dimensions(_capped_dimensions(add_final_exterior_padding(_raw_carton_dimensions(split_result, box_index))))


def _exterior_cap_violations(split_result: SplitResult) -> list[int]:
    violations = []
    for index, _carton in enumerate(split_result.cartons):
        if split_result.cartons[index].dimensions_are_final:
            exterior = _raw_carton_dimensions(split_result, index)
        else:
            exterior = add_final_exterior_padding(_raw_carton_dimensions(split_result, index))
        if not _within_carton_cap(exterior):
            violations.append(index)
    return violations


def _forced_box_result(items: list[PackedItem], dimensions: Dimensions) -> OptimizedCartonResult:
    result = pack_items(items, dimensions)
    if not result.success:
        return OptimizedCartonResult(
            success=False,
            length_cm=None,
            width_cm=None,
            height_cm=None,
            chargeable_weight_kg=None,
            volume_cm3=None,
            placements=result.placements,
            unplaced_items=result.unplaced_items,
        )
    total_weight_kg = sum(item.weight_kg for item in _expand_items(items))
    return OptimizedCartonResult(
        success=True,
        length_cm=dimensions.length,
        width_cm=dimensions.width,
        height_cm=dimensions.height,
        chargeable_weight_kg=max(
            packed_actual_weight_kg(total_weight_kg),
            dimensional_weight_kg(dimensions),
        ),
        volume_cm3=volume(dimensions),
        placements=result.placements,
        unplaced_items=[],
    )


def _prepacked_final_result(items: list[PackedItem]) -> OptimizedCartonResult:
    expanded = _expand_items(items)
    dimensions = Dimensions(
        length=max(item.padded_dimensions.length for item in expanded),
        width=max(item.padded_dimensions.width for item in expanded),
        height=sum(item.padded_dimensions.height for item in expanded),
    )
    total_weight_kg = sum(item.weight_kg for item in expanded)
    placements = [
        Placement(
            canonical_sku=item.canonical_sku,
            quantity=1,
            dimensions=item.padded_dimensions,
            origin=None,
            weight_kg=item.weight_kg,
        )
        for item in expanded
    ]
    return OptimizedCartonResult(
        success=True,
        length_cm=dimensions.length,
        width_cm=dimensions.width,
        height_cm=dimensions.height,
        chargeable_weight_kg=max(
            packed_actual_weight_kg(total_weight_kg),
            dimensional_weight_kg(dimensions),
        ),
        volume_cm3=volume(dimensions),
        placements=placements,
        unplaced_items=[],
    )


def _prepacked_final_cartons(items: list[PackedItem], rule: SKUCampaignRule) -> list[SplitCarton]:
    cartons = []
    for box_number, item in enumerate(_expand_items(items), start=1):
        result = _prepacked_final_result([item])
        cartons.append(
            SplitCarton(
                box_number=box_number,
                result=result,
                box_type=rule.box_type or "PREPACKED-FINAL-CARTON",
                rule_applied=rule.key,
                warning=rule.warning_note,
                dimensions_are_final=True,
            )
        )
    return cartons


def _carton_is_prepacked_final(
    carton: SplitCarton,
    sku_rules: dict[str, SKUCampaignRule],
) -> bool:
    if not carton.result.placements:
        return False
    placement_rules = [sku_rules.get(placement.canonical_sku) for placement in carton.result.placements]
    if all(_final_shipping_carton_rule(rule) for rule in placement_rules):
        return True
    return len(placement_rules) == 1 and bool(placement_rules[0] and placement_rules[0].prepacked)


def _merge_split_results(results: list[SplitResult]) -> SplitResult:
    cartons = []
    for result in results:
        for carton in result.cartons:
            cartons.append(
                SplitCarton(
                    box_number=len(cartons) + 1,
                    result=carton.result,
                    box_type=carton.box_type,
                    rule_applied=carton.rule_applied,
                    warning=carton.warning,
                    dimensions_are_final=carton.dimensions_are_final,
                )
            )
    unplaced = [
        item
        for result in results
        for item in result.unplaced_items
    ]
    return SplitResult(
        success=not unplaced and all(result.success for result in results),
        box_qty=len(cartons),
        cartons=cartons,
        unplaced_items=unplaced,
    )


def _split_rule_group_records(
    lines: list[OrderLine],
    sku_rules: dict[str, SKUCampaignRule],
    config: dict,
) -> list[RuleSplitGroup]:
    trigger_keys = set()
    for order_rule in config.get("order_rules") or []:
        if not isinstance(order_rule, dict):
            continue
        if order_rule.get("split_strategy") != "trigger_skus_ship_alone_addons_separate":
            continue
        for value in order_rule.get("trigger_skus", []):
            trigger_keys.add(normalize_sku(value))
            trigger_keys.add(_rule_key(value))

    groups: list[RuleSplitGroup] = []
    remainder = []
    for line in lines:
        rule = sku_rules.get(line.canonical_sku)
        line_triggered = bool(_rule_candidates_for_line(line) & trigger_keys)
        if line_triggered:
            groups.append(
                RuleSplitGroup(
                    lines=[line],
                    reason="explicit order rule",
                    rule_keys=(rule.key,) if rule else (),
                )
            )
        elif rule and rule.ships_alone:
            groups.append(
                RuleSplitGroup(
                    lines=[line],
                    reason="ships_alone=true",
                    rule_keys=(rule.key,),
                )
            )
        elif rule and not rule.can_mix_with_other_items:
            groups.append(
                RuleSplitGroup(
                    lines=[line],
                    reason="can_mix_with_other_items=false",
                    rule_keys=(rule.key,),
                )
            )
        else:
            remainder.append(line)
    if remainder:
        groups.append(RuleSplitGroup(lines=remainder))
    return groups or [RuleSplitGroup(lines=lines)]


def _split_rule_groups(
    lines: list[OrderLine],
    sku_rules: dict[str, SKUCampaignRule],
    config: dict,
) -> list[list[OrderLine]]:
    return [group.lines for group in _split_rule_group_records(lines, sku_rules, config)]


def _line_split_priority(
    line: OrderLine,
    sku_lookup: dict[str, SKUItem],
    sku_rules: dict[str, SKUCampaignRule],
    bundle_footprint_tolerance_cm: float,
) -> tuple[float, float, str]:
    items = _packed_items_for_order(
        [line],
        sku_lookup,
        sku_rules,
        bundle_footprint_tolerance_cm=bundle_footprint_tolerance_cm,
    )
    dim_weight = sum(dimensional_weight_kg(item.padded_dimensions) * item.quantity for item in items)
    packed_volume = sum(volume(item.padded_dimensions) * item.quantity for item in items)
    return (dim_weight, packed_volume, line.canonical_sku)


def _line_packed_volume(
    line: OrderLine,
    sku_lookup: dict[str, SKUItem],
    sku_rules: dict[str, SKUCampaignRule],
    bundle_footprint_tolerance_cm: float,
) -> float:
    items = _packed_items_for_order(
        [line],
        sku_lookup,
        sku_rules,
        bundle_footprint_tolerance_cm=bundle_footprint_tolerance_cm,
    )
    return sum(volume(item.padded_dimensions) * item.quantity for item in items)


def _line_is_low_volume_accessory(
    line: OrderLine,
    sku_lookup: dict[str, SKUItem],
    sku_rules: dict[str, SKUCampaignRule],
    bundle_footprint_tolerance_cm: float,
) -> bool:
    items = _packed_items_for_order(
        [line],
        sku_lookup,
        sku_rules,
        bundle_footprint_tolerance_cm=bundle_footprint_tolerance_cm,
    )
    if not items:
        return False
    aggregate_volume = sum(volume(item.padded_dimensions) * item.quantity for item in items)
    aggregate_dim_weight = sum(dimensional_weight_kg(item.padded_dimensions) * item.quantity for item in items)
    max_height = max(item.padded_dimensions.height for item in items)
    max_axis = max(
        max(item.padded_dimensions.length, item.padded_dimensions.width, item.padded_dimensions.height)
        for item in items
    )
    return aggregate_volume <= 2500 or (
        aggregate_volume <= 7000
        and aggregate_dim_weight <= 1.5
        and (max_height <= 8 or max_axis <= 22)
    )


def _line_has_campaign_rule(line: OrderLine, sku_rules: dict[str, SKUCampaignRule]) -> bool:
    rule_keys = set(sku_rules)
    return bool(_rule_candidates_for_line(line) & rule_keys)


def _line_blocks_candidate_split(line: OrderLine, sku_rules: dict[str, SKUCampaignRule]) -> bool:
    rule = sku_rules.get(line.canonical_sku)
    return bool(
        rule
        and (
            rule.prepacked
            or rule.ships_alone
            or not rule.can_mix_with_other_items
            or rule.forced_box_cm
        )
    )


def _line_is_repeat_retail_candidate(
    line: OrderLine,
    sku_lookup: dict[str, SKUItem],
    sku_rules: dict[str, SKUCampaignRule],
    bundle_footprint_tolerance_cm: float,
    min_repeated_units: int,
) -> bool:
    if line.quantity < min_repeated_units or _line_blocks_candidate_split(line, sku_rules):
        return False
    items = _packed_items_for_order(
        [line],
        sku_lookup,
        sku_rules,
        bundle_footprint_tolerance_cm=bundle_footprint_tolerance_cm,
    )
    if not items:
        return False
    unit_weight = max((item.weight_kg for item in items), default=0)
    unit_dim_weight = max((dimensional_weight_kg(item.padded_dimensions) for item in items), default=0)
    unit_volume = max((volume(item.padded_dimensions) for item in items), default=0)
    return unit_weight >= 0.4 or unit_dim_weight >= 0.5 or unit_volume >= 2500


def _line_is_repeat_retail_addon(
    line: OrderLine,
    sku_lookup: dict[str, SKUItem],
    sku_rules: dict[str, SKUCampaignRule],
    bundle_footprint_tolerance_cm: float,
) -> bool:
    if _line_blocks_candidate_split(line, sku_rules):
        return False
    rule = sku_rules.get(line.canonical_sku)
    if rule and rule.wrap_around_largest_item:
        return True
    sku_item = sku_lookup.get(line.canonical_sku)
    if sku_item:
        unit_dimensions = _line_dimensions(sku_item)
        unit_max_axis = max(unit_dimensions.length, unit_dimensions.width, unit_dimensions.height)
        unit_volume = volume(unit_dimensions)
        if sku_item.weight_kg <= 0.35 and unit_max_axis <= 25 and unit_volume <= 3500:
            return True
    items = _packed_items_for_order(
        [line],
        sku_lookup,
        sku_rules,
        bundle_footprint_tolerance_cm=bundle_footprint_tolerance_cm,
    )
    if not items:
        return False
    unit_weight = max((item.weight_kg for item in items), default=0)
    unit_dim_weight = max((dimensional_weight_kg(item.padded_dimensions) for item in items), default=0)
    max_height = max((item.padded_dimensions.height for item in items), default=0)
    max_axis = max(
        (
            max(item.padded_dimensions.length, item.padded_dimensions.width, item.padded_dimensions.height)
            for item in items
        ),
        default=0,
    )
    if rule and rule.no_padding and unit_weight <= 0.75:
        return True
    return _line_is_low_volume_accessory(
        line,
        sku_lookup,
        sku_rules,
        bundle_footprint_tolerance_cm,
    ) or (unit_weight <= 0.75 and unit_dim_weight <= 1.0 and (max_height <= 8 or max_axis <= 38))


def _line_with_quantity(line: OrderLine, quantity: int) -> OrderLine:
    return replace(line, quantity=quantity)


def _quantity_chunks(quantity: int, chunk_size: int) -> list[int]:
    if chunk_size <= 0 or quantity <= 0:
        return []
    full_chunks, remainder_quantity = divmod(quantity, chunk_size)
    chunks = [chunk_size for _ in range(full_chunks)]
    if remainder_quantity:
        chunks.append(remainder_quantity)
    return chunks


def _distributed_chunks(quantity: int, bucket_count: int) -> list[int]:
    if quantity <= 0 or bucket_count <= 0:
        return []
    base, remainder = divmod(quantity, bucket_count)
    return [base + (1 if index < remainder else 0) for index in range(bucket_count)]


def _quantity_split_groups(
    *,
    line: OrderLine,
    chunks: list[int],
    group_lines: list[OrderLine],
) -> list[RuleSplitGroup]:
    remaining_lines = [original for original in group_lines if original is not line]
    split_lines = [_line_with_quantity(line, quantity) for quantity in chunks]
    return [
        RuleSplitGroup(lines=[split_lines[0]]),
        *[RuleSplitGroup(lines=[split_line]) for split_line in split_lines[1:-1]],
        RuleSplitGroup(lines=[split_lines[-1], *remaining_lines]),
    ]


def _quantity_batch_groups(
    *,
    line: OrderLine,
    batch_size: int,
    group_lines: list[OrderLine],
) -> list[RuleSplitGroup]:
    if batch_size <= 0 or line.quantity <= batch_size:
        return []
    chunks = _quantity_chunks(line.quantity, batch_size)
    if len(chunks) <= 1:
        return []
    remaining_lines = [original for original in group_lines if original is not line]
    batch_lines = [_line_with_quantity(line, quantity) for quantity in chunks]
    groups = [RuleSplitGroup(lines=[batch_line]) for batch_line in batch_lines[:-1]]
    groups.append(RuleSplitGroup(lines=[batch_lines[-1], *remaining_lines]))
    return groups


def _repeat_retail_distributed_batch_groups(
    *,
    line: OrderLine,
    batch_size: int,
    addon_lines: list[OrderLine],
) -> list[RuleSplitGroup]:
    chunks = _quantity_chunks(line.quantity, batch_size)
    if len(chunks) <= 1:
        return []
    grouped_lines: list[list[OrderLine]] = [[_line_with_quantity(line, quantity)] for quantity in chunks]
    for addon in addon_lines:
        for index, quantity in enumerate(_distributed_chunks(addon.quantity, len(chunks))):
            if quantity:
                grouped_lines[index].append(_line_with_quantity(addon, quantity))
    return [RuleSplitGroup(lines=lines) for lines in grouped_lines]


def _repeat_retail_accessory_split_groups(
    *,
    line: OrderLine,
    batch_size: int,
    addon_lines: list[OrderLine],
) -> list[RuleSplitGroup]:
    chunks = _quantity_chunks(line.quantity, batch_size)
    if len(chunks) <= 1 or not addon_lines:
        return []
    return [
        *[RuleSplitGroup(lines=[_line_with_quantity(line, quantity)]) for quantity in chunks],
        RuleSplitGroup(lines=addon_lines),
    ]


def _candidate_group_signature(groups: list[RuleSplitGroup]) -> tuple:
    signature = []
    for group in groups:
        quantities = defaultdict(int)
        for line in group.lines:
            quantities[line.canonical_sku] += line.quantity
        signature.append(
            (
                group.reason,
                tuple(group.rule_keys),
                tuple((sku, quantities[sku]) for sku in sorted(quantities)),
            )
        )
    return tuple(signature)


def _replace_group_candidate(
    *,
    candidates: list[tuple[str, list[RuleSplitGroup]]],
    seen: set[tuple],
    name: str,
    groups: list[RuleSplitGroup],
    group_index: int,
    replacement_groups: list[RuleSplitGroup],
) -> None:
    candidate_groups = [
        replacement
        for index, original in enumerate(groups)
        for replacement in (replacement_groups if index == group_index else [original])
    ]
    signature = _candidate_group_signature(candidate_groups)
    if signature in seen:
        return
    seen.add(signature)
    candidates.append((name, candidate_groups))


def _chargeable_candidate_group_sets(
    groups: list[RuleSplitGroup],
    sku_lookup: dict[str, SKUItem],
    sku_rules: dict[str, SKUCampaignRule],
    bundle_footprint_tolerance_cm: float,
    allow_soft_rule_splits: bool = False,
    cfg: dict | None = None,
) -> list[tuple[str, list[RuleSplitGroup]]]:
    candidates = [("current", groups)]
    seen = {_candidate_group_signature(groups)}
    cfg = cfg or {}
    repeat_retail_enabled = bool(
        cfg.get("repeat_retail_batch_planning_enabled")
        if cfg.get("repeat_retail_batch_planning_enabled") is not None
        else DEFAULT_CONFIG["repeat_retail_batch_planning_enabled"]
    )
    repeat_retail_enabled = repeat_retail_enabled and float(
        cfg.get("max_optimization_seconds")
        if cfg.get("max_optimization_seconds") is not None
        else DEFAULT_CONFIG["max_optimization_seconds"]
    ) >= float(
        cfg.get("repeat_retail_min_optimization_seconds")
        if cfg.get("repeat_retail_min_optimization_seconds") is not None
        else DEFAULT_CONFIG["repeat_retail_min_optimization_seconds"]
    )
    repeat_retail_min_units = int(
        cfg.get("repeat_retail_min_repeated_units")
        if cfg.get("repeat_retail_min_repeated_units") is not None
        else DEFAULT_CONFIG["repeat_retail_min_repeated_units"]
    )
    repeat_retail_batch_sizes = [
        int(size)
        for size in (
            cfg.get("repeat_retail_batch_sizes")
            or DEFAULT_CONFIG["repeat_retail_batch_sizes"]
        )
        if isinstance(size, int) and size > 0
    ]
    repeat_retail_max_candidate_boxes = int(
        cfg.get("repeat_retail_max_candidate_boxes")
        if cfg.get("repeat_retail_max_candidate_boxes") is not None
        else DEFAULT_CONFIG["repeat_retail_max_candidate_boxes"]
    )
    for group_index, group in enumerate(groups):
        if group.reason:
            continue
        candidate_lines = [
            line for line in group.lines if not _line_blocks_candidate_split(line, sku_rules)
        ]
        if not candidate_lines:
            continue
        ranked_lines = sorted(
            candidate_lines,
            key=lambda line: (
                _line_split_priority(
                    line,
                    sku_lookup,
                    sku_rules,
                    bundle_footprint_tolerance_cm,
                ),
                line.quantity,
            ),
            reverse=True,
        )
        normal_splits_allowed = (
            allow_soft_rule_splits
            or not any(_line_has_campaign_rule(line, sku_rules) for line in group.lines)
        )
        if normal_splits_allowed and len(candidate_lines) > 1:
            split_line = ranked_lines[0]
            remainder = [line for line in group.lines if line is not split_line]
            if remainder:
                _replace_group_candidate(
                    candidates=candidates,
                    seen=seen,
                    name=f"split {split_line.canonical_sku}",
                    groups=groups,
                    group_index=group_index,
                    replacement_groups=[
                        RuleSplitGroup(lines=[split_line]),
                        RuleSplitGroup(lines=remainder),
                    ],
                )

            if len(ranked_lines) >= 3:
                top_two = ranked_lines[:2]
                remainder = [line for line in group.lines if line not in top_two]
                if remainder:
                    _replace_group_candidate(
                        candidates=candidates,
                        seen=seen,
                        name=f"split top DIM {top_two[0].canonical_sku}, {top_two[1].canonical_sku}",
                        groups=groups,
                        group_index=group_index,
                        replacement_groups=[
                            RuleSplitGroup(lines=[top_two[0]]),
                            RuleSplitGroup(lines=[top_two[1]]),
                            RuleSplitGroup(lines=remainder),
                        ],
                    )

        if normal_splits_allowed:
            accessories = [
                line
                for line in group.lines
                if _line_is_low_volume_accessory(
                    line,
                    sku_lookup,
                    sku_rules,
                    bundle_footprint_tolerance_cm,
                )
            ]
            non_accessories = [line for line in group.lines if line not in accessories]
            if accessories and non_accessories:
                _replace_group_candidate(
                    candidates=candidates,
                    seen=seen,
                    name="split low-volume accessories",
                    groups=groups,
                    group_index=group_index,
                    replacement_groups=[
                        RuleSplitGroup(lines=non_accessories),
                        RuleSplitGroup(lines=accessories),
                    ],
                )

            quantity_split_lines = [
                line for line in ranked_lines if line.quantity > 1
            ]
            for line in quantity_split_lines[:2]:
                one_unit = _line_with_quantity(line, 1)
                remaining_unit = _line_with_quantity(line, line.quantity - 1)
                remainder = [
                    remaining_unit if original is line else original
                    for original in group.lines
                ]
                _replace_group_candidate(
                    candidates=candidates,
                    seen=seen,
                    name=f"split one {line.canonical_sku}",
                    groups=groups,
                    group_index=group_index,
                    replacement_groups=[
                        RuleSplitGroup(lines=[one_unit]),
                        RuleSplitGroup(lines=remainder),
                    ],
                )
                if line.quantity > 2:
                    first_qty = line.quantity // 2
                    second_qty = line.quantity - first_qty
                    first_chunk = _line_with_quantity(line, first_qty)
                    second_chunk = _line_with_quantity(line, second_qty)
                    remainder_without_line = [
                        original for original in group.lines if original is not line
                    ]
                    _replace_group_candidate(
                        candidates=candidates,
                        seen=seen,
                        name=f"split quantity {line.canonical_sku}",
                        groups=groups,
                        group_index=group_index,
                        replacement_groups=[
                            RuleSplitGroup(lines=[first_chunk]),
                            RuleSplitGroup(lines=[second_chunk, *remainder_without_line]),
                        ],
                    )
                if line.quantity >= 5:
                    _replace_group_candidate(
                        candidates=candidates,
                        seen=seen,
                        name=f"split quantity 3+2 {line.canonical_sku}",
                        groups=groups,
                        group_index=group_index,
                        replacement_groups=_quantity_split_groups(
                            line=line,
                            chunks=[3, line.quantity - 3],
                            group_lines=group.lines,
                        ),
                    )
                if line.quantity >= 5:
                    remaining = line.quantity - 4
                    chunks = [2, 2] + ([remaining] if remaining else [])
                    _replace_group_candidate(
                        candidates=candidates,
                        seen=seen,
                        name=f"split quantity 2+2+1 {line.canonical_sku}",
                        groups=groups,
                        group_index=group_index,
                        replacement_groups=_quantity_split_groups(
                            line=line,
                            chunks=chunks,
                            group_lines=group.lines,
                        ),
                    )

        if repeat_retail_enabled:
            repeat_lines = [
                line
                for line in ranked_lines
                if _line_is_repeat_retail_candidate(
                    line,
                    sku_lookup,
                    sku_rules,
                    bundle_footprint_tolerance_cm,
                    repeat_retail_min_units,
                )
            ]
            for line in repeat_lines[:1]:
                addon_lines = [
                    candidate_line
                    for candidate_line in group.lines
                    if candidate_line is not line
                    and _line_is_repeat_retail_addon(
                        candidate_line,
                        sku_lookup,
                        sku_rules,
                        bundle_footprint_tolerance_cm,
                    )
                ]
                all_other_lines_are_addons = len(addon_lines) == len(
                    [candidate_line for candidate_line in group.lines if candidate_line is not line]
                )
                feasible_sizes = [
                    size
                    for size in repeat_retail_batch_sizes
                    if 1 < size < line.quantity
                    and ((line.quantity + size - 1) // size) <= repeat_retail_max_candidate_boxes
                ]
                for batch_size in sorted(set(feasible_sizes)):
                    replacement_groups = _quantity_batch_groups(
                        line=line,
                        batch_size=batch_size,
                        group_lines=group.lines,
                    )
                    if not replacement_groups:
                        continue
                    _replace_group_candidate(
                        candidates=candidates,
                        seen=seen,
                        name=f"repeat retail batch {batch_size} {line.canonical_sku}",
                        groups=groups,
                        group_index=group_index,
                        replacement_groups=replacement_groups,
                    )
                    if all_other_lines_are_addons:
                        distributed_groups = _repeat_retail_distributed_batch_groups(
                            line=line,
                            batch_size=batch_size,
                            addon_lines=addon_lines,
                        )
                        if distributed_groups:
                            _replace_group_candidate(
                                candidates=candidates,
                                seen=seen,
                                name=f"repeat retail distributed {batch_size} {line.canonical_sku}",
                                groups=groups,
                                group_index=group_index,
                                replacement_groups=distributed_groups,
                            )
                        accessory_split_groups = _repeat_retail_accessory_split_groups(
                            line=line,
                            batch_size=batch_size,
                            addon_lines=addon_lines,
                        )
                        if accessory_split_groups:
                            _replace_group_candidate(
                                candidates=candidates,
                                seen=seen,
                                name=f"repeat retail accessory split {batch_size} {line.canonical_sku}",
                                groups=groups,
                                group_index=group_index,
                                replacement_groups=accessory_split_groups,
                            )
    return candidates


def _rule_split_message(groups: list[RuleSplitGroup]) -> str:
    reasons = [group.reason for group in groups if group.reason]
    if not reasons:
        return ""
    return "Order split due to " + ", ".join(dict.fromkeys(reasons)) + "."


def _rule_split_keys(groups: list[RuleSplitGroup]) -> str:
    keys = [key for group in groups for key in group.rule_keys if key]
    return ", ".join(dict.fromkeys(keys))


def _final_shipping_carton_rule(rule: SKUCampaignRule | None) -> bool:
    return bool(rule and rule.prepacked and (rule.ships_alone or not rule.can_mix_with_other_items))


def _dimensions_cache_tuple(dimensions: Dimensions | None) -> tuple[float, float, float] | None:
    if dimensions is None:
        return None
    return (float(dimensions.length), float(dimensions.width), float(dimensions.height))


def _rule_cache_signature(rule: SKUCampaignRule | None) -> dict:
    if rule is None:
        return {}
    return {
        "key": rule.key,
        "no_padding": rule.no_padding,
        "prepacked": rule.prepacked,
        "ships_alone": rule.ships_alone,
        "can_mix_with_other_items": rule.can_mix_with_other_items,
        "forced_box_cm": _dimensions_cache_tuple(rule.forced_box_cm),
        "must_stay_flat": rule.must_stay_flat,
        "allow_rotation": rule.allow_rotation,
        "allowed_orientations": tuple(_dimensions_cache_tuple(dimensions) for dimensions in (rule.allowed_orientations or ())),
        "extra_padding_cm": _dimensions_cache_tuple(rule.extra_padding_cm),
        "wrap_around_largest_item": rule.wrap_around_largest_item,
        "wrapped_height_cm": float(rule.wrapped_height_cm),
        "compressible": rule.compressible,
        "compressed_height_ratio": float(rule.compressed_height_ratio),
        "compressed_volume_ratio": float(rule.compressed_volume_ratio),
        "exclude_from_standardization": rule.exclude_from_standardization,
        "box_type": rule.box_type,
        "warning_note": rule.warning_note,
    }


def _packed_item_cache_signature(item: PackedItem) -> dict:
    return {
        "canonical_sku": item.canonical_sku,
        "quantity": item.quantity,
        "unpadded_dimensions": _dimensions_cache_tuple(item.unpadded_dimensions),
        "padded_dimensions": _dimensions_cache_tuple(item.padded_dimensions),
        "weight_kg": float(item.weight_kg),
        "rule_key": item.rule_key,
        "rule_applied": item.rule_applied,
        "box_type": item.box_type,
        "warning_note": item.warning_note,
        "exclude_from_standardization": item.exclude_from_standardization,
        "allow_rotation": item.allow_rotation,
        "must_stay_flat": item.must_stay_flat,
        "allowed_orientations": tuple(_dimensions_cache_tuple(dimensions) for dimensions in (item.allowed_orientations or ())),
    }


def _group_cache_signature(group: RuleSplitGroup) -> dict:
    quantities = defaultdict(int)
    for line in group.lines:
        quantities[line.canonical_sku] += line.quantity
    return {
        "reason": group.reason,
        "rule_keys": group.rule_keys,
        "skus": tuple((sku, quantities[sku]) for sku in sorted(quantities)),
    }


def _config_cache_signature(cfg: dict) -> dict:
    return {
        "packing_mode": cfg.get("packing_mode"),
        "max_optimization_seconds": cfg.get("max_optimization_seconds"),
        "balanced_max_items_for_deep_search": cfg.get("balanced_max_items_for_deep_search"),
        "balanced_max_item_quantity_for_recombine": cfg.get("balanced_max_item_quantity_for_recombine"),
        "balanced_min_remaining_seconds": cfg.get("balanced_min_remaining_seconds"),
        "bundle_footprint_tolerance_cm": cfg.get("bundle_footprint_tolerance_cm"),
        "chargeable_weight_split_savings_threshold_kg": cfg.get("chargeable_weight_split_savings_threshold_kg"),
        "chargeable_weight_split_savings_threshold_pct": cfg.get("chargeable_weight_split_savings_threshold_pct"),
        "chargeable_weight_split_two_extra_box_threshold_kg": cfg.get("chargeable_weight_split_two_extra_box_threshold_kg"),
        "max_extra_boxes_per_order": cfg.get("max_extra_boxes_per_order"),
        "oversized_vendor_box_ids": tuple(cfg.get("oversized_vendor_box_ids") or ()),
        "oversized_vendor_box_chargeable_threshold_kg": cfg.get("oversized_vendor_box_chargeable_threshold_kg"),
        "oversized_max_extra_boxes_per_order": cfg.get("oversized_max_extra_boxes_per_order"),
        "non_preferred_extra_box_savings_threshold_kg": cfg.get("non_preferred_extra_box_savings_threshold_kg"),
        "non_preferred_extra_box_savings_threshold_pct": cfg.get("non_preferred_extra_box_savings_threshold_pct"),
        "non_preferred_two_extra_box_savings_threshold_kg": cfg.get("non_preferred_two_extra_box_savings_threshold_kg"),
        "non_preferred_two_extra_box_savings_threshold_pct": cfg.get("non_preferred_two_extra_box_savings_threshold_pct"),
        "company_protection_extra_box_guardrail": cfg.get("company_protection_extra_box_guardrail"),
        "company_protection_rate_bands": cfg.get("company_protection_rate_bands"),
        "company_protection_zone": cfg.get("company_protection_zone"),
        "company_protection_zone_markups": cfg.get("company_protection_zone_markups"),
        "company_protection_country_zones": cfg.get("company_protection_country_zones"),
        "company_protection_max_rate_weight_kg": cfg.get("company_protection_max_rate_weight_kg"),
        "company_protection_min_margin_delta": cfg.get("company_protection_min_margin_delta"),
        "repeat_retail_batch_planning_enabled": cfg.get("repeat_retail_batch_planning_enabled"),
        "repeat_retail_min_repeated_units": cfg.get("repeat_retail_min_repeated_units"),
        "repeat_retail_batch_sizes": tuple(cfg.get("repeat_retail_batch_sizes") or ()),
        "repeat_retail_max_extra_boxes_per_order": cfg.get("repeat_retail_max_extra_boxes_per_order"),
        "repeat_retail_max_candidate_boxes": cfg.get("repeat_retail_max_candidate_boxes"),
        "repeat_retail_min_optimization_seconds": cfg.get("repeat_retail_min_optimization_seconds"),
        "repeat_retail_min_savings_threshold_kg": cfg.get("repeat_retail_min_savings_threshold_kg"),
        "repeat_retail_min_savings_threshold_pct": cfg.get("repeat_retail_min_savings_threshold_pct"),
        "repeat_retail_max_margin_giveback": cfg.get("repeat_retail_max_margin_giveback"),
        "repeat_retail_min_customer_savings": cfg.get("repeat_retail_min_customer_savings"),
        "vendor_box_fit_mode": cfg.get("vendor_box_fit_mode"),
        "vendor_box_fit_tolerance_cm": cfg.get("vendor_box_fit_tolerance_cm"),
        "vendor_box_fit_tolerance_max_cm": cfg.get("vendor_box_fit_tolerance_max_cm"),
        "vendor_box_fit_tolerance_guardrail": cfg.get("vendor_box_fit_tolerance_guardrail"),
        "vendor_box_fit_tolerance_max_chargeable_increase_kg": cfg.get(
            "vendor_box_fit_tolerance_max_chargeable_increase_kg"
        ),
        "standardization_tolerance_cm": cfg.get("standardization_tolerance_cm"),
        "use_vendor_box_menu": cfg.get("use_vendor_box_menu"),
        "billing_band_kg": cfg.get("billing_band_kg"),
        "custom_box_min_units": cfg.get("custom_box_min_units"),
        "non_preferred_box_min_units": cfg.get("non_preferred_box_min_units"),
        "box_menu": cfg.get("box_menu"),
        "order_rules": cfg.get("order_rules"),
        "max_carton_cm": cfg.get("max_carton_cm"),
        "dimensional_divisor": cfg.get("dimensional_divisor"),
        "packing_weight_uplift": cfg.get("packing_weight_uplift"),
        "final_exterior_padding_cm": (2, 2, 2),
        "carton_cap_cm": _dimensions_cache_tuple(MAX_CARTON_DIMENSIONS),
    }



def _short_hash(value: str) -> str:
    return hashlib.sha1(value.encode("utf-8")).hexdigest()[:10]

def _packing_cache_key(
    combo: str,
    lines: list[OrderLine],
    items: list[PackedItem],
    groups: list[RuleSplitGroup],
    sku_rules: dict[str, SKUCampaignRule],
    cfg: dict,
) -> str:
    relevant_skus = sorted({line.canonical_sku for line in lines})
    payload = {
        "sku_combination": combo,
        "items": sorted(
            (_packed_item_cache_signature(item) for item in items),
            key=lambda item: (item["canonical_sku"], item["quantity"]),
        ),
        "groups": sorted(
            (_group_cache_signature(group) for group in groups),
            key=lambda group: (group["reason"], group["skus"], group["rule_keys"]),
        ),
        "rules": {sku: _rule_cache_signature(sku_rules.get(sku)) for sku in relevant_skus},
        "company_protection_zone": _company_protection_zone(lines, cfg) if _configured_rate_bands(cfg) else None,
        "config": _config_cache_signature(cfg),
    }
    return json.dumps(payload, sort_keys=True, separators=(",", ":"), default=str)


def _clone_split_result(split_result: SplitResult) -> SplitResult:
    return deepcopy(split_result)


def _materialize_cached_warnings(
    warnings: tuple[WorkflowWarning, ...],
    order_id: str,
    combo: str,
) -> list[WorkflowWarning]:
    return [
        replace(
            warning,
            order_id=order_id if warning.order_id else warning.order_id,
            sku_breakdown=combo if warning.sku_breakdown else warning.sku_breakdown,
        )
        for warning in warnings
    ]


def _pack_group(
    group: list[OrderLine],
    items: list[PackedItem],
    sku_rules: dict[str, SKUCampaignRule],
    packing_mode: str,
    cfg: dict | None = None,
    remaining_budget_seconds: float | None = None,
) -> tuple[SplitResult, list[WorkflowWarning]]:
    cfg = cfg or DEFAULT_CONFIG
    warnings = []
    group_rule = sku_rules.get(group[0].canonical_sku) if len(group) == 1 else None
    if group_rule and _final_shipping_carton_rule(group_rule) and not group_rule.forced_box_cm:
        cartons = _prepacked_final_cartons(items, group_rule)
        return (
            SplitResult(
                True,
                len(cartons),
                cartons,
                [],
            ),
            warnings,
        )
    if group_rule and group_rule.forced_box_cm:
        if not _within_carton_cap(group_rule.forced_box_cm):
            warnings.append(
                WorkflowWarning(
                    order_id=group[0].order_id,
                    sku=group[0].canonical_sku,
                    stage="packing",
                    error_type="ForcedBoxExceedsCap",
                    message="forced_box_cm exceeds the 74 x 37 x 44 cm carton cap.",
                    rule_applied=group_rule.key,
                    sku_breakdown=_sku_breakdown(group),
                )
            )
            return SplitResult(False, 0, [], items), warnings
        result = _forced_box_result(items, group_rule.forced_box_cm)
        if not result.success:
            warnings.append(
                WorkflowWarning(
                    order_id=group[0].order_id,
                    sku=group[0].canonical_sku,
                    stage="packing",
                    error_type="ForcedBoxTooSmall",
                    message="forced_box_cm is smaller than the item or group dimensions.",
                    rule_applied=group_rule.key,
                    sku_breakdown=_sku_breakdown(group),
                )
            )
            return SplitResult(False, 0, [], result.unplaced_items), warnings
        return (
            SplitResult(
                True,
                1,
                [
                    SplitCarton(
                        box_number=1,
                        result=result,
                        box_type=group_rule.box_type or "FORCED-BOX",
                        rule_applied=group_rule.key,
                        warning=group_rule.warning_note,
                        dimensions_are_final=True,
                    )
                ],
                [],
            ),
            warnings,
        )

    group_rules = [
        sku_rules[line.canonical_sku]
        for line in group
        if line.canonical_sku in sku_rules
    ]
    result = split_order_into_cartons(
        items,
        packing_mode=packing_mode,
        force_simple_split=False,
        balanced_max_items_for_deep_search=int(cfg.get("balanced_max_items_for_deep_search", 18)),
        balanced_max_item_quantity_for_recombine=int(cfg.get("balanced_max_item_quantity_for_recombine", 10)),
        balanced_min_remaining_seconds=float(cfg.get("balanced_min_remaining_seconds", 3)),
        remaining_budget_seconds=remaining_budget_seconds,
    )
    if result.success and group_rules:
        box_type = " + ".join(
            dict.fromkeys(rule.box_type for rule in group_rules if rule.box_type)
        ) or None
        rule_applied = ", ".join(dict.fromkeys(rule.key for rule in group_rules))
        warning = " | ".join(dict.fromkeys(rule.warning_note for rule in group_rules if rule.warning_note))
        result = SplitResult(
            True,
            result.box_qty,
            [
                SplitCarton(
                    box_number=carton.box_number,
                    result=carton.result,
                    box_type=box_type or carton.box_type,
                    rule_applied=rule_applied,
                    warning=warning or carton.warning,
                    dimensions_are_final=carton.dimensions_are_final
                    or _carton_is_prepacked_final(carton, sku_rules),
                )
                for carton in result.cartons
            ],
            [],
        )
    return result, warnings


def _pack_group_records(
    *,
    groups: list[RuleSplitGroup],
    sku_items: dict[str, SKUItem],
    sku_rules: dict[str, SKUCampaignRule],
    packing_mode: str,
    cfg: dict,
    remaining_budget_seconds: float | None,
) -> tuple[SplitResult, list[WorkflowWarning]]:
    group_results = []
    group_warnings = []
    for group_record in groups:
        group = group_record.lines
        group_items = _packed_items_for_order(
            group,
            sku_items,
            sku_rules,
            bundle_footprint_tolerance_cm=float(cfg.get("bundle_footprint_tolerance_cm", 5)),
        )
        group_result, warnings = _pack_group(
            group,
            group_items,
            sku_rules,
            packing_mode,
            cfg=cfg,
            remaining_budget_seconds=remaining_budget_seconds,
        )
        group_results.append(group_result)
        group_warnings.extend(warnings)
    return _merge_split_results(group_results), group_warnings


def _select_chargeable_weight_plan(
    *,
    context: PackingOrderContext,
    sku_items: dict[str, SKUItem],
    sku_rules: dict[str, SKUCampaignRule],
    packing_mode: str,
    cfg: dict,
    remaining_budget_seconds: float | None,
) -> tuple[SplitResult, list[WorkflowWarning]]:
    if not cfg.get("use_vendor_box_menu", True):
        return _pack_group_records(
            groups=context.groups,
            sku_items=sku_items,
            sku_rules=sku_rules,
            packing_mode=packing_mode,
            cfg=cfg,
            remaining_budget_seconds=remaining_budget_seconds,
        )
    threshold_kg = float(
        cfg.get("chargeable_weight_split_savings_threshold_kg")
        if cfg.get("chargeable_weight_split_savings_threshold_kg") is not None
        else DEFAULT_CONFIG["chargeable_weight_split_savings_threshold_kg"]
    )
    threshold_pct = float(
        cfg.get("chargeable_weight_split_savings_threshold_pct")
        if cfg.get("chargeable_weight_split_savings_threshold_pct") is not None
        else DEFAULT_CONFIG["chargeable_weight_split_savings_threshold_pct"]
    )
    two_extra_box_threshold_kg = float(
        cfg.get("chargeable_weight_split_two_extra_box_threshold_kg")
        if cfg.get("chargeable_weight_split_two_extra_box_threshold_kg") is not None
        else DEFAULT_CONFIG["chargeable_weight_split_two_extra_box_threshold_kg"]
    )
    max_extra_boxes = int(
        cfg.get("max_extra_boxes_per_order")
        if cfg.get("max_extra_boxes_per_order") is not None
        else DEFAULT_CONFIG["max_extra_boxes_per_order"]
    )
    oversized_max_extra_boxes = int(
        cfg.get("oversized_max_extra_boxes_per_order")
        if cfg.get("oversized_max_extra_boxes_per_order") is not None
        else DEFAULT_CONFIG["oversized_max_extra_boxes_per_order"]
    )
    non_preferred_extra_threshold_kg = float(
        cfg.get("non_preferred_extra_box_savings_threshold_kg")
        if cfg.get("non_preferred_extra_box_savings_threshold_kg") is not None
        else DEFAULT_CONFIG["non_preferred_extra_box_savings_threshold_kg"]
    )
    non_preferred_extra_threshold_pct = float(
        cfg.get("non_preferred_extra_box_savings_threshold_pct")
        if cfg.get("non_preferred_extra_box_savings_threshold_pct") is not None
        else DEFAULT_CONFIG["non_preferred_extra_box_savings_threshold_pct"]
    )
    non_preferred_two_extra_threshold_kg = float(
        cfg.get("non_preferred_two_extra_box_savings_threshold_kg")
        if cfg.get("non_preferred_two_extra_box_savings_threshold_kg") is not None
        else DEFAULT_CONFIG["non_preferred_two_extra_box_savings_threshold_kg"]
    )
    non_preferred_two_extra_threshold_pct = float(
        cfg.get("non_preferred_two_extra_box_savings_threshold_pct")
        if cfg.get("non_preferred_two_extra_box_savings_threshold_pct") is not None
        else DEFAULT_CONFIG["non_preferred_two_extra_box_savings_threshold_pct"]
    )
    candidate_groups = _chargeable_candidate_group_sets(
        context.groups,
        sku_items,
        sku_rules,
        float(cfg.get("bundle_footprint_tolerance_cm", 5)),
        cfg=cfg,
    )
    scored_candidates = []
    scored_signatures = set()
    baseline_score = None
    baseline_split_result = None
    baseline_warnings: list[WorkflowWarning] = []
    def score_candidate_groups(
        candidate_name: str,
        groups: list[RuleSplitGroup],
        candidate_index: int,
    ) -> int:
        nonlocal baseline_score, baseline_split_result, baseline_warnings
        signature = _candidate_group_signature(groups)
        if signature in scored_signatures:
            return candidate_index + 1
        scored_signatures.add(signature)
        try:
            split_result, candidate_warnings = _pack_group_records(
                groups=groups,
                sku_items=sku_items,
                sku_rules=sku_rules,
                packing_mode=packing_mode,
                cfg=cfg,
                remaining_budget_seconds=remaining_budget_seconds,
            )
        except Exception:
            if candidate_index == 0:
                raise
            return candidate_index + 1
        if not split_result.success:
            if candidate_index == 0:
                return split_result, candidate_warnings
            return candidate_index + 1
        try:
            score = _score_assigned_split_result(
                order_id=context.order_id,
                combo=context.combo,
                split_result=split_result,
                cfg=cfg,
                sku_rules=sku_rules,
            )
        except Exception:
            if candidate_index == 0:
                raise
            return candidate_index + 1
        if candidate_index == 0:
            baseline_score = score
            baseline_split_result = split_result
            baseline_warnings = candidate_warnings
        scored_candidates.append((candidate_index, candidate_name, groups, split_result, candidate_warnings, score))
        return candidate_index + 1

    next_candidate_index = 0
    for candidate_name, groups in candidate_groups:
        scored = score_candidate_groups(candidate_name, groups, next_candidate_index)
        if isinstance(scored, tuple):
            return scored
        next_candidate_index = scored

    if baseline_score is None or baseline_split_result is None:
        return SplitResult(False, 0, [], context.items), []
    if baseline_score.oversized_box_count:
        soft_candidate_groups = _chargeable_candidate_group_sets(
            context.groups,
            sku_items,
            sku_rules,
            float(cfg.get("bundle_footprint_tolerance_cm", 5)),
            allow_soft_rule_splits=True,
            cfg=cfg,
        )
        for candidate_name, groups in soft_candidate_groups:
            next_candidate_index = score_candidate_groups(
                f"oversized {candidate_name}" if candidate_name != "current" else candidate_name,
                groups,
                next_candidate_index,
            )

    best = scored_candidates[0]
    for candidate in scored_candidates[1:]:
        candidate_index, _candidate_name, _groups, _split_result, _candidate_warnings, score = candidate
        if not _candidate_beats_baseline(
            score,
            baseline_score,
            candidate_index,
            threshold_kg,
            threshold_pct,
            two_extra_box_threshold_kg,
            max_extra_boxes,
            oversized_max_extra_boxes,
            non_preferred_extra_threshold_kg,
            non_preferred_extra_threshold_pct,
            non_preferred_two_extra_threshold_kg,
            non_preferred_two_extra_threshold_pct,
            context.lines,
            cfg,
            _candidate_name,
        ):
            continue
        if _candidate_score_tuple(score, candidate_index) < _candidate_score_tuple(best[5], best[0]):
            best = candidate

    best_index, best_name, _best_groups, best_split_result, best_warnings, best_score = best
    if best_index == 0:
        return baseline_split_result, baseline_warnings

    savings = baseline_score.total_chargeable_weight_kg - best_score.total_chargeable_weight_kg
    best_warnings = list(best_warnings)
    best_warnings.append(
        WorkflowWarning(
            order_id=context.order_id,
            stage="packing",
            error_type=(
                "OversizedVendorBoxPlanSelected"
                if baseline_score.oversized_box_count > best_score.oversized_box_count
                else "ChargeableWeightPlanSelected"
            ),
            message=_candidate_plan_selection_message(best_score, baseline_score, savings),
            sku_breakdown=context.combo,
            rule_applied=best_name,
        )
    )
    return best_split_result, best_warnings


def _box_word(count: int) -> str:
    return "box" if count == 1 else "boxes"


def _candidate_plan_selection_message(
    best_score: CandidatePlanScore,
    baseline_score: CandidatePlanScore,
    savings_kg: float,
) -> str:
    savings = _format_weight_display(savings_kg)
    if best_score.box_qty == baseline_score.box_qty:
        return (
            f"Selected alternate {best_score.box_qty}-box layout; "
            f"saved {savings} kg chargeable weight."
        )
    if best_score.box_qty < baseline_score.box_qty:
        box_savings = baseline_score.box_qty - best_score.box_qty
        return (
            f"Selected fewer-box plan; saved {box_savings} {_box_word(box_savings)} "
            f"and {savings} kg chargeable weight."
        )
    return (
        f"Selected {best_score.box_qty}-box plan over "
        f"{baseline_score.box_qty}-box plan; "
        f"saved {savings} kg chargeable weight."
    )


def _carton_cap_warning_for_order(
    *,
    order_id: str,
    combo: str,
    split_result: SplitResult,
    assignments_by_key: dict[str, StandardizedBoxAssignment],
) -> WorkflowWarning | None:
    violation_indexes = _exterior_cap_violations(split_result)
    if not violation_indexes:
        return None

    material_indexes = []
    for index in violation_indexes:
        carton = split_result.cartons[index]
        assignment = assignments_by_key.get(f"{order_id}#{index + 1}")
        if carton.dimensions_are_final or carton.box_type:
            continue
        if assignment and assignment.vendor_box_id:
            continue
        material_indexes.append(index)

    if not material_indexes:
        return None

    box_numbers = ", ".join(str(index + 1) for index in material_indexes)
    return WorkflowWarning(
        order_id=order_id,
        stage="packing",
        error_type="CartonCapWarning",
        message=(
            "Final exterior padding would exceed the carton cap on "
            f"box {box_numbers}; reported carton dimensions were capped at 74 x 37 x 44 cm."
        ),
        sku_breakdown=combo,
    )


def _row_float(row: dict, key: str) -> float:
    try:
        return float(row.get(key) or 0)
    except (TypeError, ValueError):
        return 0.0


def _weight_driver_note(packed_actual_kg: float, dimensional_kg: float) -> str:
    if dimensional_kg > packed_actual_kg + 0.5:
        return "dimensional weight"
    if packed_actual_kg > dimensional_kg + 0.5:
        return "packed actual weight"
    return "mixed actual and dimensional weight"


def _split_reason_note(order_id: str, box_qty: int, warning_rows: list[WorkflowWarning]) -> str:
    order_warnings = [warning for warning in warning_rows if warning.order_id == order_id]
    reasons = []
    if any(warning.error_type == "RuleBasedOrderSplit" for warning in order_warnings):
        reasons.append("rule-based final-carton split")
    if any(
        warning.error_type in {"ChargeableWeightPlanSelected", "OversizedVendorBoxPlanSelected"}
        for warning in order_warnings
    ):
        reasons.append("optimized alternate layout")
    if box_qty > 1 and not reasons:
        reasons.append("physical fit or carton selection")
    return ", ".join(reasons) if reasons else "single selected carton layout"


def _retail_bulk_review_warning(
    *,
    order_id: str,
    lines: list[OrderLine],
    order_box_rows: list[dict],
    warning_rows: list[WorkflowWarning],
    combo: str,
) -> WorkflowWarning | None:
    if not order_box_rows:
        return None
    total_units = _total_units(lines)
    box_qty = int(_row_float(order_box_rows[0], "Box Qty"))
    if total_units < 25 and box_qty < 3:
        return None

    packed_actual_kg = sum(_row_float(row, "Packed Actual Weight kg") for row in order_box_rows)
    dimensional_kg = sum(_row_float(row, "Dimensional Weight kg (/5000)") for row in order_box_rows)
    chargeable_kg = sum(_row_float(row, "Chargeable Weight g") / 1000 for row in order_box_rows)
    driver = _weight_driver_note(packed_actual_kg, dimensional_kg)
    split_reason = _split_reason_note(order_id, box_qty, warning_rows)
    return WorkflowWarning(
        order_id=order_id,
        stage="report",
        error_type="RetailBulkReview",
        message=(
            f"Retail/bulk review: {total_units} units across {box_qty} {_box_word(box_qty)}; "
            f"chargeable weight is driven mainly by {driver} "
            f"({_format_weight_display(dimensional_kg)} kg dimensional vs "
            f"{_format_weight_display(packed_actual_kg)} kg packed actual; "
            f"{_format_weight_display(chargeable_kg)} kg chargeable). "
            f"Split reason: {split_reason}."
        ),
        sku_breakdown=combo,
    )


def _warning_row(warning: WorkflowWarning) -> dict:
    return {
        "Order ID": warning.order_id,
        "SKU": warning.sku,
        "Stage": warning.stage,
        "Error Type": warning.error_type,
        "Message": warning.message,
        "Rule Applied": warning.rule_applied,
        "SKU Breakdown": warning.sku_breakdown,
    }


def _rule_allows_vendor_box_fit_tolerance(rule: SKUCampaignRule | None) -> bool:
    if not rule:
        return False
    return (
        rule.wrap_around_largest_item
        or rule.compressible
        or bool(getattr(rule, "foldable", False))
        or bool(getattr(rule, "flexible", False))
    )


def _carton_allows_vendor_box_fit_tolerance(
    carton: SplitCarton,
    sku_rules: dict[str, SKUCampaignRule],
) -> bool:
    for placement in carton.result.placements:
        if _rule_allows_vendor_box_fit_tolerance(sku_rules.get(placement.canonical_sku)):
            return True
    return False


def _vendor_box_fit_mode(config: dict) -> str:
    mode = str(config.get("vendor_box_fit_mode") or "auto").strip().lower()
    return mode if mode in {"auto", "off", "on"} else "auto"


def _build_standardization_inputs(
    split_results: dict[str, SplitResult],
    combo_by_order: dict[str, str],
    sku_rules: dict[str, SKUCampaignRule] | None = None,
    vendor_box_fit_mode: str = "auto",
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
                    chargeable_weight_kg=max(
                        carton.result.chargeable_weight_kg or 0,
                        dimensional_weight_kg(dimensions),
                    ),
                    placements=carton.result.placements,
                    allow_vendor_box_fit_tolerance=(
                        vendor_box_fit_mode == "on"
                        or (
                            vendor_box_fit_mode == "auto"
                            and _carton_allows_vendor_box_fit_tolerance(carton, sku_rules or {})
                        )
                    ),
                )
            )
    return optimized


def _assignment_lookup(
    assignments: list[StandardizedBoxAssignment],
) -> dict[str, StandardizedBoxAssignment]:
    return {assignment.order_id: assignment for assignment in assignments}


def _standardize_split_result(
    split_results: dict[str, SplitResult],
    combo_by_order: dict[str, str],
    cfg: dict,
    sku_rules: dict[str, SKUCampaignRule] | None = None,
) -> list[StandardizedBoxAssignment]:
    try:
        vendor_fit_tolerance = float(cfg.get("vendor_box_fit_tolerance_cm", 0) or 0)
    except (TypeError, ValueError):
        vendor_fit_tolerance = 0.0
    try:
        vendor_fit_tolerance_max = float(cfg.get("vendor_box_fit_tolerance_max_cm", 2.0) or 2.0)
    except (TypeError, ValueError):
        vendor_fit_tolerance_max = 2.0
    vendor_fit_tolerance = max(0.0, min(vendor_fit_tolerance, vendor_fit_tolerance_max, 2.0))
    vendor_fit_mode = _vendor_box_fit_mode(cfg)
    if vendor_fit_mode == "off":
        vendor_fit_tolerance = 0.0
    return standardize_optimized_cartons(
        _build_standardization_inputs(split_results, combo_by_order, sku_rules or {}, vendor_fit_mode),
        tolerance_cm=cfg["standardization_tolerance_cm"],
        use_vendor_box_menu=cfg.get("use_vendor_box_menu", True),
        billing_band_kg=cfg.get("billing_band_kg", 1.0),
        custom_box_min_units=cfg.get("custom_box_min_units", 400),
        non_preferred_box_min_units=cfg.get("non_preferred_box_min_units", 100),
        vendor_box_fit_tolerance_cm=vendor_fit_tolerance,
        vendor_box_fit_tolerance_guardrail=cfg.get("vendor_box_fit_tolerance_guardrail", True),
        vendor_box_fit_tolerance_max_chargeable_increase_kg=float(
            cfg.get("vendor_box_fit_tolerance_max_chargeable_increase_kg", 1.0) or 1.0
        ),
    )


def _assigned_carton_dimensions_and_type(
    *,
    carton: SplitCarton,
    assignment: StandardizedBoxAssignment,
    optimized_dimensions: Dimensions,
) -> tuple[Dimensions, str, str]:
    raw_carton_box_type = carton.box_type or assignment.box_type
    vendor_box_id = "" if carton.box_type else (assignment.vendor_box_id or "")
    assigned_dimensions = _display_dimensions(
        Dimensions(
            optimized_dimensions.length if carton.box_type else assignment.assigned_length_cm,
            optimized_dimensions.width if carton.box_type else assignment.assigned_width_cm,
            optimized_dimensions.height if carton.box_type else assignment.assigned_height_cm,
        ),
        cap=bool(carton.box_type) and not carton.dimensions_are_final,
    )
    assigned_dimensions, cutdown_note = _vendor_height_cutdown(
        dimensions=assigned_dimensions,
        optimized_dimensions=optimized_dimensions,
        placements=carton.result.placements,
        vendor_box_id=vendor_box_id,
    )
    if (
        not cutdown_note
        and not carton.box_type
        and assignment.vendor_box_id
        and "Vendor box height cut down" in assignment.box_standardization_note
    ):
        cutdown_note = f"Vendor box height cut down to {int(assigned_dimensions.height)} cm."
    carton_box_type = _compact_box_type(raw_carton_box_type, assigned_dimensions)
    if cutdown_note:
        carton_box_type = f"{carton_box_type} cutdown"
    return assigned_dimensions, carton_box_type, cutdown_note


def _score_assigned_split_result(
    *,
    order_id: str,
    combo: str,
    split_result: SplitResult,
    cfg: dict,
    sku_rules: dict[str, SKUCampaignRule] | None = None,
) -> CandidatePlanScore:
    assignments = _standardize_split_result({order_id: split_result}, {order_id: combo}, cfg, sku_rules)
    assignments_by_key = _assignment_lookup(assignments)
    total_chargeable = 0.0
    total_volume = 0.0
    box_types = set()
    oversized_box_count = 0
    non_preferred_box_count = 0
    package_weights = []
    oversized_vendor_ids = {str(value) for value in (cfg.get("oversized_vendor_box_ids") or DEFAULT_CONFIG["oversized_vendor_box_ids"])}
    oversized_chargeable_threshold_kg = float(
        cfg.get("oversized_vendor_box_chargeable_threshold_kg")
        if cfg.get("oversized_vendor_box_chargeable_threshold_kg") is not None
        else DEFAULT_CONFIG["oversized_vendor_box_chargeable_threshold_kg"]
    )
    for index, carton in enumerate(split_result.cartons):
        assignment = assignments_by_key[f"{order_id}#{index + 1}"]
        optimized_dimensions = _carton_dimensions(split_result, index)
        assigned_dimensions, carton_box_type, _cutdown_note = _assigned_carton_dimensions_and_type(
            carton=carton,
            assignment=assignment,
            optimized_dimensions=optimized_dimensions,
        )
        box_actual_weight_kg = sum(placement.weight_kg for placement in carton.result.placements)
        box_packed_weight_kg = packed_actual_weight_kg(box_actual_weight_kg)
        box_chargeable = max(box_packed_weight_kg, dimensional_weight_kg(assigned_dimensions))
        total_chargeable += box_chargeable
        total_volume += volume(assigned_dimensions)
        box_types.add(carton_box_type)
        if assignment.vendor_box_id and assignment.vendor_box_id not in PREFERRED_VENDOR_BOX_IDS:
            non_preferred_box_count += 1
        if (assignment.vendor_box_id and assignment.vendor_box_id in oversized_vendor_ids) or box_chargeable >= oversized_chargeable_threshold_kg:
            oversized_box_count += 1
        package_weights.append(box_chargeable)
    return CandidatePlanScore(
        total_chargeable_weight_kg=total_chargeable,
        box_qty=split_result.box_qty,
        box_type_count=len(box_types),
        total_assigned_volume_cm3=total_volume,
        oversized_box_count=oversized_box_count,
        non_preferred_box_count=non_preferred_box_count,
        package_chargeable_weights_kg=tuple(package_weights),
    )


def _candidate_score_tuple(score: CandidatePlanScore, candidate_index: int) -> tuple[float, int, int, float, int]:
    return (
        round(score.total_chargeable_weight_kg, 6),
        score.box_qty,
        score.box_type_count,
        round(score.total_assigned_volume_cm3, 6),
        candidate_index,
    )


def _configured_rate_bands(cfg: dict) -> dict:
    raw = cfg.get("company_protection_rate_bands")
    return raw if isinstance(raw, dict) else {}


def _company_protection_zone(lines: list[OrderLine], cfg: dict) -> str:
    configured = cfg.get("company_protection_zone")
    if configured:
        return str(configured)
    if not lines:
        return "Zone USA"
    first_line = lines[0]
    country = _normalize_country(first_line.country)
    if country == "United States":
        state = _state_abbreviation(first_line.country, first_line.state_province)
        return "Zone 1" if state in _USA_PLUS_ZONE_STATES else "Zone USA"
    country_zone_map = cfg.get("company_protection_country_zones")
    if isinstance(country_zone_map, dict):
        return str(
            country_zone_map.get(country)
            or country_zone_map.get(str(first_line.country or "").strip())
            or country_zone_map.get("default")
            or country
        )
    return country or "Zone USA"


def _rate_band_lookup(rate_bands: dict, zone: str) -> dict[float, float]:
    zone_bands = rate_bands.get(zone)
    if zone_bands is None:
        zone_bands = rate_bands.get("default")
    if not isinstance(zone_bands, dict):
        return {}
    parsed = {}
    for weight, amount in zone_bands.items():
        try:
            parsed[float(weight)] = float(amount)
        except (TypeError, ValueError):
            continue
    return dict(sorted(parsed.items()))


def _rounded_rate_weight_kg(weight_kg: float) -> float:
    if weight_kg <= 0:
        return 0.5
    return math.ceil((weight_kg - 1e-9) * 2) / 2


def _band_rate_for_weight(weight_kg: float, bands: dict[float, float]) -> float | None:
    if not bands:
        return None
    target = _rounded_rate_weight_kg(weight_kg)
    for band_weight, amount in bands.items():
        if band_weight + 1e-9 >= target:
            return amount
    return None


def _rated_charge_for_weight(weight_kg: float, bands: dict[float, float], max_rate_weight_kg: float) -> float | None:
    if not bands:
        return None
    max_weight = max_rate_weight_kg if max_rate_weight_kg > 0 else max(bands)
    remaining = _rounded_rate_weight_kg(weight_kg)
    total = 0.0
    while remaining > max_weight + 1e-9:
        max_rate = _band_rate_for_weight(max_weight, bands)
        if max_rate is None:
            return None
        total += max_rate
        remaining = _rounded_rate_weight_kg(remaining - max_weight)
    final_rate = _band_rate_for_weight(remaining, bands)
    if final_rate is None:
        return None
    return total + final_rate


def _zone_markup(zone: str, cfg: dict) -> float:
    markups = cfg.get("company_protection_zone_markups")
    if not isinstance(markups, dict):
        markups = DEFAULT_CONFIG["company_protection_zone_markups"]
    try:
        markup = float(markups.get(zone, markups.get("default", 1.3)))
    except (TypeError, ValueError):
        markup = 1.3
    return markup if markup > 0 else 1.0


def _company_protection_client_charge_delta(
    *,
    candidate_score: CandidatePlanScore,
    baseline_score: CandidatePlanScore,
    lines: list[OrderLine],
    cfg: dict,
) -> float | None:
    rate_bands = _configured_rate_bands(cfg)
    if not rate_bands:
        return None
    zone = _company_protection_zone(lines, cfg)
    bands = _rate_band_lookup(rate_bands, zone)
    if not bands:
        return None
    max_rate_weight = float(
        cfg.get("company_protection_max_rate_weight_kg")
        if cfg.get("company_protection_max_rate_weight_kg") is not None
        else DEFAULT_CONFIG["company_protection_max_rate_weight_kg"]
    )
    baseline_charge = _rated_charge_for_weight(
        baseline_score.total_chargeable_weight_kg,
        bands,
        max_rate_weight,
    )
    candidate_charge = _rated_charge_for_weight(
        candidate_score.total_chargeable_weight_kg,
        bands,
        max_rate_weight,
    )
    if baseline_charge is None or candidate_charge is None:
        return None
    return baseline_charge - candidate_charge


def _company_protection_margin_delta(
    *,
    candidate_score: CandidatePlanScore,
    baseline_score: CandidatePlanScore,
    lines: list[OrderLine],
    cfg: dict,
) -> float | None:
    if cfg.get("company_protection_extra_box_guardrail") is False:
        return None
    rate_bands = _configured_rate_bands(cfg)
    if not rate_bands:
        return None
    zone = _company_protection_zone(lines, cfg)
    bands = _rate_band_lookup(rate_bands, zone)
    if not bands:
        return None
    max_rate_weight = float(
        cfg.get("company_protection_max_rate_weight_kg")
        if cfg.get("company_protection_max_rate_weight_kg") is not None
        else DEFAULT_CONFIG["company_protection_max_rate_weight_kg"]
    )
    markup = _zone_markup(zone, cfg)

    def margin_for(score: CandidatePlanScore) -> float | None:
        client_charge = _rated_charge_for_weight(score.total_chargeable_weight_kg, bands, max_rate_weight)
        if client_charge is None:
            return None
        package_weights = score.package_chargeable_weights_kg or (score.total_chargeable_weight_kg,)
        package_charge = 0.0
        for package_weight in package_weights:
            package_rate = _rated_charge_for_weight(package_weight, bands, max_rate_weight)
            if package_rate is None:
                return None
            package_charge += package_rate
        actual_shipping_cost = package_charge / markup
        return client_charge - actual_shipping_cost

    baseline_margin = margin_for(baseline_score)
    candidate_margin = margin_for(candidate_score)
    if baseline_margin is None or candidate_margin is None:
        return None
    return candidate_margin - baseline_margin


def _candidate_beats_baseline(
    candidate_score: CandidatePlanScore,
    baseline_score: CandidatePlanScore,
    candidate_index: int,
    threshold_kg: float,
    threshold_pct: float,
    two_extra_box_threshold_kg: float,
    max_extra_boxes: int,
    oversized_max_extra_boxes: int,
    non_preferred_extra_threshold_kg: float,
    non_preferred_extra_threshold_pct: float,
    non_preferred_two_extra_threshold_kg: float,
    non_preferred_two_extra_threshold_pct: float,
    lines: list[OrderLine],
    cfg: dict,
    candidate_name: str = "",
) -> bool:
    is_repeat_retail_candidate = candidate_name.startswith("repeat retail")
    if candidate_score.box_qty > baseline_score.box_qty:
        extra_boxes = candidate_score.box_qty - baseline_score.box_qty
        effective_max_extra_boxes = max_extra_boxes
        if baseline_score.oversized_box_count and candidate_score.oversized_box_count < baseline_score.oversized_box_count:
            effective_max_extra_boxes = max(effective_max_extra_boxes, oversized_max_extra_boxes)
        if is_repeat_retail_candidate:
            effective_max_extra_boxes = max(
                effective_max_extra_boxes,
                int(
                    cfg.get("repeat_retail_max_extra_boxes_per_order")
                    if cfg.get("repeat_retail_max_extra_boxes_per_order") is not None
                    else DEFAULT_CONFIG["repeat_retail_max_extra_boxes_per_order"]
                ),
            )
        if extra_boxes > effective_max_extra_boxes:
            return False
        savings = baseline_score.total_chargeable_weight_kg - candidate_score.total_chargeable_weight_kg
        required_savings = max(
            threshold_kg,
            baseline_score.total_chargeable_weight_kg * threshold_pct,
        )
        if is_repeat_retail_candidate:
            required_savings = max(
                required_savings,
                float(
                    cfg.get("repeat_retail_min_savings_threshold_kg")
                    if cfg.get("repeat_retail_min_savings_threshold_kg") is not None
                    else DEFAULT_CONFIG["repeat_retail_min_savings_threshold_kg"]
                ),
                baseline_score.total_chargeable_weight_kg
                * float(
                    cfg.get("repeat_retail_min_savings_threshold_pct")
                    if cfg.get("repeat_retail_min_savings_threshold_pct") is not None
                    else DEFAULT_CONFIG["repeat_retail_min_savings_threshold_pct"]
                ),
            )
        if extra_boxes >= 2:
            required_savings = max(required_savings, two_extra_box_threshold_kg)
        if candidate_score.non_preferred_box_count > baseline_score.non_preferred_box_count:
            required_savings = max(
                required_savings,
                non_preferred_extra_threshold_kg,
                baseline_score.total_chargeable_weight_kg * non_preferred_extra_threshold_pct,
            )
            if extra_boxes >= 2:
                required_savings = max(
                    required_savings,
                    non_preferred_two_extra_threshold_kg,
                    baseline_score.total_chargeable_weight_kg * non_preferred_two_extra_threshold_pct,
                )
        if savings + 1e-9 < required_savings:
            return False
        margin_delta = _company_protection_margin_delta(
            candidate_score=candidate_score,
            baseline_score=baseline_score,
            lines=lines,
            cfg=cfg,
        )
        if margin_delta is not None:
            required_margin_delta = float(
                cfg.get("company_protection_min_margin_delta")
                if cfg.get("company_protection_min_margin_delta") is not None
                else DEFAULT_CONFIG["company_protection_min_margin_delta"]
            )
            if margin_delta + 1e-9 < required_margin_delta:
                allowed_repeat_retail_giveback = False
                if is_repeat_retail_candidate:
                    max_giveback = float(
                        cfg.get("repeat_retail_max_margin_giveback")
                        if cfg.get("repeat_retail_max_margin_giveback") is not None
                        else DEFAULT_CONFIG["repeat_retail_max_margin_giveback"]
                    )
                    min_customer_savings = float(
                        cfg.get("repeat_retail_min_customer_savings")
                        if cfg.get("repeat_retail_min_customer_savings") is not None
                        else DEFAULT_CONFIG["repeat_retail_min_customer_savings"]
                    )
                    customer_savings = _company_protection_client_charge_delta(
                        candidate_score=candidate_score,
                        baseline_score=baseline_score,
                        lines=lines,
                        cfg=cfg,
                    )
                    allowed_repeat_retail_giveback = (
                        customer_savings is not None
                        and customer_savings + 1e-9 >= min_customer_savings
                        and margin_delta + max_giveback + 1e-9 >= 0
                    )
                if not allowed_repeat_retail_giveback:
                    return False
    return _candidate_score_tuple(candidate_score, candidate_index) < _candidate_score_tuple(baseline_score, 0)


def _rule_summary(sku_rules: dict[str, SKUCampaignRule]) -> list[str]:
    summaries = []
    for rule in sorted(sku_rules.values(), key=lambda item: item.key):
        flags = []
        if rule.ships_alone:
            flags.append("ship-alone")
        if rule.prepacked:
            flags.append("prepacked")
        if rule.no_padding:
            flags.append("no-padding")
        if rule.forced_box_cm:
            flags.append(
                f"fixed carton {_compact_box_type(rule.box_type, rule.forced_box_cm)} "
                f"{int(rule.forced_box_cm.length)}x{int(rule.forced_box_cm.width)}x{int(rule.forced_box_cm.height)}"
            )
        if not rule.can_mix_with_other_items:
            flags.append("no-mix")
        if rule.must_stay_flat:
            flags.append("must-stay-flat")
        if rule.extra_padding_cm:
            flags.append("extra padding")
        if rule.wrap_around_largest_item:
            flags.append(f"wrap around largest item, height {rule.wrapped_height_cm:g} cm")
        if rule.compressible:
            flags.append(
                f"compressible, height ratio {rule.compressed_height_ratio:g}, volume ratio {rule.compressed_volume_ratio:g}"
            )
        if flags:
            summaries.append(f"{rule.key}: {', '.join(flags)}")
    return summaries


def _base_box_type(box_type: str) -> str:
    match = re.match(r"^(VB\s+[^ ]+)", str(box_type or ""))
    return match.group(1) if match else str(box_type or "")


def _clean_summary_rows(
    result: dict,
    box_size_rows: list[dict],
    unmatched_rows: list[dict],
    sku_rules: dict[str, SKUCampaignRule],
) -> list[dict]:
    rows = [
        {
            "Section": "Run Summary",
            "Metric": "Orders Processed",
            "Value": result["orders_processed"],
            "Detail": "",
            "Box Not Available - Substituted Up To VB Box X": "",
        },
        {"Section": "Run Summary", "Metric": "Boxes Created", "Value": result["boxes_created"], "Detail": "", "Box Not Available - Substituted Up To VB Box X": ""},
        {"Section": "Run Summary", "Metric": "Box Types", "Value": result["box_types"], "Detail": "", "Box Not Available - Substituted Up To VB Box X": ""},
        {"Section": "Run Summary", "Metric": "Unmatched SKUs", "Value": result["unmatched_skus"], "Detail": "", "Box Not Available - Substituted Up To VB Box X": ""},
        {"Section": "Cost Placeholder", "Metric": "Total Chargeable Cost", "Value": "Pending rate integration", "Detail": "", "Box Not Available - Substituted Up To VB Box X": ""},
        {"Section": "Cost Placeholder", "Metric": "Estimated Cost", "Value": "Pending rate integration", "Detail": "", "Box Not Available - Substituted Up To VB Box X": ""},
    ]
    boxes_by_base: dict[str, dict] = {}
    for box in box_size_rows:
        base_type = _base_box_type(box.get("Box Type", ""))
        entry = boxes_by_base.setdefault(
            base_type,
            {
                "Section": "Boxes Needed",
                "Metric": base_type,
                "Value": 0,
                "Detail": "",
                "Box Not Available - Substituted Up To VB Box X": "",
            },
        )
        entry["Value"] += int(float(box.get("Box Count") or 0))
        if not entry["Detail"]:
            entry["Detail"] = f"{box.get('Length cm', '')}x{box.get('Width cm', '')}x{box.get('Height cm', '')} cm"
    rows.extend(boxes_by_base.values())
    if unmatched_rows:
        for unmatched in unmatched_rows[:20]:
            rows.append(
                {
                    "Section": "Unmatched SKU Summary",
                    "Metric": unmatched.get("Canonical SKU") or unmatched.get("Raw SKU", ""),
                    "Value": unmatched.get("Reason", ""),
                    "Detail": unmatched.get("Order ID", ""),
                    "Box Not Available - Substituted Up To VB Box X": "",
                }
            )
    else:
        rows.append({"Section": "Unmatched SKU Summary", "Metric": "Unmatched SKUs", "Value": 0, "Detail": "None", "Box Not Available - Substituted Up To VB Box X": ""})

    rule_summaries = _rule_summary(sku_rules)
    if rule_summaries:
        for summary in rule_summaries:
            rows.append({"Section": "Rules Applied Summary", "Metric": "Rule", "Value": summary, "Detail": "", "Box Not Available - Substituted Up To VB Box X": ""})
    else:
        rows.append({"Section": "Rules Applied Summary", "Metric": "Rules", "Value": 0, "Detail": "No special packing rules matched", "Box Not Available - Substituted Up To VB Box X": ""})
    return rows


def _debug_summary_rows(
    result: dict,
    box_size_rows: list[dict],
    unmatched_rows: list[dict],
    warning_rows: list[WorkflowWarning],
    sku_rules: dict[str, SKUCampaignRule],
) -> list[dict]:
    rows = [
        {"Section": "Run Summary", "Metric": "Orders Processed", "Value": result["orders_processed"], "Detail": ""},
        {"Section": "Run Summary", "Metric": "Boxes Created", "Value": result["boxes_created"], "Detail": ""},
        {"Section": "Run Summary", "Metric": "Box Types", "Value": result["box_types"], "Detail": ""},
        {"Section": "Run Summary", "Metric": "Unmatched SKUs", "Value": result["unmatched_skus"], "Detail": ""},
        {"Section": "Run Summary", "Metric": "Warning Count", "Value": result["warning_count"], "Detail": ""},
        {"Section": "Run Summary", "Metric": "Multi-box Orders", "Value": result["multi_box_order_count"], "Detail": ""},
        {"Section": "Run Summary", "Metric": "Rules Applied", "Value": result["rules_applied_count"], "Detail": ""},
        {
            "Section": "Run Summary",
            "Metric": "Chargeable Weight Plans Selected",
            "Value": result.get("chargeable_weight_plan_selected_count", 0),
            "Detail": "",
        },
    ]
    for box in box_size_rows:
        rows.append(
            {
                "Section": "Boxes Needed",
                "Metric": box.get("Box Type", ""),
                "Value": box.get("Box Count", ""),
                "Detail": f"{box.get('Length cm', '')}x{box.get('Width cm', '')}x{box.get('Height cm', '')} cm",
            }
        )
    if unmatched_rows:
        for unmatched in unmatched_rows[:20]:
            rows.append(
                {
                    "Section": "Unmatched SKU Summary",
                    "Metric": unmatched.get("Canonical SKU") or unmatched.get("Raw SKU", ""),
                    "Value": unmatched.get("Reason", ""),
                    "Detail": unmatched.get("Order ID", ""),
                }
            )
    else:
        rows.append({"Section": "Unmatched SKU Summary", "Metric": "Unmatched SKUs", "Value": 0, "Detail": "None"})

    unique_warning_messages = list(dict.fromkeys(warning.message for warning in warning_rows))
    if unique_warning_messages:
        for message in unique_warning_messages[:20]:
            rows.append({"Section": "Unique Warning Summary", "Metric": "Warning", "Value": message, "Detail": ""})
    else:
        rows.append({"Section": "Unique Warning Summary", "Metric": "Warnings", "Value": 0, "Detail": "None"})

    rule_summaries = _rule_summary(sku_rules)
    if rule_summaries:
        for summary in rule_summaries:
            rows.append({"Section": "Rules Applied Summary", "Metric": "Rule", "Value": summary, "Detail": ""})
    else:
        rows.append({"Section": "Rules Applied Summary", "Metric": "Rules", "Value": 0, "Detail": "No special packing rules matched"})
    return rows


def _campaign_info(config: dict) -> dict:
    return config.get("campaign") if isinstance(config.get("campaign"), dict) else {}


def _campaign_name(config: dict) -> str:
    return str(_campaign_info(config).get("name") or "").strip()


def _campaign_vfi_prefix(config: dict) -> str:
    campaign = _campaign_info(config)
    raw = str(
        campaign.get("code")
        or campaign.get("project_name")
        or campaign.get("project")
        or campaign.get("name")
        or "VFI"
    ).strip()
    cleaned = re.sub(r"\s+", "-", raw)
    cleaned = re.sub(r"[^A-Za-z0-9_.-]+", "", cleaned)
    return cleaned or "VFI"



_XLSX_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def _xlsx_column_index(cell_reference: str) -> int:
    letters = re.sub(r"[^A-Z]", "", cell_reference.upper())
    index = 0
    for letter in letters:
        index = index * 26 + ord(letter) - ord("A") + 1
    return index - 1


def _xlsx_shared_strings_raw(archive: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ElementTree.fromstring(archive.read("xl/sharedStrings.xml"))
    values = []
    for item in root.findall("main:si", _XLSX_NS):
        values.append("".join(node.text or "" for node in item.findall(".//main:t", _XLSX_NS)))
    return values


def _xlsx_first_sheet_path(archive: zipfile.ZipFile) -> str:
    workbook = ElementTree.fromstring(archive.read("xl/workbook.xml"))
    rels = ElementTree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    rel_targets = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels.findall("pkgrel:Relationship", _XLSX_NS)
    }
    sheet = workbook.find("main:sheets/main:sheet", _XLSX_NS)
    if sheet is None:
        return ""
    rel_id = sheet.attrib[f"{{{_XLSX_NS['rel']}}}id"]
    target = rel_targets[rel_id].lstrip("/")
    return target if target.startswith("xl/") else f"xl/{target}"


def _xlsx_cell_value_raw(cell, shared_strings: list[str]) -> object:
    inline_node = cell.find("main:is/main:t", _XLSX_NS)
    if inline_node is not None:
        return inline_node.text or ""
    value_node = cell.find("main:v", _XLSX_NS)
    if value_node is None:
        return ""
    value = value_node.text or ""
    if cell.attrib.get("t") == "s":
        return shared_strings[int(value)]
    try:
        number = float(value)
    except ValueError:
        return value
    return int(number) if number.is_integer() else number


def _read_xlsx_first_sheet_table(path: str) -> list[list[object]]:
    with zipfile.ZipFile(path) as archive:
        shared_strings = _xlsx_shared_strings_raw(archive)
        sheet_path = _xlsx_first_sheet_path(archive)
        if not sheet_path:
            return []
        root = ElementTree.fromstring(archive.read(sheet_path))
        raw_rows = []
        for row in root.findall(".//main:row", _XLSX_NS):
            values = {}
            for cell in row.findall("main:c", _XLSX_NS):
                values[_xlsx_column_index(cell.attrib.get("r", ""))] = _xlsx_cell_value_raw(cell, shared_strings)
            if values:
                raw_rows.append(values)
        if not raw_rows:
            return []
        max_index = max(max(row) for row in raw_rows)
        return [[row.get(index, "") for index in range(max_index + 1)] for row in raw_rows]


def _clean_rate_country(value: object) -> str:
    text = str(value or "").strip()
    text = re.sub(r"^[-\s]+", "", text)
    text = re.sub(r"\s+is\s+.+$", "", text, flags=re.I)
    text = text.replace("*", "")
    text = re.sub(r"\s+", " ", text).strip()
    return text.lower()


@dataclass(frozen=True)
class CustomerRateSheet:
    rates_by_zone: dict[str, dict[float, float]]
    zone_by_country: dict[str, str]
    max_weight_kg: float = 49.0


def _load_customer_rate_sheet(path: str) -> CustomerRateSheet | None:
    if not path or not Path(path).exists():
        return None
    table = _read_xlsx_first_sheet_table(path)
    if len(table) < 4:
        return None
    weights = []
    for value in table[1][1:]:
        try:
            weights.append(float(value))
        except (TypeError, ValueError):
            break
    rates_by_zone: dict[str, dict[float, float]] = {}
    for row in table[2:8]:
        zone = str(row[0] or "").strip()
        if not zone.startswith("Zone"):
            continue
        zone_rates = {}
        for weight, amount in zip(weights, row[1:], strict=False):
            try:
                zone_rates[float(weight)] = float(amount)
            except (TypeError, ValueError):
                continue
        if zone_rates:
            rates_by_zone[zone] = dict(sorted(zone_rates.items()))
    zone_by_country: dict[str, str] = {}
    zone_header_index = next(
        (
            index
            for index, row in enumerate(table)
            if any(str(value or "").strip() == "Zone USA" for value in row[:6])
            and any(str(value or "").strip() == "Zone 1" for value in row[:6])
        ),
        -1,
    )
    if zone_header_index >= 0:
        zone_headers = [str(value or "").strip() for value in table[zone_header_index]]
        for row in table[zone_header_index + 1:]:
            for column, zone in enumerate(zone_headers):
                if not zone.startswith("Zone"):
                    continue
                country_key = _clean_rate_country(row[column] if column < len(row) else "")
                if country_key:
                    zone_by_country[country_key] = zone
    return CustomerRateSheet(
        rates_by_zone=rates_by_zone,
        zone_by_country=zone_by_country,
        max_weight_kg=max(weights) if weights else 49.0,
    )


def _customer_rate_sheet_path(cfg: dict) -> str:
    configured = str(cfg.get("rate_sheet_path") or "").strip()
    if not configured:
        return ""
    path = Path(configured)
    if path.is_absolute():
        return str(path)
    return str(Path.cwd() / path)


def _zone_for_cost_summary_row(row: dict, rate_sheet: CustomerRateSheet | None) -> str:
    country = _normalize_country(row.get("Country", ""))
    state = str(row.get("US State Abbreviation") or row.get("State/Province") or "").strip().upper()
    if country in {"United States", "US", "USA"}:
        return "Zone 1" if state in _USA_PLUS_ZONE_STATES else "Zone USA"
    if rate_sheet:
        zone = rate_sheet.zone_by_country.get(_clean_rate_country(country))
        if zone:
            return zone
    return country


def _customer_handling_fee(total_units: object) -> float:
    try:
        units = int(float(total_units))
    except (TypeError, ValueError):
        units = 1
    units = max(units, 1)
    return 2.0 + max(units - 1, 0) * 0.25


def _customer_shipping_fee(
    weight_kg: object,
    zone: str,
    rate_sheet: CustomerRateSheet | None,
    total_units: object = 1,
) -> float | str:
    if not rate_sheet:
        return "Pending rate integration"
    rates = rate_sheet.rates_by_zone.get(zone)
    if not rates:
        return ""
    try:
        weight = float(weight_kg)
    except (TypeError, ValueError):
        return ""
    charge = _rated_charge_for_weight(weight, rates, rate_sheet.max_weight_kg)
    if charge is None:
        return ""
    return round(charge + _customer_handling_fee(total_units), 2)

def _cost_summary_sheet_name(config: dict) -> str:
    name = _campaign_name(config)
    return f"Cost Summary - {name}" if name else "Cost Summary"




def _first_present(row: dict, keys: list[str]) -> object:
    for key in keys:
        value = row.get(key, "")
        if str(value or "").strip():
            return value
    return ""

def _cost_summary_rows(order_rows: list[dict], cfg: dict) -> list[dict]:
    rows = []
    rate_sheet = _load_customer_rate_sheet(_customer_rate_sheet_path(cfg))
    for row in order_rows:
        zone = _zone_for_cost_summary_row(row, rate_sheet)
        rows.append(
            {
                "Backer ID": _first_present(row, ["Backer ID", "BackerKit ID", "Id", "Reward Id"]),
                "VFI #": row.get("VFI #", ""),
                "Shipping name": _first_present(row, ["Shipping name", "Address Name", "Name"]),
                "phone": _first_present(row, ["phone", "Phone", "Address Phone Number"]),
                "email": _first_present(row, ["email", "Email"]),
                "Country": row.get("Country", ""),
                "State/Province": row.get("State/Province", ""),
                "US State Abbreviation": row.get("US State Abbreviation", ""),
                "SKU Breakdown": row.get("SKU Breakdown", ""),
                "Packed Actual Weight kg": row.get("Packed Actual Weight kg", ""),
                "Dimensional Weight kg (/5000)": row.get("Dimensional Weight kg (/5000)", ""),
                "Chargeable Weight kg": row.get("Chargeable Weight kg", ""),
                "Chargeable Weight g": row.get("Chargeable Weight g", ""),
                "Total Units": row.get("Total Units", ""),
                "Box Qty": row.get("Box Qty", ""),
                "Zone": zone,
                "Customer Shipping Fee": _customer_shipping_fee(
                    row.get("Chargeable Weight kg", ""),
                    zone,
                    rate_sheet,
                    row.get("Total Units", ""),
                ),
                "Express": "Pending future rate table",
                "Slow Post": "Pending future rate table",
            }
        )
    return rows


def _label_number_for_box(row: dict, campaign_label_prefix: str = "") -> str:
    vfi_number = str(row.get("VFI #") or row.get("Order ID") or "").strip()
    if not vfi_number:
        return str(row.get("Order Box ID", ""))
    try:
        box_qty = int(float(row.get("Box Qty") or 0))
        box_number = int(float(row.get("Box Number") or 0))
    except (TypeError, ValueError):
        box_qty = 0
        box_number = 0
    match = re.match(r"^(.+)-(\d+)$", vfi_number)
    if match:
        barcode_base = f"{match.group(1)} {match.group(2)}"
    elif campaign_label_prefix:
        barcode_base = f"{campaign_label_prefix} {vfi_number}"
    else:
        barcode_base = vfi_number
    return f"{barcode_base}-{box_number}" if box_qty > 1 and box_number else barcode_base


def _pledge_config_by_combo(box_rows: list[dict]) -> dict[str, int]:
    return {
        combo: index
        for index, (combo, _entry) in enumerate(_combo_entries_for_optimized_to_pack(box_rows), start=1)
    }


def _label_generator_rows(
    box_rows: list[dict],
    pledge_config_by_combo: dict[str, int] | None = None,
    campaign_label_prefix: str = "",
) -> list[dict]:
    pledge_config_by_combo = pledge_config_by_combo or {}
    return [
        {
            "Pledge Configuration": pledge_config_by_combo.get(str(row.get("SKU Breakdown", "")), ""),
            "Order ID": row.get("Order ID", ""),
            "Total Units": row.get("Unit Count", ""),
            "Label numbers": _label_number_for_box(row, campaign_label_prefix),
            "Box Plan": row.get("Box Type", ""),
            "Per-Box Chargeable Weight": row.get("Chargeable Weight kg", ""),
            "SKU Breakdown": row.get("SKUs in Box", ""),
            "Backer ID": _first_present(row, ["Backer ID", "BackerKit ID", "Id", "Reward Id"]),
            "Shipping name": _first_present(row, ["Shipping name", "Address Name", "Name"]),
            "phone": _first_present(row, ["phone", "Phone", "Address Phone Number"]),
            "email": _first_present(row, ["email", "Email"]),
            "add 1": _first_present(row, ["add 1", "Address Line 1", "Shipping Address 1", "Address1"]),
            "add 2": _first_present(row, ["add 2", "Address Line 2", "Shipping Address 2", "Address2"]),
            "Shipping City": _first_present(row, ["Shipping City", "Address City", "City"]),
            "Shipping Postal Code": _first_present(row, ["Shipping Postal Code", "Address Postal Code", "Postal Code", "Zip"]),
            "Country Name": _first_present(row, ["Country Name", "Full Country", "Country"]),
            "Ship to Country Code": _first_present(row, ["Ship to Country Code", "Address Country"]),
            "Name in Chinese": row.get("Name in Chinese", ""),
            "Shipping State": _first_present(row, ["Shipping State", "Address State", "State/Province"]),
        }
        for row in box_rows
    ]


def _labels_rows() -> list[dict]:
    return [{"Note": "Printable label layout and QR generation will be added in a later phase."}]


def _vfi_intake_form_rows(sku_master_path: str) -> list[dict]:
    try:
        source_rows = read_workbook(sku_master_path)
    except Exception:
        return [{"Note": "VFI Intake Form source could not be preserved from uploaded SKU file."}]
    if not source_rows:
        return [{"Note": "No VFI Intake Form rows found in uploaded SKU file."}]
    # Phase A preserves the parsed values from the first SKU/intake sheet. Exact Excel formatting/blank columns can be copied in a later raw-sheet phase.
    return source_rows[0].rows or [{"Note": "VFI Intake Form source sheet was empty."}]


def _box_size_summary(box_rows: list[dict]) -> list[dict]:
    by_type: dict[str, dict] = {}
    for row in box_rows:
        box_type = row["Box Type"]
        summary = by_type.setdefault(
            box_type,
            {
                "Box Type": box_type,
                "Length cm": row["Length cm"],
                "Width cm": row["Width cm"],
                "Height cm": row["Height cm"],
                "Box Count": 0,
                "Order IDs": set(),
                "Unit Count": 0,
                "Chargeable Weights": [],
                "Regions Used": set(),
            },
        )
        summary["Box Count"] += 1
        summary["Order IDs"].add(row["Order ID"])
        summary["Unit Count"] += row.get("Unit Count", 0)
        summary["Chargeable Weights"].append(row["Chargeable Weight kg"])
        if row.get("Region"):
            summary["Regions Used"].add(row["Region"])

    output = []
    for summary in by_type.values():
        weights = summary["Chargeable Weights"] or [0]
        output.append(
            {
                "Box Type": summary["Box Type"],
                "Length cm": summary["Length cm"],
                "Width cm": summary["Width cm"],
                "Height cm": summary["Height cm"],
                "Box Count": summary["Box Count"],
                "Order Count": len(summary["Order IDs"]),
                "Unit Count": summary["Unit Count"],
                "Max Chargeable Weight kg": _format_weight_display(max(weights)),
                "Regions Used": " | ".join(sorted(summary["Regions Used"])),
            }
        )
    return sorted(output, key=lambda row: row["Box Type"])


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
                x, y, z = placement.origin or ("N/A", "N/A", "N/A")
                dimensions = _display_dimensions(placement.dimensions)
                rows.append(
                    {
                        "Order ID": order_id,
                        "Order Box ID": f"{order_id}-{carton.box_number}",
                        "Box Number": carton.box_number,
                        "Canonical SKU": placement.canonical_sku,
                        "Quantity": placement.quantity,
                        "Placement X cm": x,
                        "Placement Y cm": y,
                        "Placement Z cm": z,
                        "Placement Note": "Placed item coordinate" if placement.origin is not None else "Prepacked/forced box; no item coordinate",
                        "Length cm": dimensions.length,
                        "Width cm": dimensions.width,
                        "Height cm": dimensions.height,
                    }
                )
    return rows


def _multi_box_rows(box_rows: list[dict]) -> list[dict]:
    return [
        {
            "Region": row["Region"],
            "Order ID": row["Order ID"],
            "Order Box ID": row["Order Box ID"],
            "Box Number": row["Box Number"],
            "Box Qty": row["Box Qty"],
            "Box Type": row["Box Type"],
            "Vendor Box ID": row.get("Vendor Box ID", ""),
            "Box Selection Decision": row.get("Box Selection Decision", ""),
            "Box Standardization Note": row.get("Box Standardization Note", ""),
            "Length cm": row["Length cm"],
            "Width cm": row["Width cm"],
            "Height cm": row["Height cm"],
            "Actual Weight kg": row["Actual Weight kg"],
            "Packed Actual Weight kg": row["Packed Actual Weight kg"],
            "Dimensional Weight kg": row["Dimensional Weight kg (/5000)"],
            "Chargeable Weight kg": row["Chargeable Weight kg"],
            "SKUs in Box": row["SKUs in Box"],
            "Placement Summary": row["Placement Summary"],
            "Rule Applied": row["Rule Applied"],
            "Warning": row["Warning"],
        }
        for row in box_rows
    ]


def _box_plan(box_rows: list[dict]) -> str:
    return "; ".join(
        f"Box {row['Box Number']}: {row['Box Type']} {row['Length cm']}x{row['Width cm']}x{row['Height cm']} cm"
        for row in sorted(box_rows, key=lambda row: row["Box Number"])
    )


def _joined_box_types(box_rows: list[dict]) -> str:
    types = [row["Box Type"] for row in box_rows if row.get("Box Type")]
    unique_types = []
    for box_type in types:
        if box_type not in unique_types:
            unique_types.append(box_type)
    if len(unique_types) == 1:
        return unique_types[0]
    return " | ".join(unique_types)


def _max_dimensions(box_rows: list[dict], prefix: str = "") -> Dimensions:
    return Dimensions(
        length=max(int(float(row[f"{prefix}Length cm"])) for row in box_rows),
        width=max(int(float(row[f"{prefix}Width cm"])) for row in box_rows),
        height=max(int(float(row[f"{prefix}Height cm"])) for row in box_rows),
    )


def _warning_summary(order_id: str, warnings: list[WorkflowWarning]) -> str:
    return " | ".join(
        dict.fromkeys(warning.message for warning in warnings if warning.order_id == order_id)
    )


def _dedupe_text_warnings(warnings: list[str]) -> list[str]:
    return list(dict.fromkeys(warnings))


def _dedupe_workflow_warnings(warnings: list[WorkflowWarning]) -> list[WorkflowWarning]:
    deduped = []
    seen = set()
    for warning in warnings:
        key = (
            warning.order_id,
            warning.sku,
            warning.stage,
            warning.error_type,
            warning.message,
            warning.rule_applied,
            warning.sku_breakdown,
        )
        if key in seen:
            continue
        seen.add(key)
        deduped.append(warning)
    return deduped


def _with_vfi_number(row: dict, vfi_by_order: dict[str, str]) -> dict:
    order_id = str(row.get("Order ID", ""))
    vfi_number = vfi_by_order.get(order_id, "")
    output = {}
    inserted = False
    for key, value in row.items():
        output[key] = value
        if key == "Order ID":
            output["VFI #"] = vfi_number
            inserted = True
    if not inserted:
        output["VFI #"] = vfi_number
    return output


def _with_vfi_numbers(rows: list[dict], vfi_by_order: dict[str, str]) -> list[dict]:
    return [_with_vfi_number(row, vfi_by_order) for row in rows]


def _combo_entries_for_optimized_to_pack(box_rows: list[dict]) -> list[tuple[str, dict]]:
    combos: dict[str, dict] = {}
    for row in box_rows:
        combo = row["SKU Breakdown"]
        entry = combos.setdefault(
            combo,
            {
                "order_ids": set(),
                "boxes": defaultdict(list),
            },
        )
        entry["order_ids"].add(row["Order ID"])
        entry["boxes"][int(float(row["Box Number"]))].append(row)
    return sorted(
        combos.items(),
        key=lambda item: (-len(item[1]["order_ids"]), item[0]),
    )


def _vfi_numbers_by_order(
    grouped_orders: dict[str, list[OrderLine]],
    combo_by_order: dict[str, str],
    box_rows: list[dict],
    cfg: dict,
) -> dict[str, str]:
    prefix = _campaign_vfi_prefix(cfg)
    vfi_by_order = {}
    next_number = 1
    for combo, _entry in _combo_entries_for_optimized_to_pack(box_rows):
        for order_id in grouped_orders:
            if order_id in vfi_by_order:
                continue
            if combo_by_order.get(order_id) != combo:
                continue
            vfi_by_order[order_id] = f"{prefix}-{next_number}"
            next_number += 1
    return vfi_by_order


def _order_rows_in_vfi_sequence(rows: list[dict], vfi_by_order: dict[str, str]) -> list[dict]:
    def sequence(row: dict) -> tuple[int, str]:
        vfi = str(vfi_by_order.get(str(row.get("Order ID", "")), ""))
        match = re.search(r"-(\d+)$", vfi)
        return (int(match.group(1)) if match else 10**9, str(row.get("Order ID", "")))
    return sorted(rows, key=sequence)


def _box_rows_in_vfi_sequence(rows: list[dict], vfi_by_order: dict[str, str]) -> list[dict]:
    def sequence(row: dict) -> tuple[int, int, str]:
        vfi = str(vfi_by_order.get(str(row.get("Order ID", "")), ""))
        match = re.search(r"-(\d+)$", vfi)
        box_number = int(float(row.get("Box Number") or 0))
        return (
            int(match.group(1)) if match else 10**9,
            box_number,
            str(row.get("Order ID", "")),
        )
    return sorted(rows, key=sequence)


def _order_summary_rows(
    box_rows: list[dict],
    grouped_orders: dict[str, list[OrderLine]],
    items_by_order: dict[str, list[PackedItem]],
    combo_by_order: dict[str, str],
    warning_rows: list[WorkflowWarning],
) -> list[dict]:
    rows = []
    rows_by_order: dict[str, list[dict]] = defaultdict(list)
    for row in box_rows:
        rows_by_order[row["Order ID"]].append(row)

    for order_id, order_box_rows in rows_by_order.items():
        lines = grouped_orders[order_id]
        first_line = lines[0]
        packed_weight = packed_actual_weight_kg(_actual_weight_kg(items_by_order[order_id]))
        dim_weight = sum(float(row["Dimensional Weight kg (/5000)"]) for row in order_box_rows)
        chargeable = sum(float(row["Chargeable Weight g"]) / 1000 for row in order_box_rows)
        box_qty = int(float(order_box_rows[0]["Box Qty"]))
        row = {
            "Region": first_line.region or "",
            "Order ID": order_id,
            "Country": _normalize_country(first_line.country),
            "State/Province": first_line.state_province or "",
            "US State Abbreviation": _state_abbreviation(first_line.country, first_line.state_province),
            "Packed Actual Weight kg": _format_weight_display(packed_weight),
            "Dimensional Weight kg (/5000)": _format_weight_display(dim_weight),
            "Chargeable Weight kg": _format_weight_display(chargeable),
            "Chargeable Weight g": int(chargeable * 1000),
            "Total Units": _total_units(lines),
            "Box Qty": box_qty,
            "Box Plan": _box_plan(order_box_rows),
            "Per-Box Chargeable Weight": _per_box_chargeable_weight_summary(order_box_rows),
            "SKU Breakdown": combo_by_order[order_id],
        }
        rows.append(_append_metadata(row, _metadata_for_order(lines)))
    return rows


def _pledge_combination_summary_rows(order_rows: list[dict]) -> list[dict]:
    return _pledge_combination_summary_rows_from_boxes(order_rows)


def _optimized_to_pack_rows(box_rows: list[dict]) -> list[dict]:
    sorted_entries = _combo_entries_for_optimized_to_pack(box_rows)
    max_box = max((max(entry["boxes"]) for _combo, entry in sorted_entries if entry["boxes"]), default=0)
    output = []
    for index, (combo, entry) in enumerate(sorted_entries, start=1):
        row = {
            "Pledge Configuration": index,
            "Total Pledges": len(entry["order_ids"]),
            "All Items": combo.replace(" | ", ", "),
        }
        for box_number in range(1, max_box + 1):
            rows = entry["boxes"].get(box_number, [])
            if not rows:
                row[f"Box {box_number}"] = ""
                continue
            first = rows[0]
            row[f"Box {box_number}"] = f"{first['Box Type']}: {first['SKUs in Box'].replace(' | ', ', ')}"
        output.append(row)
    return output


def _pledge_combination_summary_rows_from_boxes(box_rows: list[dict]) -> list[dict]:
    grouped: dict[tuple[str, int], list[dict]] = defaultdict(list)
    order_ids_by_combo: dict[str, set[str]] = defaultdict(set)
    units_by_combo: dict[str, dict[str, int]] = defaultdict(dict)
    for row in box_rows:
        combo = row["SKU Breakdown"]
        grouped[(combo, int(float(row["Box Number"])))].append(row)
        order_ids_by_combo[combo].add(row["Order ID"])
        units_by_combo[combo][row["Order ID"]] = int(float(row["Total Units"]))

    output = []
    for (sku_breakdown, box_number), rows in grouped.items():
        chargeable_weights = [float(row["Chargeable Weight kg"]) for row in rows]
        packed_weights = [float(row["Packed Actual Weight kg"]) for row in rows]
        dim_weights = [float(row["Dimensional Weight kg (/5000)"]) for row in rows]
        output.append(
            {
                "SKU Breakdown": sku_breakdown,
                "Order Count": len(order_ids_by_combo[sku_breakdown]),
                "Total Units": sum(units_by_combo[sku_breakdown].values()),
                "Box Number": box_number,
                "Box Qty": " | ".join(sorted({str(row["Box Qty"]) for row in rows})),
                "Order Box Pattern": " | ".join(sorted({row["Order Box ID"] for row in rows})),
                "Box Type": " | ".join(sorted({row["Box Type"] for row in rows})),
                "Length cm": max(int(float(row["Length cm"])) for row in rows),
                "Width cm": max(int(float(row["Width cm"])) for row in rows),
                "Height cm": max(int(float(row["Height cm"])) for row in rows),
                "Packed Actual Weight kg": _format_weight_display(sum(packed_weights) / len(packed_weights)),
                "Dimensional Weight kg": _format_weight_display(sum(dim_weights) / len(dim_weights)),
                "Chargeable Weight kg": _format_weight_display(sum(chargeable_weights) / len(chargeable_weights)),
                "Box Count": len(rows),
                "Regions Used": " | ".join(sorted({row["Region"] for row in rows if row.get("Region")})),
            }
        )
    return sorted(output, key=lambda row: (row["SKU Breakdown"], row["Box Number"]))


def optimize_workbook(
    sku_master_path: str,
    orders_path: str,
    output_path: str,
    config: dict | None = None,
) -> dict:
    """Optimize a SKU/order workbook pair and write an output workbook."""
    started = time.perf_counter()
    cfg = _config(config)
    warnings = []
    warning_rows: list[WorkflowWarning] = []
    if cfg["max_carton_cm"] != DEFAULT_CONFIG["max_carton_cm"]:
        warnings.append("Custom max_carton_cm is not yet supported; using 74 x 37 x 44 cm.")
    if cfg["dimensional_divisor"] != DEFAULT_CONFIG["dimensional_divisor"]:
        warnings.append("Custom dimensional_divisor is not yet supported; using 5000.")
    if cfg["packing_weight_uplift"] != DEFAULT_CONFIG["packing_weight_uplift"]:
        warnings.append("Custom packing_weight_uplift is not yet supported; using 1.15.")

    _log_event("sku_parsing_started")
    _log_event("order_parsing_started")
    intake = read_intake(sku_master_path, orders_path)
    intake = _limit_intake(intake, cfg.get("max_orders"))
    _log_event(
        "matching_finished",
        sku_items=len(intake.sku_items),
        order_lines=len(intake.order_lines),
        matched=len(intake.matched_order_lines),
        unmatched=len(intake.unmatched_skus),
    )
    if intake.unmatched_skus:
        warnings.append(f"{len(intake.unmatched_skus)} unmatched SKU rows were preserved.")
    warnings.extend(_diagnostic_warnings(intake.debug))

    sku_items = _sku_lookup(intake.sku_items)
    sku_rules, matched_rule_keys, unmatched_rule_keys = _match_sku_rules(
        cfg,
        intake.sku_items,
        intake.order_lines,
    )
    for rule_key in unmatched_rule_keys:
        warning_rows.append(
            WorkflowWarning(
                order_id="",
                sku=rule_key,
                stage="config",
                error_type="UnmatchedRuleKey",
                message="sku_rules key did not match any SKU, product name, or order product header.",
                rule_applied=rule_key,
            )
        )
    grouped_orders = _group_order_lines(intake.matched_order_lines)
    split_results = {}
    combo_by_order = {}
    items_by_order = {}
    failed_orders = []
    packing_cache: dict[str, CachedPackingPlan] = {}

    packing_mode = cfg.get("packing_mode", "normal")
    _log_event("packing_started", order_count=len(grouped_orders), packing_mode=packing_mode)
    order_contexts: list[PackingOrderContext] = []
    for first_index, (order_id, lines) in enumerate(grouped_orders.items()):
        order_started = time.perf_counter()
        combo = _sku_breakdown(lines)
        try:
            items = _packed_items_for_order(
                lines,
                sku_items,
                sku_rules,
                bundle_footprint_tolerance_cm=float(cfg.get("bundle_footprint_tolerance_cm", 5)),
            )
            groups = _split_rule_group_records(lines, sku_rules, cfg)
            cache_key = _packing_cache_key(combo, lines, items, groups, sku_rules, cfg)
            order_contexts.append(
                PackingOrderContext(
                    order_id=order_id,
                    lines=lines,
                    combo=combo,
                    items=items,
                    groups=groups,
                    cache_key=cache_key,
                    first_index=first_index,
                )
            )
        except Exception as exc:
            failed_orders.append(order_id)
            message = f"Order packing setup failed and was skipped: {exc}"
            warning_rows.append(
                WorkflowWarning(
                    order_id=order_id,
                    stage="packing",
                    error_type=type(exc).__name__,
                    message=message,
                    sku_breakdown=combo,
                )
            )
            _log_event(
                "order_packing_failed",
                order_id=order_id,
                error_type=type(exc).__name__,
                elapsed_seconds=round(time.perf_counter() - order_started, 3),
            )

    cache_key_counts = Counter(context.cache_key for context in order_contexts)
    first_index_by_cache_key: dict[str, int] = {}
    for context in order_contexts:
        first_index_by_cache_key.setdefault(context.cache_key, context.first_index)
    ranked_cache_keys = sorted(
        cache_key_counts,
        key=lambda key: (-cache_key_counts[key], first_index_by_cache_key[key]),
    )
    cache_key_rank = {key: index + 1 for index, key in enumerate(ranked_cache_keys)}
    if packing_mode == "balanced":
        packing_order = sorted(
            order_contexts,
            key=lambda context: (cache_key_rank[context.cache_key], context.first_index),
        )
    else:
        packing_order = order_contexts

    packing_budget_start = time.perf_counter()
    packing_budget_seconds = (
        float(cfg.get("max_optimization_seconds") or DEFAULT_CONFIG["max_optimization_seconds"])
        if packing_mode == "balanced"
        else None
    )

    for context in packing_order:
        order_started = time.perf_counter()
        order_id = context.order_id
        lines = context.lines
        combo = context.combo
        items = context.items
        groups = context.groups
        cache_key = context.cache_key
        combo_rank = cache_key_rank.get(cache_key, context.first_index + 1)
        item_count = sum(item.quantity for item in items)
        requested_mode = packing_mode
        effective_packing_mode = "cached"
        _log_event(
            "order_packing_started",
            order_id=order_id,
            combo_rank=combo_rank,
            combo_hash=_short_hash(cache_key),
            line_count=len(lines),
            pledge_count=cache_key_counts[cache_key],
            mode=requested_mode,
            item_count=item_count,
        )
        try:
            cached_plan = packing_cache.get(cache_key)
            if cached_plan is not None:
                split_result = _clone_split_result(cached_plan.split_result)
                warning_rows.extend(
                    _materialize_cached_warnings(cached_plan.group_warnings, order_id, combo)
                )
            else:
                effective_packing_mode = packing_mode
                fallback_reason = ""
                remaining_budget = None
                if packing_mode == "balanced" and packing_budget_seconds is not None:
                    elapsed_budget = time.perf_counter() - packing_budget_start
                    remaining_budget = packing_budget_seconds - elapsed_budget
                    min_remaining = float(cfg.get("balanced_min_remaining_seconds", 3))
                    low_budget_threshold = max(min_remaining, packing_budget_seconds * 0.1)
                    max_deep_items = int(cfg.get("balanced_max_items_for_deep_search", 18))
                    if remaining_budget <= low_budget_threshold:
                        effective_packing_mode = "fast"
                        fallback_reason = "budget_exhausted"
                    elif item_count > max_deep_items:
                        effective_packing_mode = "fast"
                        fallback_reason = "combo_too_complex"
                    if fallback_reason:
                        _log_event(
                            "balanced_fallback_fast",
                            order_id=order_id,
                            combo_rank=combo_rank,
                            combo_hash=_short_hash(cache_key),
                            pledge_count=cache_key_counts[cache_key],
                            reason=fallback_reason,
                            elapsed_seconds=round(elapsed_budget, 3),
                            remaining_seconds=round(max(remaining_budget, 0), 3),
                            item_count=item_count,
                            mode="balanced_fallback_fast",
                        )

                split_result, group_warnings_for_cache = _select_chargeable_weight_plan(
                    context=context,
                    sku_items=sku_items,
                    sku_rules=sku_rules,
                    packing_mode=effective_packing_mode,
                    cfg=cfg,
                    remaining_budget_seconds=remaining_budget,
                )
                warning_rows.extend(group_warnings_for_cache)
                packing_cache[cache_key] = CachedPackingPlan(
                    split_result=_clone_split_result(split_result),
                    group_warnings=tuple(group_warnings_for_cache),
                )
            rule_split_message = _rule_split_message(groups)
            if rule_split_message:
                warning_rows.append(
                    WorkflowWarning(
                        order_id=order_id,
                        stage="packing",
                        error_type="RuleBasedOrderSplit",
                        message=rule_split_message,
                        rule_applied=_rule_split_keys(groups),
                        sku_breakdown=combo,
                    )
                )
            _log_event(
                "order_packing_finished",
                order_id=order_id,
                combo_rank=combo_rank,
                combo_hash=_short_hash(cache_key),
                mode=effective_packing_mode or requested_mode,
                cartons=split_result.box_qty,
                success=split_result.success,
                elapsed_seconds=round(time.perf_counter() - order_started, 3),
            )
            if split_result.success:
                split_results[order_id] = split_result
                combo_by_order[order_id] = combo
                items_by_order[order_id] = items
            else:
                failed_orders.append(order_id)
                failed_skus = ", ".join(
                    sorted({item.canonical_sku for item in split_result.unplaced_items})
                )
                message = (
                    "Order contains item(s) that cannot fit inside the 74 x 37 x 44 cm "
                    f"carton cap in any rotation: {failed_skus or 'unknown SKU'}."
                )
                warning_rows.append(
                    WorkflowWarning(
                        order_id=order_id,
                        sku=failed_skus,
                        stage="packing",
                        error_type="OversizedItem",
                        message=message,
                        sku_breakdown=combo,
                    )
                )
        except Exception as exc:
            failed_orders.append(order_id)
            message = f"Order packing failed and was skipped: {exc}"
            warning_rows.append(
                WorkflowWarning(
                    order_id=order_id,
                    stage="packing",
                    error_type=type(exc).__name__,
                    message=message,
                    sku_breakdown=combo,
                )
            )
            _log_event(
                "order_packing_failed",
                order_id=order_id,
                error_type=type(exc).__name__,
                elapsed_seconds=round(time.perf_counter() - order_started, 3),
            )

    if packing_mode == "balanced":
        split_results = {order_id: split_results[order_id] for order_id in grouped_orders if order_id in split_results}
        combo_by_order = {order_id: combo_by_order[order_id] for order_id in grouped_orders if order_id in combo_by_order}
        items_by_order = {order_id: items_by_order[order_id] for order_id in grouped_orders if order_id in items_by_order}

    if failed_orders:
        warnings.append(f"{len(failed_orders)} orders could not be packed.")
    _log_event(
        "packing_finished",
        successful_orders=len(split_results),
        failed_orders=len(failed_orders),
    )
    warnings = _dedupe_text_warnings(warnings)
    warning_rows = _dedupe_workflow_warnings(warning_rows)

    try:
        assignments = _standardize_split_result(split_results, combo_by_order, cfg, sku_rules)
    except Exception as exc:
        warning_rows.append(
            WorkflowWarning(
                order_id="",
                stage="standardization",
                error_type=type(exc).__name__,
                message=f"Box standardization failed; using optimized capped dimensions: {exc}",
            )
        )
        assignments = [
            StandardizedBoxAssignment(
                order_id=carton.order_id,
                combination_key=carton.combination_key,
                box_type=f"Box Type {index + 1}",
                optimized_length_cm=carton.optimized_dimensions.length,
                optimized_width_cm=carton.optimized_dimensions.width,
                optimized_height_cm=carton.optimized_dimensions.height,
                assigned_length_cm=carton.optimized_dimensions.length,
                assigned_width_cm=carton.optimized_dimensions.width,
                assigned_height_cm=carton.optimized_dimensions.height,
                box_standardization_note="Optimized capped dimensions used after standardization warning.",
                placements=carton.placements,
            )
            for index, carton in enumerate(_build_standardization_inputs(split_results, combo_by_order, sku_rules, _vendor_box_fit_mode(cfg)))
        ]
    warning_rows = _dedupe_workflow_warnings(warning_rows)
    assignments_by_key = _assignment_lookup(assignments)
    box_rows = []

    for order_id, lines in grouped_orders.items():
        if order_id not in split_results:
            continue
        split_result = split_results[order_id]
        first_line = lines[0]
        order_box_rows = []
        for index, carton in enumerate(split_result.cartons):
            assignment = assignments_by_key[f"{order_id}#{index + 1}"]
            optimized_dimensions = _carton_dimensions(split_result, index)
            vendor_box_id = "" if carton.box_type else (assignment.vendor_box_id or "")
            selection_decision = "rule_assigned_box" if carton.box_type else assignment.selection_decision
            carton_note_parts = [assignment.box_standardization_note]
            if carton.rule_applied:
                carton_note_parts.append(f"Rule applied: {carton.rule_applied}")
            if carton.warning:
                carton_note_parts.append(carton.warning)
            assigned_dimensions, carton_box_type, cutdown_note = _assigned_carton_dimensions_and_type(
                carton=carton,
                assignment=assignment,
                optimized_dimensions=optimized_dimensions,
            )
            if cutdown_note:
                carton_note_parts.append(cutdown_note)
            box_actual_weight_kg = sum(placement.weight_kg for placement in carton.result.placements)
            box_packed_weight_kg = packed_actual_weight_kg(box_actual_weight_kg)
            dim_weight_kg = dimensional_weight_kg(assigned_dimensions)
            chargeable_kg = max(box_packed_weight_kg, dim_weight_kg)
            skus_in_box = defaultdict(int)
            for placement in carton.result.placements:
                skus_in_box[placement.canonical_sku] += placement.quantity
            row = {
                "Region": first_line.region or "",
                "Order ID": order_id,
                "Order Box ID": f"{order_id}-{carton.box_number}",
                "Box Number": carton.box_number,
                "Country": _normalize_country(first_line.country),
                "State/Province": first_line.state_province or "",
                "US State Abbreviation": _state_abbreviation(first_line.country, first_line.state_province),
                "Actual Weight kg": _format_weight_display(box_actual_weight_kg),
                "Packed Actual Weight kg": _format_weight_display(box_packed_weight_kg),
                "Dimensional Weight kg (/5000)": _format_weight_display(dim_weight_kg),
                "Chargeable Weight kg": _format_weight_display(chargeable_kg),
                "Chargeable Weight g": int(chargeable_kg * 1000),
                "Total Units": _total_units(lines),
                "Unit Count": sum(skus_in_box.values()),
                "Box Qty": split_result.box_qty,
                "Box Type": carton_box_type,
                "Vendor Box ID": vendor_box_id,
                "Box Selection Decision": selection_decision,
                "Length cm": assigned_dimensions.length,
                "Width cm": assigned_dimensions.width,
                "Height cm": assigned_dimensions.height,
                "Optimized Length cm": optimized_dimensions.length,
                "Optimized Width cm": optimized_dimensions.width,
                "Optimized Height cm": optimized_dimensions.height,
                "Assigned Box Length cm": assigned_dimensions.length,
                "Assigned Box Width cm": assigned_dimensions.width,
                "Assigned Box Height cm": assigned_dimensions.height,
                "Box Standardization Note": " ".join(part for part in carton_note_parts if part),
                "Actual Item Weight lb": _format_lb_display(box_actual_weight_kg * KG_TO_LB),
                "Packed Actual Weight lb (+15%)": _format_lb_display(box_packed_weight_kg * KG_TO_LB),
                "Bundled/Padded Volume cm³": _padded_volume_cm3(items_by_order[order_id]),
                "Dimensional Weight lb": _format_lb_display(dim_weight_kg * KG_TO_LB),
                "Chargeable Weight lb": _format_lb_display(chargeable_kg * KG_TO_LB),
                "Distinct SKUs": _distinct_skus(lines),
                "SKU Breakdown": combo_by_order[order_id],
                "SKUs in Box": " | ".join(f"{sku} x{qty}" for sku, qty in sorted(skus_in_box.items())),
                "Placement Summary": f"{len(carton.result.placements)} item placements",
                "Rule Applied": carton.rule_applied,
                "Warning": carton.warning,
            }
            output_row = _append_metadata(row, _metadata_for_order(lines))
            box_rows.append(output_row)
            order_box_rows.append(output_row)

        cap_warning = _carton_cap_warning_for_order(
            order_id=order_id,
            combo=combo_by_order.get(order_id, ""),
            split_result=split_result,
            assignments_by_key=assignments_by_key,
        )
        if cap_warning:
            warning_rows.append(cap_warning)
        retail_warning = _retail_bulk_review_warning(
            order_id=order_id,
            lines=lines,
            order_box_rows=order_box_rows,
            warning_rows=warning_rows,
            combo=combo_by_order.get(order_id, ""),
        )
        if retail_warning:
            warning_rows.append(retail_warning)

    for order_id, lines in grouped_orders.items():
        if order_id not in split_results or not lines:
            continue
        first_line = lines[0]
        if _normalize_country(first_line.country) == "United States" and not _state_abbreviation(first_line.country, first_line.state_province):
            warning_rows.append(
                WorkflowWarning(
                    order_id=order_id,
                    stage="report",
                    error_type="MissingUSStateAbbreviation",
                    message="US order is missing a valid state/territory/armed forces abbreviation for costing calculator.",
                    sku_breakdown=combo_by_order.get(order_id, ""),
                )
            )
    warning_rows = _dedupe_workflow_warnings(warning_rows)

    order_summary_rows = _order_summary_rows(
        box_rows,
        grouped_orders,
        items_by_order,
        combo_by_order,
        warning_rows,
    )
    vfi_by_order = _vfi_numbers_by_order(grouped_orders, combo_by_order, box_rows, cfg)
    box_rows = _with_vfi_numbers(box_rows, vfi_by_order)
    order_summary_rows = _with_vfi_numbers(order_summary_rows, vfi_by_order)
    cost_order_summary_rows = list(order_summary_rows)
    output_granularity = cfg.get("output_granularity", "order_summary")
    order_rows = (
        _box_rows_in_vfi_sequence(box_rows, vfi_by_order)
        if output_granularity == "box_detail"
        else _order_rows_in_vfi_sequence(order_summary_rows, vfi_by_order)
    )

    result = {
        "output_path": str(Path(output_path)),
        "orders_processed": len(split_results),
        "boxes_created": sum(split_result.box_qty for split_result in split_results.values()),
        "box_types": len({row["Box Type"] for row in box_rows}),
        "unmatched_skus": len(intake.unmatched_skus),
        "warnings": warnings,
        "warning_count": len(warnings) + len(warning_rows),
        "multi_box_order_count": sum(
            1 for split_result in split_results.values() if split_result.box_qty > 1
        ),
        "rules_applied_count": len(matched_rule_keys),
        "chargeable_weight_plan_selected_count": sum(
            1
            for warning in warning_rows
            if warning.error_type in {"ChargeableWeightPlanSelected", "OversizedVendorBoxPlanSelected"}
        ),
    }

    region_sheets = {}
    if cfg["preserve_region_sheets"]:
        rows_by_region = defaultdict(list)
        for row in order_rows:
            if row.get("Region"):
                rows_by_region[f"Region - {row['Region']}"].append(row)
        region_sheets = dict(rows_by_region)

    box_size_rows = _box_size_summary(box_rows)
    unmatched_rows = _unmatched_rows(intake.unmatched_skus)
    optimized_to_pack_rows = _optimized_to_pack_rows(box_rows)
    pledge_config_by_combo = _pledge_config_by_combo(box_rows)

    cost_summary_rows = _cost_summary_rows(cost_order_summary_rows, cfg)
    label_generator_rows = _label_generator_rows(
        _box_rows_in_vfi_sequence(box_rows, vfi_by_order),
        pledge_config_by_combo,
        _campaign_vfi_prefix(cfg),
    )
    labels_rows = _labels_rows()
    vfi_intake_form_rows = _vfi_intake_form_rows(sku_master_path)
    workbook_sheets = dict(region_sheets)
    cost_sheet_name = _cost_summary_sheet_name(cfg)
    if cost_sheet_name != "Cost Summary":
        workbook_sheets[cost_sheet_name] = cost_summary_rows

    _log_event("excel_writing_started", output_path=str(Path(output_path).name))
    write_workbook(
        output_path,
        summary_rows=_clean_summary_rows(result, box_size_rows, unmatched_rows, sku_rules),
        cost_summary_rows=cost_summary_rows,
        vfi_intake_form_rows=vfi_intake_form_rows,
        optimized_to_pack_rows=optimized_to_pack_rows,
        label_generator_rows=label_generator_rows,
        labels_rows=labels_rows,
        order_volume_weights_rows=order_rows,
        box_size_summary_rows=box_size_rows,
        unmatched_skus_rows=unmatched_rows,
        packing_detail_rows=_packing_detail_rows(split_results),
        multi_box_detail_rows=_multi_box_rows(box_rows),
        pledge_combination_summary_rows=_pledge_combination_summary_rows_from_boxes(box_rows),
        debug_summary_rows=_debug_summary_rows(result, box_size_rows, unmatched_rows, warning_rows, sku_rules),
        input_column_mapping_rows=intake.column_mappings,
        errors_and_warnings_rows=[
            {
                "Order ID": "",
                "SKU": "",
                "Stage": "workflow",
                "Error Type": "Warning",
                "Message": warning,
                "Rule Applied": "",
                "SKU Breakdown": "",
            }
            for warning in warnings
        ]
        + [_warning_row(warning) for warning in warning_rows],
        sheets=workbook_sheets,
    )
    _log_event(
        "excel_writing_finished",
        output_path=str(Path(output_path).name),
        elapsed_seconds=round(time.perf_counter() - started, 3),
    )

    return result
