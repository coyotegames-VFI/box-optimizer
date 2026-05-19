"""Top-level workbook optimization workflow."""

import json
import logging
import math
import re
import time
from collections import defaultdict
from copy import deepcopy
from dataclasses import dataclass, replace
from pathlib import Path

from box_optimizer.bundling import sku_combination_key
from box_optimizer.io.excel_reader import IntakeResult, read_intake
from box_optimizer.io.excel_writer import write_workbook
from box_optimizer.models import Dimensions, OrderLine, PackedItem, SKUItem
from box_optimizer.normalize import normalize_sku
from box_optimizer.packing.geometry import volume
from box_optimizer.packing.packer import (
    MAX_CARTON_DIMENSIONS,
    OptimizedCartonResult,
    _expand_items,
    pack_items,
)
from box_optimizer.packing.splitter import SplitCarton, SplitResult, split_order_into_cartons
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
    "standardization_tolerance_cm": 4,
    "use_vendor_box_menu": True,
    "billing_band_kg": 1.0,
    "custom_box_min_units": 400,
    "preserve_region_sheets": True,
    "debug": False,
    "max_orders": None,
    "packing_mode": "normal",
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
    logger.info(event, extra={"box_optimizer": {"event": event, **fields}})


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


def _packed_items_for_order(
    lines: list[OrderLine],
    sku_lookup: dict[str, SKUItem],
    sku_rules: dict[str, SKUCampaignRule] | None = None,
) -> list[PackedItem]:
    items = []
    for line in lines:
        sku_item = sku_lookup[line.canonical_sku]
        rule = (sku_rules or {}).get(line.canonical_sku)
        unpadded = Dimensions(
            length=sku_item.length_cm,
            width=sku_item.width_cm,
            height=sku_item.height_cm,
        )
        padded = unpadded if rule and (rule.no_padding or rule.prepacked) else add_padding(unpadded)
        if rule and rule.extra_padding_cm:
            padded = Dimensions(
                length=padded.length + rule.extra_padding_cm.length,
                width=padded.width + rule.extra_padding_cm.width,
                height=padded.height + rule.extra_padding_cm.height,
            )
        items.append(
            PackedItem(
                canonical_sku=line.canonical_sku,
                quantity=line.quantity,
                unpadded_dimensions=unpadded,
                padded_dimensions=padded,
                weight_kg=sku_item.weight_kg,
                raw_sku=sku_item.raw_sku,
                product_name=sku_item.product_name,
                rule_key=rule.key if rule else None,
                rule_applied=rule.key if rule else "",
                box_type=rule.box_type if rule else None,
                warning_note=rule.warning_note if rule else "",
                exclude_from_standardization=rule.exclude_from_standardization if rule else False,
                allow_rotation=rule.allow_rotation if rule else True,
                must_stay_flat=rule.must_stay_flat if rule else False,
                allowed_orientations=rule.allowed_orientations if rule else None,
            )
        )
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
        return _display_dimensions(_capped_dimensions(_raw_carton_dimensions(split_result, box_index)))
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
        "standardization_tolerance_cm": cfg.get("standardization_tolerance_cm"),
        "use_vendor_box_menu": cfg.get("use_vendor_box_menu"),
        "billing_band_kg": cfg.get("billing_band_kg"),
        "custom_box_min_units": cfg.get("custom_box_min_units"),
        "box_menu": cfg.get("box_menu"),
        "order_rules": cfg.get("order_rules"),
        "max_carton_cm": cfg.get("max_carton_cm"),
        "dimensional_divisor": cfg.get("dimensional_divisor"),
        "packing_weight_uplift": cfg.get("packing_weight_uplift"),
        "final_exterior_padding_cm": (2, 2, 2),
        "carton_cap_cm": _dimensions_cache_tuple(MAX_CARTON_DIMENSIONS),
    }


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
) -> tuple[SplitResult, list[WorkflowWarning]]:
    warnings = []
    group_rule = sku_rules.get(group[0].canonical_sku) if len(group) == 1 else None
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
                    chargeable_weight_kg=max(
                        carton.result.chargeable_weight_kg or 0,
                        dimensional_weight_kg(dimensions),
                    ),
                    placements=carton.result.placements,
                )
            )
    return optimized


def _assignment_lookup(
    assignments: list[StandardizedBoxAssignment],
) -> dict[str, StandardizedBoxAssignment]:
    return {assignment.order_id: assignment for assignment in assignments}


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
        if flags:
            summaries.append(f"{rule.key}: {', '.join(flags)}")
    return summaries


def _summary_rows(
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
                "Average Chargeable Weight kg": _format_weight_display(sum(weights) / len(weights)),
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
        packed_weight = sum(float(row["Packed Actual Weight kg"]) for row in order_box_rows)
        dim_weight = sum(float(row["Dimensional Weight kg (/5000)"]) for row in order_box_rows)
        chargeable = sum(float(row["Chargeable Weight kg"]) for row in order_box_rows)
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
            "Box Type": _joined_box_types(order_box_rows),
            "Box Plan": _box_plan(order_box_rows),
            "Per-Box Chargeable Weight": _per_box_chargeable_weight_summary(order_box_rows),
            "SKU Breakdown": combo_by_order[order_id],
        }
        rows.append(_append_metadata(row, _metadata_for_order(lines)))
    return rows


def _pledge_combination_summary_rows(order_rows: list[dict]) -> list[dict]:
    return _pledge_combination_summary_rows_from_boxes(order_rows)


def _optimized_to_pack_rows(box_rows: list[dict]) -> list[dict]:
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

    sorted_entries = sorted(
        combos.items(),
        key=lambda item: (-len(item[1]["order_ids"]), item[0]),
    )
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
    for order_id, lines in grouped_orders.items():
        order_started = time.perf_counter()
        combo = _sku_breakdown(lines)
        _log_event("order_packing_started", order_id=order_id, line_count=len(lines))
        try:
            items = _packed_items_for_order(lines, sku_items, sku_rules)
            groups = _split_rule_group_records(lines, sku_rules, cfg)
            cache_key = _packing_cache_key(combo, lines, items, groups, sku_rules, cfg)
            cached_plan = packing_cache.get(cache_key)
            if cached_plan is not None:
                split_result = _clone_split_result(cached_plan.split_result)
                warning_rows.extend(
                    _materialize_cached_warnings(cached_plan.group_warnings, order_id, combo)
                )
            else:
                group_results = []
                group_warnings_for_cache = []
                for group_record in groups:
                    group = group_record.lines
                    group_items = _packed_items_for_order(group, sku_items, sku_rules)
                    group_result, group_warnings = _pack_group(
                        group,
                        group_items,
                        sku_rules,
                        packing_mode,
                    )
                    group_results.append(group_result)
                    group_warnings_for_cache.extend(group_warnings)
                    warning_rows.extend(group_warnings)
                split_result = _merge_split_results(group_results)
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
            if split_result.success and _exterior_cap_violations(split_result):
                if _exterior_cap_violations(split_result):
                    message = (
                        "Final exterior padding would exceed the carton cap; "
                        "reported carton dimensions were capped at 74 x 37 x 44 cm."
                    )
                    warning_rows.append(
                        WorkflowWarning(
                            order_id=order_id,
                            stage="packing",
                            error_type="CartonCapWarning",
                            message=message,
                            sku_breakdown=combo,
                        )
                    )
            _log_event(
                "order_packing_finished",
                order_id=order_id,
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
        assignments = standardize_optimized_cartons(
            _build_standardization_inputs(split_results, combo_by_order),
            tolerance_cm=cfg["standardization_tolerance_cm"],
            use_vendor_box_menu=cfg.get("use_vendor_box_menu", True),
            billing_band_kg=cfg.get("billing_band_kg", 1.0),
            custom_box_min_units=cfg.get("custom_box_min_units", 400),
        )
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
            for index, carton in enumerate(_build_standardization_inputs(split_results, combo_by_order))
        ]
    warning_rows = _dedupe_workflow_warnings(warning_rows)
    assignments_by_key = _assignment_lookup(assignments)
    box_rows = []

    for order_id, lines in grouped_orders.items():
        if order_id not in split_results:
            continue
        split_result = split_results[order_id]
        first_line = lines[0]
        for index, carton in enumerate(split_result.cartons):
            assignment = assignments_by_key[f"{order_id}#{index + 1}"]
            optimized_dimensions = _carton_dimensions(split_result, index)
            raw_carton_box_type = carton.box_type or assignment.box_type
            vendor_box_id = "" if carton.box_type else (assignment.vendor_box_id or "")
            selection_decision = "rule_assigned_box" if carton.box_type else assignment.selection_decision
            carton_note_parts = [assignment.box_standardization_note]
            if carton.rule_applied:
                carton_note_parts.append(f"Rule applied: {carton.rule_applied}")
            if carton.warning:
                carton_note_parts.append(carton.warning)
            assigned_dimensions = _display_dimensions(
                Dimensions(
                    optimized_dimensions.length if carton.box_type else assignment.assigned_length_cm,
                    optimized_dimensions.width if carton.box_type else assignment.assigned_width_cm,
                    optimized_dimensions.height if carton.box_type else assignment.assigned_height_cm,
                ),
                cap=bool(carton.box_type),
            )
            carton_box_type = _compact_box_type(raw_carton_box_type, assigned_dimensions)
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
            box_rows.append(_append_metadata(row, _metadata_for_order(lines)))

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
    output_granularity = cfg.get("output_granularity", "order_summary")
    order_rows = box_rows if output_granularity == "box_detail" else order_summary_rows

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

    _log_event("excel_writing_started", output_path=str(Path(output_path).name))
    write_workbook(
        output_path,
        summary_rows=_summary_rows(result, box_size_rows, unmatched_rows, warning_rows, sku_rules),
        order_volume_weights_rows=order_rows,
        optimized_to_pack_rows=optimized_to_pack_rows,
        box_size_summary_rows=box_size_rows,
        unmatched_skus_rows=unmatched_rows,
        packing_detail_rows=_packing_detail_rows(split_results),
        multi_box_detail_rows=_multi_box_rows(box_rows),
        pledge_combination_summary_rows=_pledge_combination_summary_rows_from_boxes(box_rows),
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
        sheets=region_sheets,
    )
    _log_event(
        "excel_writing_finished",
        output_path=str(Path(output_path).name),
        elapsed_seconds=round(time.perf_counter() - started, 3),
    )

    return result
