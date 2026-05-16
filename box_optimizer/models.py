"""Core data models used by box_optimizer."""

from dataclasses import dataclass, field


@dataclass(frozen=True)
class Dimensions:
    """Simple rectangular dimensions."""

    length: float
    width: float
    height: float


@dataclass(frozen=True)
class SKU:
    """A product SKU with dimensions and weight."""

    sku: str
    dimensions: Dimensions
    weight: float


@dataclass(frozen=True)
class SKUItem:
    """Canonical item master data for a SKU."""

    raw_sku: str
    canonical_sku: str
    product_name: str
    length_cm: float
    width_cm: float
    height_cm: float
    weight_kg: float
    is_flat: bool = False
    aliases: tuple[str, ...] = ()
    metadata: dict = field(default_factory=dict)


@dataclass(frozen=True)
class OrderLine:
    """A single order line before packing."""

    order_id: str
    raw_sku: str
    canonical_sku: str
    quantity: int
    region: str | None = None
    country: str | None = None
    state_province: str | None = None
    metadata: dict = field(default_factory=dict)


@dataclass(frozen=True)
class UnmatchedSKURecord:
    """An order SKU that could not be matched to SKU master data."""

    order_line: OrderLine
    reason: str
    metadata: dict = field(default_factory=dict)


@dataclass(frozen=True)
class PackedItem:
    """A packed SKU with dimensions before and after item-level padding."""

    canonical_sku: str
    quantity: int
    unpadded_dimensions: Dimensions
    padded_dimensions: Dimensions
    weight_kg: float
    placement_coordinates: tuple[float, float, float] | None = None


@dataclass(frozen=True)
class Carton:
    """Final carton dimensions and weight metrics."""

    length_cm: float
    width_cm: float
    height_cm: float
    items: list[PackedItem] = field(default_factory=list)
    actual_weight_kg: float = 0.0
    packed_actual_weight_kg: float = 0.0
    dimensional_weight_kg: float = 0.0
    chargeable_weight_kg: float = 0.0
    box_type: str | None = None
    standardization_note: str | None = None
