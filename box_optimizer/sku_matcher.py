"""SKU matching utilities."""

from box_optimizer.normalize import normalize_sku


def sku_matches(left: object, right: object) -> bool:
    """Return whether two SKU values match after normalization."""
    return normalize_sku(left) == normalize_sku(right)
