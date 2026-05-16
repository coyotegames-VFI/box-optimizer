"""Bundling helpers for grouped SKU quantities."""

from collections import Counter

from box_optimizer.models import OrderLine


def bundle_quantity(quantity: int, bundle_size: int) -> int:
    """Return how many bundles are needed for a quantity."""
    if bundle_size <= 0:
        raise ValueError("bundle_size must be greater than zero")
    return (quantity + bundle_size - 1) // bundle_size


def sku_combination_key(order_lines: list[OrderLine]) -> str:
    """Return a stable exact key for an order's canonical SKU quantities."""
    quantities = Counter()
    for line in order_lines:
        quantities[line.canonical_sku] += line.quantity

    return " | ".join(
        f"{canonical_sku} x{quantities[canonical_sku]}"
        for canonical_sku in sorted(quantities)
    )
