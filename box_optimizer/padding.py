"""Padding helpers for packed dimensions."""

from box_optimizer.models import Dimensions


def _sorted_dimensions(dimensions: Dimensions) -> Dimensions:
    length, width, height = sorted(
        [dimensions.length, dimensions.width, dimensions.height],
        reverse=True,
    )
    return Dimensions(length=length, width=width, height=height)


def add_padding(dimensions: Dimensions) -> Dimensions:
    """Add per-item tiered padding after normalization and bundling."""
    normalized = _sorted_dimensions(dimensions)

    is_small = normalized.length <= 7 and normalized.width <= 7 and normalized.height <= 3
    is_medium = (
        normalized.length <= 22
        and normalized.width <= 22
        and normalized.height <= 5
    )

    if is_small:
        padding = Dimensions(length=2, width=2, height=2)
    elif is_medium:
        padding = Dimensions(length=3, width=2, height=2)
    else:
        padding = Dimensions(length=3, width=3, height=2)

    return Dimensions(
        length=normalized.length + padding.length,
        width=normalized.width + padding.width,
        height=normalized.height + padding.height,
    )


def add_final_exterior_padding(dimensions: Dimensions) -> Dimensions:
    """Add final exterior box padding after all items have been packed together."""
    normalized = _sorted_dimensions(dimensions)
    return Dimensions(
        length=normalized.length + 2,
        width=normalized.width + 2,
        height=normalized.height + 2,
    )
