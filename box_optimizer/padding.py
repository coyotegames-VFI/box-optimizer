"""Padding helpers for packed dimensions."""

from box_optimizer.models import Dimensions


def _sorted_dimensions(dimensions: Dimensions) -> Dimensions:
    length, width, height = sorted(
        [dimensions.length, dimensions.width, dimensions.height],
        reverse=True,
    )
    return Dimensions(length=length, width=width, height=height)


def add_padding(dimensions: Dimensions) -> Dimensions:
    """Add normal per-item or per-bundle padding after normalization."""
    normalized = _sorted_dimensions(dimensions)
    return Dimensions(
        length=normalized.length + 2,
        width=normalized.width + 2,
        height=normalized.height + 2,
    )


def add_final_exterior_padding(dimensions: Dimensions) -> Dimensions:
    """Add final exterior box padding after all items have been packed together."""
    normalized = _sorted_dimensions(dimensions)
    return Dimensions(
        length=normalized.length + 2,
        width=normalized.width + 2,
        height=normalized.height + 2,
    )
