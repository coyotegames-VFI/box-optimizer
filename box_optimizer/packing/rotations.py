"""Rotation helpers for rectangular items."""

from itertools import permutations

from box_optimizer.models import Dimensions


def unique_rotations(dimensions: Dimensions) -> list[Dimensions]:
    """Return up to six unique axis-aligned rotations for dimensions."""
    seen = {
        Dimensions(length=length, width=width, height=height)
        for length, width, height in permutations(
            [dimensions.length, dimensions.width, dimensions.height],
            3,
        )
    }
    return sorted(seen, key=lambda item: (item.length, item.width, item.height))


def valid_rotations_for_carton(
    dimensions: Dimensions,
    carton_dimensions: Dimensions,
) -> list[Dimensions]:
    """Return rotations that fit inside the carton at the origin."""
    from box_optimizer.packing.geometry import fits_within_boundaries

    return [
        rotation
        for rotation in unique_rotations(dimensions)
        if fits_within_boundaries(rotation, carton_dimensions)
    ]
