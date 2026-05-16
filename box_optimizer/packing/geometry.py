"""Geometry helpers for packing."""

from box_optimizer.models import Dimensions

Coordinate = tuple[float, float, float]


def volume(dimensions: Dimensions) -> float:
    """Return rectangular volume."""
    return dimensions.length * dimensions.width * dimensions.height


def fill_percentage(used_dimensions: Dimensions, carton_dimensions: Dimensions) -> float:
    """Return used volume as a percentage of carton volume."""
    carton_volume = volume(carton_dimensions)
    if carton_volume == 0:
        raise ValueError("carton volume must be greater than zero")
    return volume(used_dimensions) / carton_volume * 100


def fits_within_boundaries(
    item_dimensions: Dimensions,
    carton_dimensions: Dimensions,
    origin: Coordinate = (0, 0, 0),
) -> bool:
    """Return whether an item at origin stays inside carton boundaries."""
    x, y, z = origin
    return (
        x >= 0
        and y >= 0
        and z >= 0
        and x + item_dimensions.length <= carton_dimensions.length
        and y + item_dimensions.width <= carton_dimensions.width
        and z + item_dimensions.height <= carton_dimensions.height
    )


def boxes_overlap(
    first_origin: Coordinate,
    first_dimensions: Dimensions,
    second_origin: Coordinate,
    second_dimensions: Dimensions,
) -> bool:
    """Return whether two axis-aligned rectangular boxes overlap."""
    first_x, first_y, first_z = first_origin
    second_x, second_y, second_z = second_origin

    return (
        first_x < second_x + second_dimensions.length
        and first_x + first_dimensions.length > second_x
        and first_y < second_y + second_dimensions.width
        and first_y + first_dimensions.width > second_y
        and first_z < second_z + second_dimensions.height
        and first_z + first_dimensions.height > second_z
    )
