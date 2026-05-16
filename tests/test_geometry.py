from box_optimizer.models import Dimensions
from box_optimizer.packing.geometry import (
    boxes_overlap,
    fill_percentage,
    fits_within_boundaries,
    volume,
)


def test_volume_multiplies_dimensions():
    assert volume(Dimensions(2, 3, 4)) == 24


def test_fill_percentage_uses_volume_ratio():
    assert fill_percentage(Dimensions(2, 5, 5), Dimensions(10, 10, 10)) == 5


def test_boxes_overlap_when_spaces_intersect():
    assert boxes_overlap(
        (0, 0, 0),
        Dimensions(5, 5, 5),
        (4, 0, 0),
        Dimensions(5, 5, 5),
    )


def test_boxes_do_not_overlap_when_touching_faces():
    assert not boxes_overlap(
        (0, 0, 0),
        Dimensions(5, 5, 5),
        (5, 0, 0),
        Dimensions(5, 5, 5),
    )


def test_boxes_do_not_overlap_when_separated():
    assert not boxes_overlap(
        (0, 0, 0),
        Dimensions(5, 5, 5),
        (6, 0, 0),
        Dimensions(5, 5, 5),
    )


def test_item_fits_within_carton_boundaries():
    assert fits_within_boundaries(
        Dimensions(5, 5, 5),
        Dimensions(10, 10, 10),
        origin=(5, 5, 5),
    )


def test_item_cannot_exceed_carton_boundaries():
    assert not fits_within_boundaries(
        Dimensions(6, 5, 5),
        Dimensions(10, 10, 10),
        origin=(5, 5, 5),
    )


def test_item_cannot_start_outside_carton_boundaries():
    assert not fits_within_boundaries(
        Dimensions(1, 1, 1),
        Dimensions(10, 10, 10),
        origin=(-1, 0, 0),
    )
