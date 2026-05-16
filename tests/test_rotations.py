from box_optimizer.models import Dimensions
from box_optimizer.packing.rotations import unique_rotations, valid_rotations_for_carton


def test_unique_rotations_returns_all_axis_aligned_rotations():
    assert len(unique_rotations(Dimensions(1, 2, 3))) == 6


def test_unique_rotations_removes_duplicates():
    rotations = unique_rotations(Dimensions(2, 2, 3))

    assert len(rotations) == 3
    assert len(set(rotations)) == 3


def test_unique_rotations_allows_at_most_six_rotations():
    assert len(unique_rotations(Dimensions(1, 2, 3))) <= 6


def test_item_that_only_fits_when_rotated_has_valid_rotation():
    item = Dimensions(10, 4, 3)
    carton = Dimensions(4, 10, 3)

    assert item not in valid_rotations_for_carton(item, carton)
    assert Dimensions(4, 10, 3) in valid_rotations_for_carton(item, carton)
