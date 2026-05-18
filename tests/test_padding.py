from box_optimizer.models import Dimensions
from box_optimizer.padding import add_final_exterior_padding, add_padding


def test_add_padding_small_tier():
    assert add_padding(Dimensions(7, 7, 3)) == Dimensions(9, 9, 5)


def test_add_padding_medium_tier():
    assert add_padding(Dimensions(22, 22, 5)) == Dimensions(25, 24, 7)


def test_add_padding_large_tier():
    assert add_padding(Dimensions(30, 10, 6)) == Dimensions(33, 13, 8)


def test_add_padding_dimension_order_does_not_affect_result():
    assert add_padding(Dimensions(22, 5, 5)) == add_padding(Dimensions(5, 22, 5))
    assert add_padding(Dimensions(22, 5, 5)) == Dimensions(25, 7, 7)


def test_add_padding_applies_per_item_before_final_exterior_padding():
    items = [
        Dimensions(7, 7, 3),
        Dimensions(5, 5, 2),
        Dimensions(22, 22, 5),
        Dimensions(30, 10, 6),
    ]

    assert [add_padding(item) for item in items] == [
        Dimensions(9, 9, 5),
        Dimensions(7, 7, 4),
        Dimensions(25, 24, 7),
        Dimensions(33, 13, 8),
    ]


def test_add_final_exterior_padding_happens_after_packing_dimensions_are_combined():
    packed_dimensions = Dimensions(33, 24, 25)

    assert add_final_exterior_padding(packed_dimensions) == Dimensions(35, 27, 26)
