from box_optimizer.models import Dimensions
from box_optimizer.padding import add_final_exterior_padding, add_padding


def test_add_padding_always_adds_two_cm_to_all_dimensions():
    assert add_padding(Dimensions(7, 7, 3)) == Dimensions(9, 9, 5)
    assert add_padding(Dimensions(22, 22, 5)) == Dimensions(24, 24, 7)
    assert add_padding(Dimensions(30, 10, 6)) == Dimensions(32, 12, 8)


def test_add_padding_dimension_order_does_not_affect_result():
    assert add_padding(Dimensions(22, 5, 5)) == add_padding(Dimensions(5, 22, 5))
    assert add_padding(Dimensions(22, 5, 5)) == Dimensions(24, 7, 7)


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
        Dimensions(24, 24, 7),
        Dimensions(32, 12, 8),
    ]


def test_add_final_exterior_padding_happens_after_packing_dimensions_are_combined():
    packed_dimensions = Dimensions(32, 24, 24)

    assert add_final_exterior_padding(packed_dimensions) == Dimensions(34, 26, 26)
