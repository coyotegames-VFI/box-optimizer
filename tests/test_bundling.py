from box_optimizer.bundling import bundle_quantity, sku_combination_key
from box_optimizer.models import OrderLine


def test_bundle_quantity_rounds_up():
    assert bundle_quantity(quantity=7, bundle_size=3) == 3


def _line(sku: str, quantity: int) -> OrderLine:
    return OrderLine(
        order_id="order-1",
        raw_sku=sku,
        canonical_sku=sku,
        quantity=quantity,
    )


def test_sku_combination_key_includes_quantities_and_sorted_skus():
    key = sku_combination_key(
        [
            _line("Expansion A", 2),
            _line("Core Game", 1),
            _line("Dice Pack", 1),
        ]
    )

    assert key == "Core Game x1 | Dice Pack x1 | Expansion A x2"


def test_sku_combination_key_is_identical_for_identical_combinations():
    first = sku_combination_key(
        [
            _line("Core Game", 1),
            _line("Expansion A", 2),
            _line("Dice Pack", 1),
        ]
    )
    second = sku_combination_key(
        [
            _line("Dice Pack", 1),
            _line("Expansion A", 1),
            _line("Core Game", 1),
            _line("Expansion A", 1),
        ]
    )

    assert first == second
