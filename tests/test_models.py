from box_optimizer.models import Carton, Dimensions, OrderLine, PackedItem, SKUItem


def test_sku_item_captures_canonical_product_data():
    item = SKUItem(
        raw_sku=" abc ",
        canonical_sku="ABC",
        product_name="Sample Product",
        length_cm=10,
        width_cm=5,
        height_cm=1,
        weight_kg=0.25,
        is_flat=True,
        aliases=("ABC-1",),
        metadata={"source": "test"},
    )

    assert item.canonical_sku == "ABC"
    assert item.is_flat is True
    assert item.metadata["source"] == "test"


def test_order_line_allows_region_fields_and_metadata():
    line = OrderLine(
        order_id="order-1",
        raw_sku="abc",
        canonical_sku="ABC",
        quantity=2,
        region="NA",
        country="US",
        state_province="CA",
        metadata={"channel": "web"},
    )

    assert line.quantity == 2
    assert line.country == "US"
    assert line.metadata["channel"] == "web"


def test_packed_item_tracks_unpadded_padded_and_optional_placement():
    packed = PackedItem(
        canonical_sku="ABC",
        quantity=1,
        unpadded_dimensions=Dimensions(10, 5, 2),
        padded_dimensions=Dimensions(12, 7, 4),
        weight_kg=0.5,
        placement_coordinates=(0, 1, 2),
    )

    assert packed.unpadded_dimensions == Dimensions(10, 5, 2)
    assert packed.padded_dimensions == Dimensions(12, 7, 4)
    assert packed.placement_coordinates == (0, 1, 2)


def test_carton_tracks_items_and_weight_metrics():
    packed = PackedItem(
        canonical_sku="ABC",
        quantity=1,
        unpadded_dimensions=Dimensions(10, 5, 2),
        padded_dimensions=Dimensions(12, 7, 4),
        weight_kg=0.5,
    )
    carton = Carton(
        length_cm=13,
        width_cm=8,
        height_cm=5,
        items=[packed],
        actual_weight_kg=0.5,
        packed_actual_weight_kg=0.575,
        dimensional_weight_kg=0.104,
        chargeable_weight_kg=0.575,
        box_type="custom",
        standardization_note="No standard box selected.",
    )

    assert carton.items == [packed]
    assert carton.chargeable_weight_kg == 0.575
    assert carton.box_type == "custom"
