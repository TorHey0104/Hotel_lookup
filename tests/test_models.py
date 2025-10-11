from __future__ import annotations

from spirit_lookup.models import Contact, SpiritRecord


def test_display_label_with_location():
    record = SpiritRecord(
        spirit_code="ABC123",
        hotel_name="Test Hotel",
        location_city="Berlin",
        location_country="Deutschland",
    )
    assert record.display_label() == "ABC123 • Test Hotel (Berlin, Deutschland)"


def test_display_label_without_location():
    record = SpiritRecord(spirit_code="XYZ", hotel_name="Minimal Hotel")
    assert record.display_label() == "XYZ • Minimal Hotel"


def test_contact_defaults():
    contact = Contact(role="Manager", name="Dana")
    assert contact.email is None
    assert contact.phone is None
