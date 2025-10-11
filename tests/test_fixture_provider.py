from __future__ import annotations

from spirit_lookup.providers import RecordNotFoundError


def test_fixture_provider_list_and_get(fixture_provider):
    records, has_more = fixture_provider.list_records()
    assert len(records) == 3
    assert has_more is False

    record = fixture_provider.get_record("ZRH001")
    assert record.hotel_name == "Hyatt Regency Zurich"
    assert record.contacts[0].email == "max@hyatt.com"


def test_fixture_provider_filters(fixture_provider):
    records, has_more = fixture_provider.list_records("lon")
    assert len(records) == 1
    assert records[0].spirit_code == "LON123"
    assert has_more is False


def test_fixture_provider_not_found(fixture_provider):
    try:
        fixture_provider.get_record("MISSING")
    except RecordNotFoundError as exc:
        assert "Spirit Code" in str(exc)
    else:
        raise AssertionError("Expected RecordNotFoundError")
