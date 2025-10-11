from __future__ import annotations

import pytest

from spirit_lookup.controller import SpiritLookupController
from spirit_lookup.providers import RecordNotFoundError


@pytest.fixture()
def controller(fixture_provider):
    return SpiritLookupController(fixture_provider, page_size=2)


def test_controller_pagination(controller):
    result = controller.list_records(page=0)
    assert len(result.records) == 2
    assert result.has_more is True

    next_result = controller.list_records(page=1)
    assert len(next_result.records) >= 1
    assert next_result.has_more is False


def test_controller_search_by_input_with_code(controller):
    record = controller.search_by_input(spirit_code="DXB777", selected_label=None, cached_records=[])
    assert record.location_city == "Dubai"


def test_controller_search_by_input_with_label(controller):
    result = controller.list_records(page=0)
    label = result.records[0].display_label()
    record = controller.search_by_input(spirit_code=None, selected_label=label, cached_records=result.records)
    assert record.spirit_code == result.records[0].spirit_code


def test_controller_search_without_selection(controller):
    with pytest.raises(RecordNotFoundError):
        controller.search_by_input(spirit_code=None, selected_label=None, cached_records=[])
