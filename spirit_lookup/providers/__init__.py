"""Data providers for Spirit Lookup."""

from __future__ import annotations

import json
import os
import time
from abc import ABC, abstractmethod
from pathlib import Path
from typing import List, Tuple

try:  # optional dependency for SharePoint support
    import requests
except ModuleNotFoundError:  # pragma: no cover - handled gracefully in code
    requests = None  # type: ignore

from ..models import Contact, SpiritRecord


class DataProviderError(RuntimeError):
    """Raised when the data provider encounters an unexpected error."""


class RecordNotFoundError(LookupError):
    """Raised when a Spirit record cannot be found."""


class BaseDataProvider(ABC):
    """Abstract interface for provider implementations."""

    @abstractmethod
    def list_records(
        self, query: str = "", *, page: int = 0, page_size: int = 50
    ) -> Tuple[List[SpiritRecord], bool]:
        """Return a page of `SpiritRecord` results and a flag indicating more data."""

    @abstractmethod
    def get_record(self, spirit_code: str) -> SpiritRecord:
        """Return a single `SpiritRecord` by spirit code."""


class FixtureDataProvider(BaseDataProvider):
    """Load Spirit records from a local JSON fixture."""

    def __init__(self, fixture_path: Path):
        self.fixture_path = fixture_path
        self._records = self._load_fixture()

    def _load_fixture(self) -> List[SpiritRecord]:
        if not self.fixture_path.exists():
            raise DataProviderError(f"Fixture-Datei nicht gefunden: {self.fixture_path}")
        try:
            raw_data = json.loads(self.fixture_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            raise DataProviderError(f"Fixture-Datei konnte nicht gelesen werden: {exc}") from exc

        records: List[SpiritRecord] = []
        for item in raw_data:
            contacts = [
                Contact(
                    role=contact.get("role", ""),
                    name=contact.get("name", ""),
                    email=contact.get("email"),
                    phone=contact.get("phone"),
                )
                for contact in item.get("contacts", [])
            ]
            location = item.get("location", {}) or {}
            records.append(
                SpiritRecord(
                    spirit_code=item.get("spiritCode", ""),
                    hotel_name=item.get("hotelName", ""),
                    region=item.get("region"),
                    status=item.get("status"),
                    location_city=location.get("city"),
                    location_country=location.get("country"),
                    address=location.get("address"),
                    contacts=contacts,
                    meta=item.get("meta", {}),
                )
            )
        records.sort(key=lambda r: (r.spirit_code or "", r.hotel_name))
        return records

    def reload(self) -> None:
        """Reload records from the current fixture file."""

        self._records = self._load_fixture()

    def _filter_records(self, query: str) -> List[SpiritRecord]:
        if not query:
            return list(self._records)
        query_lower = query.lower()
        results = [
            record
            for record in self._records
            if query_lower in record.spirit_code.lower()
            or query_lower in record.hotel_name.lower()
            or (record.location_city and query_lower in record.location_city.lower())
        ]
        return results

    def list_records(
        self, query: str = "", *, page: int = 0, page_size: int = 50
    ) -> Tuple[List[SpiritRecord], bool]:
        records = self._filter_records(query)
        start = page * page_size
        end = start + page_size
        page_records = records[start:end]
        has_more = end < len(records)
        return page_records, has_more

    def get_record(self, spirit_code: str) -> SpiritRecord:
        spirit_code_lower = spirit_code.lower()
        for record in self._records:
            if record.spirit_code.lower() == spirit_code_lower:
                return record
        raise RecordNotFoundError(f"Spirit Code '{spirit_code}' wurde nicht gefunden.")


# coverage: ignore start
class SharePointDataProvider(BaseDataProvider):
    """Retrieve Spirit records from a Microsoft Graph list."""

    def __init__(
        self,
        *,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        site_id: str,
        list_id: str,
    ) -> None:
        if requests is None:
            raise DataProviderError("Das Paket 'requests' ist für den SharePoint-Provider erforderlich.")
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.site_id = site_id
        self.list_id = list_id
        self._token_cache: tuple[float, str] | None = None

    def _get_token(self) -> str:
        if requests is None:  # pragma: no cover - safeguarded by __init__
            raise DataProviderError("Das Paket 'requests' ist für den SharePoint-Provider erforderlich.")
        cache = self._token_cache
        if cache and cache[0] > time.time():
            return cache[1]

        token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        data = {
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": "https://graph.microsoft.com/.default",
        }
        response = requests.post(token_url, data=data, timeout=10)
        if response.status_code != 200:
            raise DataProviderError(
                f"Authentifizierung bei Microsoft Graph fehlgeschlagen: {response.status_code} {response.text}"
            )
        payload = response.json()
        expires_in = int(payload.get("expires_in", 3600))
        token = payload.get("access_token")
        if not token:
            raise DataProviderError("Zugriffstoken konnte nicht ermittelt werden.")
        self._token_cache = (time.time() + expires_in - 60, token)
        return token

    def _request(self, url: str, params: dict | None = None) -> dict:
        if requests is None:  # pragma: no cover - safeguarded by __init__
            raise DataProviderError("Das Paket 'requests' ist für den SharePoint-Provider erforderlich.")
        token = self._get_token()
        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get(url, headers=headers, params=params, timeout=10)
        if response.status_code != 200:
            raise DataProviderError(
                f"SharePoint-Abfrage fehlgeschlagen: {response.status_code} {response.text}"
            )
        return response.json()

    def _map_item(self, item: dict) -> SpiritRecord:
        fields = item.get("fields", {})
        contacts = []
        for contact_item in fields.get("contacts", []) or []:
            contacts.append(
                Contact(
                    role=contact_item.get("role", ""),
                    name=contact_item.get("name", ""),
                    email=contact_item.get("email"),
                    phone=contact_item.get("phone"),
                )
            )
        return SpiritRecord(
            spirit_code=fields.get("spiritCode", ""),
            hotel_name=fields.get("hotelName", ""),
            region=fields.get("region"),
            status=fields.get("status"),
            location_city=fields.get("city"),
            location_country=fields.get("country"),
            address=fields.get("address"),
            contacts=contacts,
            meta={k: v for k, v in fields.items() if k not in {"spiritCode", "hotelName"}},
        )

    def list_records(
        self, query: str = "", *, page: int = 0, page_size: int = 50
    ) -> Tuple[List[SpiritRecord], bool]:
        url = (
            f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/lists/{self.list_id}/items"
        )
        params = {
            "$top": page_size,
            "$skip": page * page_size,
        }
        if query:
            params["$search"] = f"\"{query}\""

        payload = self._request(url, params=params)
        items = payload.get("value", [])
        records = [self._map_item(item) for item in items]
        has_more = "@odata.nextLink" in payload
        return records, has_more

    def get_record(self, spirit_code: str) -> SpiritRecord:
        url = (
            f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/lists/{self.list_id}/items"
        )
        params = {
            "$filter": f"fields/spiritCode eq '{spirit_code}'",
            "$top": 1,
        }
        payload = self._request(url, params=params)
        items = payload.get("value", [])
        if not items:
            raise RecordNotFoundError(f"Spirit Code '{spirit_code}' wurde nicht gefunden.")
        return self._map_item(items[0])
# coverage: ignore end


def create_provider(config) -> BaseDataProvider:
    """Create the appropriate provider based on configuration."""

    if config.use_sharepoint:
        required_env = {
            "SP_TENANT_ID": os.getenv("SP_TENANT_ID"),
            "SP_CLIENT_ID": os.getenv("SP_CLIENT_ID"),
            "SP_CLIENT_SECRET": os.getenv("SP_CLIENT_SECRET"),
            "SP_SITE_ID": os.getenv("SP_SITE_ID"),
            "SP_LIST_ID": os.getenv("SP_LIST_ID"),
        }
        missing = [key for key, value in required_env.items() if not value]
        if missing:
            raise DataProviderError(
                "SharePoint-Konfiguration unvollständig. Fehlende Variablen: "
                + ", ".join(missing)
            )
        return SharePointDataProvider(
            tenant_id=required_env["SP_TENANT_ID"],
            client_id=required_env["SP_CLIENT_ID"],
            client_secret=required_env["SP_CLIENT_SECRET"],
            site_id=required_env["SP_SITE_ID"],
            list_id=required_env["SP_LIST_ID"],
        )

    return FixtureDataProvider(config.fixture_path)
