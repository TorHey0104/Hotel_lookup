"""Data providers for Spirit Lookup."""

from __future__ import annotations

import json
import os
import re
import time
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any, Dict, Iterable, List, Tuple

try:  # optional dependency for SharePoint support
    import requests
except ModuleNotFoundError:  # pragma: no cover - handled gracefully in code
    requests = None  # type: ignore

from ..models import Contact, SpiritRecord


_NORMALIZE_PATTERN = re.compile(r"[^0-9a-z]")
_EMAIL_SUFFIX_PATTERN = re.compile(r"(?i)[\s_.-]*(e[-\s]?mail|mail)$")
_META_PREFIX_PATTERN = re.compile(r"(?i)^meta[\s_.-]*")

_FIELD_ALIASES: Dict[str, List[str]] = {
    "spiritCode": ["spiritcode"],
    "displayField": [
        "hotelname",
        "hotel",
        "propertyname",
        "property",
        "projektname",
        "projectname",
        "assetname",
        "standortname",
        "locationname",
        "site",
    ],
    "region": ["region"],
    "status": ["status", "projektstatus", "pipelinestatus"],
    "city": ["city", "ort", "stadt", "locationcity"],
    "country": ["country", "land", "locationcountry"],
    "address": ["address", "adresse", "street", "strasse", "locationaddress"],
}

_NAME_SUFFIXES = ("name", "kontaktname", "contactname")
_PHONE_SUFFIXES = ("phone", "telefon", "tel")


def _normalize_key(value: str) -> str:
    return _NORMALIZE_PATTERN.sub("", value.lower())


def _looks_like_email(label: str) -> bool:
    normalized = _normalize_key(label)
    return "email" in normalized or "mail" in normalized


def _derive_field_mapping(columns: Iterable[str], existing: Dict[str, str] | None = None) -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    if existing:
        mapping.update({key: value for key, value in existing.items() if isinstance(key, str) and isinstance(value, str)})
    used = {value for value in mapping.values()}

    for field, aliases in _FIELD_ALIASES.items():
        if field in mapping:
            continue
        for column in columns:
            if column in used:
                continue
            if _normalize_key(column) in aliases:
                mapping[field] = column
                used.add(column)
                break

    if "displayField" not in mapping:
        for column in columns:
            if column in used:
                continue
            normalized = _normalize_key(column)
            if normalized.endswith("phone") or normalized.endswith("telefon"):
                continue
            if _looks_like_email(column):
                continue
            mapping["displayField"] = column
            used.add(column)
            break

    return mapping


def _strip_email_suffix(label: str) -> str:
    return _EMAIL_SUFFIX_PATTERN.sub("", label).strip(" -_:")


def _find_related_field(
    fields: Dict[str, str | None], source: str, suffixes: Iterable[str]
) -> str | None:
    base = _normalize_key(_strip_email_suffix(source))
    if not base:
        return None
    for label, value in fields.items():
        if label == source or not value:
            continue
        normalized = _normalize_key(label)
        if not normalized.startswith(base):
            continue
        if any(normalized.endswith(suffix) for suffix in suffixes):
            return value
    return None


def _derive_meta_key(header: str) -> str:
    stripped = _META_PREFIX_PATTERN.sub("", header.strip())
    if not stripped:
        stripped = "metaField"
    parts = [part for part in re.split(r"[^0-9A-Za-z]+", stripped) if part]
    if not parts:
        return "metaField"
    first, *rest = parts
    camel = first[:1].lower() + first[1:]
    for piece in rest:
        camel += piece[:1].upper() + piece[1:]
    return camel


def _coerce_optional(value: Any) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    return text or None


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

        if isinstance(raw_data, dict):
            config = raw_data.get("config", {})
            raw_records = raw_data.get("records", [])
        else:
            config = {}
            raw_records = raw_data

        if not isinstance(raw_records, list):
            raise DataProviderError("Fixture besitzt ein unerwartetes Format.")

        selected_columns = [
            str(column)
            for column in config.get("selectedColumns", [])
            if isinstance(column, str)
        ]
        email_columns = [
            str(column) for column in config.get("emailColumns", []) if isinstance(column, str)
        ]
        raw_mapping = config.get("fieldMapping")
        field_mapping = (
            {str(key): str(value) for key, value in raw_mapping.items() if isinstance(key, str) and isinstance(value, str)}
            if isinstance(raw_mapping, dict)
            else {}
        )

        records: List[SpiritRecord] = []
        for item in raw_records:
            if not isinstance(item, dict):
                continue
            if "fields" in item:
                record = self._create_record_from_fields(item, selected_columns, field_mapping, email_columns)
            else:
                record = self._create_legacy_record(item)
            records.append(record)
        records.sort(key=lambda r: (r.spirit_code or "", r.hotel_name))
        return records

    def _create_legacy_record(self, item: dict) -> SpiritRecord:
        contacts = [
            Contact(
                role=contact.get("role", ""),
                name=contact.get("name", ""),
                email=contact.get("email"),
                phone=contact.get("phone"),
            )
            for contact in item.get("contacts", [])
            if isinstance(contact, dict)
        ]
        location = item.get("location", {}) or {}
        meta = item.get("meta", {}) if isinstance(item.get("meta"), dict) else {}
        record = SpiritRecord(
            spirit_code=str(item.get("spiritCode", "")),
            hotel_name=str(item.get("hotelName", "")),
            region=item.get("region"),
            status=item.get("status"),
            location_city=location.get("city"),
            location_country=location.get("country"),
            address=location.get("address"),
            contacts=contacts,
            meta=meta,
        )
        field_entries: Dict[str, str | None] = {
            "Spirit Code": record.spirit_code,
            "Hotel": record.hotel_name,
        }
        if record.region:
            field_entries["Region"] = record.region
        if record.status:
            field_entries["Status"] = record.status
        if record.location_city:
            field_entries["City"] = record.location_city
        if record.location_country:
            field_entries["Country"] = record.location_country
        if record.address:
            field_entries["Address"] = record.address
        record.fields = field_entries
        record.field_order = list(field_entries.keys())
        for contact in contacts:
            if contact.email:
                record.email_fields[contact.role or "Kontakt"] = contact.email
        return record

    def _create_record_from_fields(
        self,
        item: dict,
        selected_columns: List[str],
        field_mapping: Dict[str, str],
        email_columns: List[str],
    ) -> SpiritRecord:
        raw_fields = item.get("fields", {})
        if not isinstance(raw_fields, dict):
            raw_fields = {}

        order: List[str] = list(selected_columns) if selected_columns else list(raw_fields.keys())
        for column in raw_fields:
            if column not in order:
                order.append(column)

        ordered_fields: Dict[str, str | None] = {
            column: _coerce_optional(raw_fields.get(column)) for column in order
        }

        mapping = _derive_field_mapping(order, field_mapping)
        spirit_column = mapping.get("spiritCode")
        spirit_value = item.get("spiritCode") or ordered_fields.get(spirit_column or "") or ""
        spirit_code = str(spirit_value)

        display_column = mapping.get("displayField")
        display_value = item.get("displayValue") or ordered_fields.get(display_column or "")
        hotel_name = str(display_value) if display_value else spirit_code

        region = _coerce_optional(ordered_fields.get(mapping.get("region", "")))
        status = _coerce_optional(ordered_fields.get(mapping.get("status", "")))
        city = _coerce_optional(ordered_fields.get(mapping.get("city", "")))
        country = _coerce_optional(ordered_fields.get(mapping.get("country", "")))
        address = _coerce_optional(ordered_fields.get(mapping.get("address", "")))

        emails: Dict[str, str] = {}
        if isinstance(item.get("emails"), dict):
            for key, value in item["emails"].items():
                value_str = _coerce_optional(value)
                if value_str:
                    emails[str(key)] = value_str
        else:
            for column in email_columns:
                value_str = _coerce_optional(ordered_fields.get(column))
                if value_str:
                    emails[column] = value_str

        contacts: List[Contact] = []
        for label, email in emails.items():
            base_label = _strip_email_suffix(label) or label
            role = _find_related_field(ordered_fields, label, ("role",)) or base_label
            name = _find_related_field(ordered_fields, label, _NAME_SUFFIXES) or role
            phone = _find_related_field(ordered_fields, label, _PHONE_SUFFIXES)
            contacts.append(Contact(role=role, name=name, email=email, phone=phone))

        meta: Dict[str, object] = {}
        for column, value in ordered_fields.items():
            if value is None:
                continue
            if column in mapping.values() or column in emails:
                continue
            if _normalize_key(column).startswith("meta"):
                meta[_derive_meta_key(column)] = value

        return SpiritRecord(
            spirit_code=spirit_code,
            hotel_name=hotel_name,
            region=region,
            status=status,
            location_city=city,
            location_country=country,
            address=address,
            contacts=contacts,
            meta=meta,
            fields=ordered_fields,
            field_order=order,
            email_fields=emails,
        )

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
            or any(
                value and query_lower in value.lower()
                for value in record.fields.values()
                if value
            )
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
