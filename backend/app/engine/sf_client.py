from __future__ import annotations

from dataclasses import dataclass
from base64 import b64encode
import json
from urllib.parse import urlencode
from urllib.parse import urljoin
from urllib.request import Request, urlopen
from xml.etree import ElementTree as ET


@dataclass(frozen=True)
class SuccessFactorsConfig:
    api_base_url: str
    company_id: str
    username: str
    password: str

    @property
    def basic_auth_username(self) -> str:
        if "@" in self.username or not self.company_id:
            return self.username
        return f"{self.username}@{self.company_id}"


class SuccessFactorsClient:
    def __init__(self, config: SuccessFactorsConfig) -> None:
        self.config = config

    def fetch_entity_metadata(self, entity: str) -> ET.Element:
        url = self._metadata_url(entity)
        with urlopen(self._request(url, "application/xml"), timeout=45) as response:
            return ET.fromstring(response.read())

    def fetch_service_metadata(self) -> ET.Element:
        url = self._service_metadata_url()
        with urlopen(self._request(url, "application/xml"), timeout=60) as response:
            return ET.fromstring(response.read())

    def fetch_entity_records(self, entity: str, top: int = 200) -> list[dict]:
        url = self._entity_url(entity, {"$format": "json", "$top": str(top)})
        with urlopen(self._request(url, "application/json"), timeout=60) as response:
            payload = json.loads(response.read().decode("utf-8"))
        data = payload.get("d", payload)
        if isinstance(data, dict) and isinstance(data.get("results"), list):
            return data["results"]
        if isinstance(data, list):
            return data
        return []

    def _metadata_url(self, entity: str) -> str:
        base = self.config.api_base_url.strip()
        if base.endswith("$metadata"):
            return base
        if not base.endswith("/"):
            base = f"{base}/"
        return urljoin(base, f"{entity}/$metadata")

    def _service_metadata_url(self) -> str:
        base = self.config.api_base_url.strip()
        if base.endswith("$metadata"):
            return base
        if not base.endswith("/"):
            base = f"{base}/"
        return urljoin(base, "$metadata")

    def _entity_url(self, entity: str, params: dict[str, str]) -> str:
        base = self.config.api_base_url.strip()
        if not base.endswith("/"):
            base = f"{base}/"
        return f"{urljoin(base, entity)}?{urlencode(params)}"

    def _request(self, url: str, accept: str) -> Request:
        token = b64encode(
            f"{self.config.basic_auth_username}:{self.config.password}".encode("utf-8")
        ).decode("ascii")
        return Request(
            url,
            headers={
                "Accept": accept,
                "Authorization": f"Basic {token}",
            },
        )


def extract_entity_fields(metadata: ET.Element, entity: str) -> list[dict[str, str]]:
    entity_type = _find_entity_type(metadata, entity)
    if not entity_type:
        return []

    fields: list[dict[str, str]] = []
    for prop in _children_named(entity_type, "Property"):
        fields.append(
            {
                "field": prop.attrib.get("Name", ""),
                "kind": "Property",
                "type": prop.attrib.get("Type", ""),
                "nullable": prop.attrib.get("Nullable", "true"),
                "max_length": prop.attrib.get("MaxLength", ""),
                "label": _annotation_value(prop, "label"),
            }
        )

    for nav in _children_named(entity_type, "NavigationProperty"):
        fields.append(
            {
                "field": nav.attrib.get("Name", ""),
                "kind": "NavigationProperty",
                "type": nav.attrib.get("Type", nav.attrib.get("Relationship", "")),
                "nullable": "",
                "max_length": "",
                "label": _annotation_value(nav, "label"),
            }
        )

    return sorted(fields, key=lambda item: (item["kind"], item["field"].lower()))


def list_entity_type_names(metadata: ET.Element) -> list[str]:
    names = []
    for element in metadata.iter():
        if _local_name(element.tag) == "EntityType":
            name = element.attrib.get("Name")
            if name:
                names.append(name)
    return sorted(set(names), key=str.lower)


def _find_entity_type(node: ET.Element, entity: str) -> ET.Element | None:
    for element in node.iter():
        if _local_name(element.tag) == "EntityType" and element.attrib.get("Name") == entity:
            return element
    return None


def _children_named(node: ET.Element, local_name: str) -> list[ET.Element]:
    return [child for child in list(node) if _local_name(child.tag) == local_name]


def _annotation_value(node: ET.Element, term_fragment: str) -> str:
    for annotation in _children_named(node, "Annotation"):
        term = annotation.attrib.get("Term", "").lower()
        if term_fragment.lower() in term:
            return annotation.attrib.get("String", "")
    return ""


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]
