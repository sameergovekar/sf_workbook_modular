from __future__ import annotations

from pydantic import BaseModel, Field


class TenantConfig(BaseModel):
    api_base_url: str
    company_id: str = ""
    username: str
    password: str


class DiscoveryRequest(TenantConfig):
    search_terms: list[str] = Field(default_factory=lambda: ["onb", "onboarding", "picklist", "rule", "responsible", "group"])


class WorkbookRequest(TenantConfig):
    module: str
    top: int = 200
    entities: list[str] = Field(default_factory=list)


class ArtifactResponse(BaseModel):
    status: str = "success"
    mode: str
    module: str
    artifact: str
    download_url: str
    summary: list[dict]

