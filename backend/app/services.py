from __future__ import annotations

from pathlib import Path

from .config import OUTPUTS_DIR, TEMPLATE_FILES
from .engine.discovery import create_entity_discovery_workbook
from .engine.modules import MODULE_ENTITIES
from .engine.onboarding_populator import populate_onboarding_template
from .engine.sf_client import (
    SuccessFactorsClient,
    SuccessFactorsConfig,
    extract_entity_fields,
)
from .engine.template_populator import populate_foundation_template
from .engine.workbook import create_module_workbook


def build_client(api_base_url: str, company_id: str, username: str, password: str) -> SuccessFactorsClient:
    config = SuccessFactorsConfig(
        api_base_url=api_base_url.strip(),
        company_id=company_id.strip(),
        username=username.strip(),
        password=password,
    )
    return SuccessFactorsClient(config)


def run_discovery(*, client: SuccessFactorsClient, search_terms: list[str]) -> tuple[Path, list[dict]]:
    return create_entity_discovery_workbook(
        client=client,
        output_dir=OUTPUTS_DIR,
        search=",".join(search_terms),
    )


def run_employee_central_generic(
    *, client: SuccessFactorsClient, module: str, entities: list[str] | None = None
) -> tuple[Path, list[dict]]:
    module_def = MODULE_ENTITIES[module]
    selected_entities = entities or module_def["entities"]
    entity_results = []
    for entity in selected_entities:
        try:
            metadata = client.fetch_entity_metadata(entity)
            fields = extract_entity_fields(metadata, entity)
            entity_results.append(
                {
                    "entity": entity,
                    "status": "Fetched" if fields else "Fetched - no fields parsed",
                    "fields": fields,
                    "error": "",
                }
            )
        except Exception as exc:
            entity_results.append(
                {
                    "entity": entity,
                    "status": "Error",
                    "fields": [],
                    "error": str(exc),
                }
            )
    output_path = create_module_workbook(
        module_label=module_def["label"],
        instance_name=client.config.company_id or client.config.api_base_url,
        entity_results=entity_results,
        output_dir=OUTPUTS_DIR,
    )
    summary = [
        {
            "sheet": result["entity"],
            "entity": result["entity"],
            "status": result["status"],
            "rows": len(result.get("fields", [])),
            "error": result.get("error", ""),
        }
        for result in entity_results
    ]
    return output_path, summary


def run_foundation_template(*, client: SuccessFactorsClient, top: int) -> tuple[Path, list[dict]]:
    return populate_foundation_template(
        template_path=TEMPLATE_FILES["ec_foundation"],
        output_dir=OUTPUTS_DIR,
        client=client,
        top=top,
    )


def run_onboarding_template(*, client: SuccessFactorsClient, top: int) -> tuple[Path, list[dict]]:
    return populate_onboarding_template(
        template_path=TEMPLATE_FILES["onboarding"],
        output_dir=OUTPUTS_DIR,
        client=client,
        top=top,
    )

