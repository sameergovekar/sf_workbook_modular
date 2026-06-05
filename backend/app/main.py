from __future__ import annotations

from pathlib import Path

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.responses import JSONResponse

from .config import OUTPUTS_DIR, STATIC_DIR, SUPPORTED_MODULES, TEMPLATE_FILES
from .schemas import ArtifactResponse, DiscoveryRequest, WorkbookRequest
from .services import (
    build_client,
    run_discovery,
    run_employee_central_generic,
    run_foundation_template,
    run_onboarding_template,
)

app = FastAPI(title="SuccessFactors Workbook Modular API", version="0.1.0")


@app.exception_handler(Exception)
async def unhandled_exception_handler(_request: Request, exc: Exception):
    return JSONResponse(status_code=500, content={"detail": str(exc)})


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    return (STATIC_DIR / "index.html").read_text(encoding="utf-8")


@app.get("/api/health")
def health() -> dict:
    return {"status": "ok"}


@app.get("/api/modules")
def modules() -> dict:
    return {"modules": SUPPORTED_MODULES}


@app.get("/api/templates")
def templates() -> dict:
    return {
        "templates": {
            "ec_foundation": {
                "label": "EC Foundation Object Workbook 2H 2025",
                "path": str(TEMPLATE_FILES["ec_foundation"]),
            },
            "onboarding": {
                "label": "Onboarding Implementation Workbook 2H 2025",
                "path": str(TEMPLATE_FILES["onboarding"]),
            },
        }
    }


@app.post("/api/discovery/entities", response_model=ArtifactResponse)
def discovery_entities(payload: DiscoveryRequest) -> ArtifactResponse:
    client = build_client(payload.api_base_url, payload.company_id, payload.username, payload.password)
    artifact, rows = run_discovery(client=client, search_terms=payload.search_terms)
    summary = [{"sheet": "Entity Discovery", "entity": "Service $metadata", "status": "Discovered", "rows": len(rows), "error": ""}]
    return ArtifactResponse(
        mode="discovery",
        module="discovery",
        artifact=artifact.name,
        download_url=f"/api/artifacts/{artifact.name}",
        summary=summary,
    )


@app.post("/api/workbooks/employee-central", response_model=ArtifactResponse)
def employee_central_workbook(payload: WorkbookRequest) -> ArtifactResponse:
    client = build_client(payload.api_base_url, payload.company_id, payload.username, payload.password)
    artifact, summary = run_employee_central_generic(client=client, module="employee_central", entities=payload.entities or None)
    return ArtifactResponse(
        mode="generic",
        module="employee_central",
        artifact=artifact.name,
        download_url=f"/api/artifacts/{artifact.name}",
        summary=summary,
    )


@app.post("/api/workbooks/foundation", response_model=ArtifactResponse)
def foundation_workbook(payload: WorkbookRequest) -> ArtifactResponse:
    client = build_client(payload.api_base_url, payload.company_id, payload.username, payload.password)
    artifact, summary = run_foundation_template(client=client, top=payload.top)
    return ArtifactResponse(
        mode="template",
        module="foundation_objects",
        artifact=artifact.name,
        download_url=f"/api/artifacts/{artifact.name}",
        summary=summary,
    )


@app.post("/api/workbooks/onboarding", response_model=ArtifactResponse)
def onboarding_workbook(payload: WorkbookRequest) -> ArtifactResponse:
    client = build_client(payload.api_base_url, payload.company_id, payload.username, payload.password)
    artifact, summary = run_onboarding_template(client=client, top=payload.top)
    return ArtifactResponse(
        mode="template",
        module="onboarding",
        artifact=artifact.name,
        download_url=f"/api/artifacts/{artifact.name}",
        summary=summary,
    )


@app.get("/api/artifacts/{filename}")
def artifact(filename: str) -> FileResponse:
    path = (OUTPUTS_DIR / Path(filename).name).resolve()
    if not path.exists() or path.parent != OUTPUTS_DIR.resolve():
        raise HTTPException(status_code=404, detail="Artifact not found")
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=path.name,
    )
