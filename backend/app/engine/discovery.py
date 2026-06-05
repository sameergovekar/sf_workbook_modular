from __future__ import annotations

from datetime import datetime
from pathlib import Path

from .sf_client import SuccessFactorsClient, list_entity_type_names


def create_entity_discovery_workbook(
    *,
    client: SuccessFactorsClient,
    output_dir: Path,
    search: str,
) -> tuple[Path, list[dict[str, str]]]:
    try:
        import openpyxl
    except ImportError as exc:
        raise RuntimeError("Discovery workbook needs openpyxl. Run with the bundled Python runtime.") from exc

    metadata = client.fetch_service_metadata()
    terms = [term.strip().lower() for term in search.split(",") if term.strip()]
    rows = []
    for name in list_entity_type_names(metadata):
        lower_name = name.lower()
        matched = [term for term in terms if term in lower_name]
        if not terms or matched:
            rows.append({"entity": name, "matched_terms": ", ".join(matched)})

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Entity Discovery"
    ws.append(["Generated At", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
    ws.append(["Search Terms", search])
    ws.append([])
    ws.append(["Entity", "Matched Terms"])
    for row in rows:
        ws.append([row["entity"], row["matched_terms"]])
    ws.freeze_panes = "A5"
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 38

    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = output_dir / f"entity_discovery_{timestamp}.xlsx"
    wb.save(output_path)
    return output_path, rows
