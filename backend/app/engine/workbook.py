from __future__ import annotations

from datetime import datetime
from pathlib import Path
from html import escape
from zipfile import ZIP_DEFLATED, ZipFile


def create_module_workbook(
    *,
    module_label: str,
    instance_name: str,
    entity_results: list[dict],
    output_dir: Path,
) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_module = module_label.lower().replace(" ", "_")
    output_path = output_dir / f"successfactors_{safe_module}_{timestamp}.xlsx"

    sheets = [("Summary", _summary_rows(module_label, instance_name, entity_results))]
    for result in entity_results:
        sheets.append((_safe_sheet_name(result["entity"]), _entity_rows(result)))

    _write_xlsx(output_path, sheets)
    return output_path


def _summary_rows(
    module_label: str, instance_name: str, entity_results: list[dict]
) -> list[list[str]]:
    rows = [["Entity", "Status", "Field Count", "Error"]]
    for result in entity_results:
        rows.append(
            [
                result["entity"],
                result["status"],
                len(result.get("fields", [])),
                result.get("error", ""),
            ]
        )
    return [
        ["SuccessFactors Configuration Workbook"],
        [],
        ["Module", module_label],
        ["Instance", instance_name],
        ["Generated At", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        [],
        *rows,
    ]


def _entity_rows(result: dict) -> list[list[str]]:
    rows = [
        [result["entity"]],
        [],
        ["Status", result["status"]],
        [],
    ]
    if result.get("error"):
        return [*rows, ["Error", result["error"]]]

    headers = ["Field", "Kind", "Type", "Nullable", "Max Length", "Label"]
    rows.append(headers)
    for field in result.get("fields", []):
        rows.append(
            [
                field.get("field", ""),
                field.get("kind", ""),
                field.get("type", ""),
                field.get("nullable", ""),
                field.get("max_length", ""),
                field.get("label", ""),
            ]
        )
    return rows


def _write_xlsx(output_path: Path, sheets: list[tuple[str, list[list[str]]]]) -> None:
    with ZipFile(output_path, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", _content_types(len(sheets)))
        archive.writestr("_rels/.rels", _root_rels())
        archive.writestr("xl/workbook.xml", _workbook_xml(sheets))
        archive.writestr("xl/_rels/workbook.xml.rels", _workbook_rels(len(sheets)))
        archive.writestr("xl/styles.xml", _styles_xml())
        for index, (_name, rows) in enumerate(sheets, start=1):
            archive.writestr(f"xl/worksheets/sheet{index}.xml", _sheet_xml(rows))


def _content_types(sheet_count: int) -> str:
    sheet_overrides = "".join(
        f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for i in range(1, sheet_count + 1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        f"{sheet_overrides}</Types>"
    )


def _root_rels() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/></Relationships>'
    )


def _workbook_xml(sheets: list[tuple[str, list[list[str]]]]) -> str:
    sheet_nodes = "".join(
        f'<sheet name="{escape(name)}" sheetId="{index}" r:id="rId{index}"/>'
        for index, (name, _rows) in enumerate(sheets, start=1)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<sheets>{sheet_nodes}</sheets></workbook>"
    )


def _workbook_rels(sheet_count: int) -> str:
    sheet_rels = "".join(
        f'<Relationship Id="rId{i}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        f'Target="worksheets/sheet{i}.xml"/>'
        for i in range(1, sheet_count + 1)
    )
    style_rel = (
        f'<Relationship Id="rId{sheet_count + 1}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f"{sheet_rels}{style_rel}</Relationships>"
    )


def _styles_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
        '<borders count="1"><border/></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs>'
        '</styleSheet>'
    )


def _sheet_xml(rows: list[list[str]]) -> str:
    row_nodes = []
    for row_index, row in enumerate(rows, start=1):
        cell_nodes = []
        for col_index, value in enumerate(row, start=1):
            ref = f"{_column_name(col_index)}{row_index}"
            text = escape(str(value or ""))
            cell_nodes.append(f'<c r="{ref}" t="inlineStr"><is><t>{text}</t></is></c>')
        row_nodes.append(f'<row r="{row_index}">{"".join(cell_nodes)}</row>')

    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<cols>'
        '<col min="1" max="1" width="28" customWidth="1"/>'
        '<col min="2" max="8" width="24" customWidth="1"/>'
        '</cols>'
        f'<sheetData>{"".join(row_nodes)}</sheetData>'
        '</worksheet>'
    )


def _column_name(index: int) -> str:
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def _safe_sheet_name(name: str) -> str:
    return "".join(char for char in name if char not in r"[]:*?/\\")[:31] or "Entity"
