from __future__ import annotations

from pathlib import Path


BASE_DIR = Path(__file__).resolve().parents[2]
TEMPLATES_DIR = BASE_DIR / "templates"
OUTPUTS_DIR = BASE_DIR / "outputs"
STATIC_DIR = Path(__file__).resolve().parent / "static"

TEMPLATE_FILES = {
    "ec_foundation": TEMPLATES_DIR / "EC_Foundation_Object_Workbook_2H_2025.xlsx",
    "onboarding": TEMPLATES_DIR / "Onboarding_Implementation_Workbook_2H_2025_Release_Rev_2.xlsx",
}

SUPPORTED_MODULES = {
    "employee_central": "Employee Central",
    "foundation_objects": "Foundation Objects",
    "onboarding": "Onboarding",
}

