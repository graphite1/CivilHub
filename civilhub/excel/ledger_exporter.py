from __future__ import annotations

from datetime import datetime
from pathlib import Path

from .ledger_mapping import DEFAULT_MAPPING_SHEET, load_ledger_cell_mappings


def sanitize_filename(value: str) -> str:
    forbidden = '<>:"/\\|?*'
    cleaned = "".join("_" if ch in forbidden else ch for ch in value.strip())
    return cleaned or "document"


def format_ledger_date(prefix: str, value: str) -> str:
    if not value:
        return ""
    import re

    match = re.match(r"^\s*(\d{4})-(\d{1,2})-(\d{1,2})\s*$", value)
    if match:
        year, month, day = match.groups()
        return f" {prefix} {int(year)}年{int(month)}月{int(day)}日"
    return value


def format_ledger_phone(value: str) -> str:
    if not value:
        return ""
    return f"（TEL　{value}）"


def build_ledger_export_path(
    project_name: str,
    company_name: str | None = None,
    exports_root: Path | None = None,
) -> Path:
    root = exports_root or Path("exports")
    root.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    parts = ["施工体制台帳", sanitize_filename(project_name)]
    if company_name:
        parts.append(sanitize_filename(company_name))
    parts.append(timestamp)
    return root / ("_".join(parts) + ".xlsx")


def write_ledger_mapped_workbook(
    workbook: object,
    project: object,
    root_company: object | None,
    child_company: object | None,
    mapping_path: Path,
    mapping_sheet: str = DEFAULT_MAPPING_SHEET,
) -> None:
    for mapping in load_ledger_cell_mappings(mapping_path, mapping_sheet):
        sheet_name = mapping["sheet_name"]
        if sheet_name not in workbook.sheetnames:
            continue
        value = ledger_field_value(mapping["field"], project, root_company, child_company)
        if value in (None, ""):
            continue
        workbook[sheet_name][mapping["cell"]] = value


def ledger_field_value(
    field: str,
    project: object,
    root_company: object | None,
    child_company: object | None,
) -> str:
    project_values = {
        "project.basic.name": _row_value(project, "name"),
        "selected_contractor.basic.project_name": _row_value(project, "name"),
        "client.name": _row_value(project, "client_name"),
        "project.contract.period_start": format_ledger_date("自", _row_value(project, "start_date")),
        "project.contract.period_end": format_ledger_date("至", _row_value(project, "end_date")),
        "prime_contractor.engineers.site_agent_name": _row_value(project, "site_agent"),
        "prime_contractor.engineers.chief_engineer_name": _row_value(project, "managing_engineer") or _row_value(project, "chief_engineer"),
    }
    if field in project_values:
        return project_values[field]

    if field.startswith("prime_contractor.") and root_company is not None:
        return company_ledger_field_value(field, root_company, is_prime=True)
    if field.startswith("selected_contractor.") and child_company is not None:
        return company_ledger_field_value(field, child_company, is_prime=False)
    return ""


def company_ledger_field_value(field: str, company: object, is_prime: bool) -> str:
    if field.endswith(".company_name_with_corporate_no") or field.endswith(".office_name"):
        return _row_value(company, "name")
    if field.endswith(".representative_name"):
        return _row_value(company, "representative")
    if field.endswith(".address"):
        return _row_value(company, "address")
    if field.endswith(".phone"):
        return format_ledger_phone(_row_value(company, "phone"))
    if field.endswith(".work_description") or field.endswith(".work_description_1") or field.endswith(".work_description_2"):
        return _row_value(company, "work_type")
    if field.endswith(".period_start"):
        return format_ledger_date("自", _row_value(company, "planned_start_date"))
    if field.endswith(".period_end"):
        return format_ledger_date("至", _row_value(company, "planned_end_date"))
    if field.endswith(".contract_date"):
        return format_ledger_date("", _row_value(company, "contract_date")).strip()
    if field.endswith(".permit_business_type_1"):
        return _row_value(company, "work_type")
    if field.endswith(".permit_number_1"):
        return _row_value(company, "license_no")
    if field.endswith(".permit_date_1"):
        return format_ledger_date("", _row_value(company, "license_expiry")).strip()
    if field.endswith(".safety_health_manager_name"):
        return _row_value(company, "safety_manager")
    if field.endswith(".chief_engineer_name"):
        return _row_value(company, "chief_engineer_name")
    if field.endswith(".chief_engineer_qualification"):
        return _row_value(company, "chief_engineer_license")
    if field.endswith(".chief_engineer_assignment_type"):
        return "専任" if _row_value(company, "chief_engineer_name") else ""
    if is_prime and field.endswith(".site_agent_name"):
        return _row_value(company, "representative")
    return ""


def _row_value(row: object, key: str) -> str:
    try:
        value = row[key]  # type: ignore[index]
    except (KeyError, IndexError, TypeError):
        value = getattr(row, key, "")
    return "" if value is None else str(value)
