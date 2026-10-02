import yaml
from pathlib import Path

EXCEL_COLUMNS = [
    "USE",
    "기안자",
    "작성날짜",
    "자산명 (Name of Asset)",
    "문서번호 (Document Number)",
    "업체명 (Vendor)",
    "사양 및 모델 번호 (Specification / Model Number)",
    "수량 (Quantity)",
    "P.O NUMBER",
    "입고유무",
    "용도",
    "비고",
]


def load_mapping(path):
    mapping_path = Path(path)
    mapping = yaml.safe_load(mapping_path.read_text(encoding="utf-8"))
    return mapping or {}


def _get_source_value(source, path):
    if path is None:
        return ""
    if isinstance(path, (int, float, bool)):
        return path
    if path in source:
        return source.get(path, "")
    if "." in path:
        current = source
        for part in path.split("."):
            if not isinstance(current, dict):
                return ""
            current = current.get(part, "")
        return current
    return source.get(path, "") or ""


def _apply_use_rules(item, mapping):
    rules = mapping.get("rules") or []
    for rule in rules:
        conditions = rule.get("conditions") or {}
        if not conditions:
            return rule.get("value")
        matches = True
        for field, expected in conditions.items():
            if str(item.get(field, "")).strip() != str(expected).strip():
                matches = False
                break
        if matches:
            return rule.get("value")
    return ""


def transform_items(raw_items, mapping):
    mapping_rules = mapping.get("source_to_excel", {}) or {}
    rows = []
    for item in raw_items:
        row = {column: "" for column in EXCEL_COLUMNS}
        for source_field, excel_column in mapping_rules.items():
            if excel_column not in EXCEL_COLUMNS:
                continue
            if excel_column == "USE":
                continue
            row[excel_column] = _get_source_value(item, source_field)

        use_value = _apply_use_rules(item, mapping)
        if not use_value and mapping_rules.get("USE"):
            use_value = _get_source_value(item, mapping_rules["USE"])
        row["USE"] = use_value or row["USE"]

        if row.get("작성날짜") and isinstance(row["작성날짜"], str):
            row["작성날짜"] = row["작성날짜"].replace("T", " ").replace("Z", "")

        rows.append({column: row.get(column, "") for column in EXCEL_COLUMNS})
    return rows
