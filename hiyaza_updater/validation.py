from __future__ import annotations

from typing import Any

from .mapping import MappingResolver
from .models import WorkbookTable
from .normalization import display_text, is_blank_source


class MappingContentValidator:
    """Validates selected GIS columns before any source file is changed."""

    _TEXT_REQUIRED_FIELDS = {
        "اسم الحوض": "اسم الحوض",
        "نوع المحصول": "نوع المحصول",
    }

    def __init__(self, mapping_resolver: MappingResolver):
        self._mapping_resolver = mapping_resolver

    def validate(self, base: WorkbookTable, gis: WorkbookTable, mapping: dict[str, str | None]) -> list[str]:
        errors: list[str] = []
        for expected_header, display_name in self._TEXT_REQUIRED_FIELDS.items():
            actual_base_header = self._mapping_resolver.find_header(list(mapping), expected_header)
            source_header = mapping.get(actual_base_header) if actual_base_header else None
            if not source_header:
                continue
            if self._is_mostly_numeric(gis, source_header):
                errors.append(
                    f"الحقل «{display_name}» مرتبط بالعمود «{source_header}» لكنه يحتوي أرقامًا في الغالب. "
                    "اختر عمود الأسماء الصحيح من ملف GIS قبل التحديث."
                )
        return errors

    @staticmethod
    def example_value(table: WorkbookTable, header: str, limit: int = 42) -> str:
        for row in table.rows:
            value = row.get(header)
            if not is_blank_source(value):
                rendered = display_text(value)
                return rendered if len(rendered) <= limit else f"{rendered[:limit - 1]}…"
        return "لا توجد قيمة"

    @staticmethod
    def _is_mostly_numeric(table: WorkbookTable, header: str) -> bool:
        values = [row.get(header) for row in table.rows if not is_blank_source(row.get(header))]
        if not values:
            return False
        numeric_values = sum(MappingContentValidator._is_numeric(value) for value in values)
        return numeric_values / len(values) >= 0.8

    @staticmethod
    def _is_numeric(value: Any) -> bool:
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return True
        rendered = display_text(value).translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))
        try:
            float(rendered.replace(",", ""))
        except ValueError:
            return False
        return bool(rendered)
