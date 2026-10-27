from __future__ import annotations

from collections import Counter
from copy import deepcopy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class WorkbookTable:
    path: Path
    headers: list[str]
    rows: list[dict[str, Any]]
    native_workbook: Any | None = None
    native_sheet: Any | None = None
    header_row: int = 1

    @property
    def is_native_excel(self) -> bool:
        return self.native_workbook is not None

    def clone_for_preview(self) -> "WorkbookTable":
        return WorkbookTable(self.path, list(self.headers), deepcopy(self.rows))


@dataclass(frozen=True)
class MappingChoice:
    target_label: str
    base_header: str | None
    suggested_gis_header: str | None


@dataclass
class UpdateReport:
    base_total: int = 0
    gis_total: int = 0
    eligible_matches: int = 0
    updated_rows: int = 0
    changes_by_field: Counter = field(default_factory=Counter)
    blank_changes_by_field: Counter = field(default_factory=Counter)
    rule_changes: int = 0
    base_without_match: list[str] = field(default_factory=list)
    gis_without_match: list[str] = field(default_factory=list)
    base_duplicates: list[str] = field(default_factory=list)
    gis_duplicates: list[str] = field(default_factory=list)
    base_blank_keys: int = 0
    gis_blank_keys: int = 0
    unresolved_crop_codes: list[tuple[str, str]] = field(default_factory=list)
    details: list[tuple[str, str, Any, Any, str]] = field(default_factory=list)

    @property
    def changed_cells(self) -> int:
        return sum(self.changes_by_field.values())
