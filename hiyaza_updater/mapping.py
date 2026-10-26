from __future__ import annotations

from .constants import BASE_TO_GIS_DEFAULTS, PARCEL_CODE_HEADER
from .models import MappingChoice, WorkbookTable
from .normalization import normalized_header


class MappingResolver:
    @staticmethod
    def find_header(headers: list[str], expected: str) -> str | None:
        target = normalized_header(expected)
        return next((header for header in headers if normalized_header(header) == target), None)

    def resolve_defaults(self, base: WorkbookTable, gis: WorkbookTable) -> tuple[list[MappingChoice], str | None, str | None]:
        choices = [
            MappingChoice(
                target_label=arabic_name,
                base_header=self.find_header(base.headers, arabic_name),
                suggested_gis_header=self.find_header(gis.headers, gis_name),
            )
            for arabic_name, gis_name in BASE_TO_GIS_DEFAULTS.items()
        ]
        return choices, self.find_header(base.headers, PARCEL_CODE_HEADER), self.find_header(gis.headers, PARCEL_CODE_HEADER)
