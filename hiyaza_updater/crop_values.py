from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from .mapping import MappingResolver
from .models import WorkbookTable
from .normalization import display_text, is_blank_source


@dataclass(frozen=True)
class CropValueResolution:
    value: Any | None
    reason: str
    should_update: bool = True
    unresolved_code: str | None = None


class CropValueResolver:
    """Resolves GIS coded crop domains into safe text values for the base workbook."""

    _DESCRIPTION_HEADER_CANDIDATES = (
        "d_Crop_type",
        "Crop_type_description",
        "Crop_type_desc",
        "Crop_type_name",
        "Crop_type_label",
    )

    def __init__(self, mapping_resolver: MappingResolver):
        self._mapping_resolver = mapping_resolver

    def find_description_header(self, gis: WorkbookTable) -> str | None:
        for candidate in self._DESCRIPTION_HEADER_CANDIDATES:
            header = self._mapping_resolver.find_header(gis.headers, candidate)
            if header:
                return header
        return None

    def resolve(
        self,
        gis_row: dict[str, Any],
        primary_source: str,
        crop_type_source: str | None,
        other_crop_type_source: str | None,
        description_source: str | None,
    ) -> CropValueResolution:
        primary_value = gis_row.get(primary_source)

        if is_blank_source(primary_value):
            return self._resolve_other_crop(gis_row, other_crop_type_source)

        if not self.is_numeric_code(primary_value):
            return CropValueResolution(primary_value, "تحديث من GIS")

        if primary_source == crop_type_source:
            description_value = gis_row.get(description_source) if description_source else None
            if not is_blank_source(description_value) and not self.is_numeric_code(description_value):
                return CropValueResolution(
                    description_value,
                    f"تحديث من {description_source} (وصف Domain لـ Crop_type)",
                )

        return CropValueResolution(
            None,
            "لم يتم تحديث نوع المحصول: كود Domain بلا وصف",
            should_update=False,
            unresolved_code=display_text(primary_value),
        )

    def _resolve_other_crop(
        self,
        gis_row: dict[str, Any],
        other_crop_type_source: str | None,
    ) -> CropValueResolution:
        fallback_value = gis_row.get(other_crop_type_source) if other_crop_type_source else None
        if is_blank_source(fallback_value):
            return CropValueResolution(None, "تحديث من GIS")
        if self.is_numeric_code(fallback_value):
            return CropValueResolution(
                None,
                "لم يتم تحديث نوع المحصول: كود Domain بلا وصف",
                should_update=False,
                unresolved_code=display_text(fallback_value),
            )
        return CropValueResolution(fallback_value, "تحديث من other_crop_type (بديل Crop_type)")

    @staticmethod
    def is_numeric_code(value: Any) -> bool:
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return True
        rendered = display_text(value).translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))
        try:
            float(rendered.replace(",", ""))
        except ValueError:
            return False
        return bool(rendered)
