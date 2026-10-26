from __future__ import annotations

from collections import defaultdict

from .exceptions import SpreadsheetValidationError
from .mapping import MappingResolver
from .models import UpdateReport, WorkbookTable
from .normalization import is_blank_source, normalize_parcel_code, normalized_text, values_equal
from .workbook_store import WorkbookStore


class ParcelUpdateService:
    def __init__(self, store: WorkbookStore, mapping_resolver: MappingResolver):
        self._store = store
        self._mapping_resolver = mapping_resolver

    def update(
        self,
        base: WorkbookTable,
        gis: WorkbookTable,
        base_key_header: str,
        gis_key_header: str,
        mapping: dict[str, str | None],
    ) -> UpdateReport:
        if base_key_header not in base.headers or gis_key_header not in gis.headers:
            raise SpreadsheetValidationError("اختر عمود كود القطعة الصحيح في الملفين.")

        report = UpdateReport(base_total=len(base.rows), gis_total=len(gis.rows))
        base_index, report.base_blank_keys = self._index_rows(base, base_key_header)
        gis_index, report.gis_blank_keys = self._index_rows(gis, gis_key_header)
        base_duplicates = {parcel for parcel, rows in base_index.items() if len(rows) > 1}
        gis_duplicates = {parcel for parcel, rows in gis_index.items() if len(rows) > 1}
        report.base_duplicates = sorted(base_duplicates)
        report.gis_duplicates = sorted(gis_duplicates)
        report.base_without_match = sorted(parcel for parcel in base_index if parcel not in gis_index and parcel not in base_duplicates)
        report.gis_without_match = sorted(parcel for parcel in gis_index if parcel not in base_index and parcel not in gis_duplicates)

        active_mapping = {
            target: source
            for target, source in mapping.items()
            if target in base.headers and source in gis.headers
        }
        usage_field = self._mapping_resolver.find_header(base.headers, "نوع الاستخدام")
        crop_field = self._mapping_resolver.find_header(base.headers, "نوع المحصول")
        stage_field = self._mapping_resolver.find_header(base.headers, "مراحل النمو")

        for parcel, base_rows in base_index.items():
            if parcel in base_duplicates or parcel in gis_duplicates or parcel not in gis_index:
                continue

            base_row, gis_row = base_rows[0], gis_index[parcel][0]
            report.eligible_matches += 1
            changed = self._apply_mapped_values(parcel, base, base_row, gis_row, active_mapping, report)
            changed = self._apply_usage_rule(parcel, base, base_row, usage_field, crop_field, stage_field, report) or changed
            if changed:
                report.updated_rows += 1
        return report

    @staticmethod
    def _index_rows(table: WorkbookTable, key_header: str):
        index = defaultdict(list)
        blank_count = 0
        for row in table.rows:
            parcel = normalize_parcel_code(row.get(key_header))
            if parcel is None:
                blank_count += 1
            else:
                index[parcel].append(row)
        return index, blank_count

    def _apply_mapped_values(self, parcel, base, base_row, gis_row, mapping, report) -> bool:
        changed = False
        for target, source in mapping.items():
            old_value = base_row.get(target)
            new_value = None if is_blank_source(gis_row.get(source)) else gis_row.get(source)
            if values_equal(old_value, new_value):
                continue
            self._store.write_cell(base, base_row, target, new_value)
            report.changes_by_field[target] += 1
            report.blank_changes_by_field[target] += int(is_blank_source(new_value))
            report.details.append((parcel, target, old_value, new_value, "تحديث من GIS"))
            changed = True
        return changed

    def _apply_usage_rule(self, parcel, base, row, usage_field, crop_field, stage_field, report) -> bool:
        if not usage_field or normalized_text(row.get(usage_field)) not in {"بور", "مباني"}:
            return False

        changed = False
        for target in (crop_field, stage_field):
            if target and not is_blank_source(row.get(target)):
                old_value = row.get(target)
                self._store.write_cell(base, row, target, None)
                report.changes_by_field[target] += 1
                report.blank_changes_by_field[target] += 1
                report.rule_changes += 1
                report.details.append((parcel, target, old_value, None, "قاعدة بور/مباني"))
                changed = True
        return changed
