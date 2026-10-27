from pathlib import Path
import unittest

from hiyaza_updater.constants import BASE_TO_GIS_DEFAULTS
from hiyaza_updater.mapping import MappingResolver
from hiyaza_updater.models import WorkbookTable
from hiyaza_updater.update_service import ParcelUpdateService
from hiyaza_updater.validation import MappingContentValidator
from hiyaza_updater.workbook_store import WorkbookStore


class ParcelUpdateServiceTests(unittest.TestCase):
    def setUp(self):
        self.store = WorkbookStore()
        self.service = ParcelUpdateService(self.store, MappingResolver())

    def test_blank_gis_value_and_usage_rule_clear_the_expected_cells(self):
        base = WorkbookTable(
            Path("base.xlsx"),
            ["كود القطعة", *BASE_TO_GIS_DEFAULTS.keys()],
            [{
                "كود القطعة": "10",
                "اسم الحوض": "قديم",
                "كود الحوض": "1",
                "فدان": 1,
                "قراط": 0,
                "سهم": 0,
                "نوع الاستخدام": "زراعي",
                "نوع المحصول": "قمح",
                "مراحل النمو": "نمو",
                "ملاحظات على الحيازة (من فريق العمل)": "ملاحظة",
                "__excel_row__": 2,
            }],
        )
        gis = WorkbookTable(
            Path("gis.xlsx"),
            ["كود القطعة", *BASE_TO_GIS_DEFAULTS.values()],
            [{
                "كود القطعة": "10",
                "Basin_Name": None,
                "Basin_Code": "2",
                "fadan_survey": 3,
                "qirat_survey": 0,
                "sahm_survey": 0,
                "using_type": "بور",
                "Crop_type": "شعير",
                "stage_of_growth": "حصاد",
                "notes_from_survey": "NULL",
                "__excel_row__": 2,
            }],
        )

        report = self.service.update(base, gis, "كود القطعة", "كود القطعة", dict(BASE_TO_GIS_DEFAULTS))

        row = base.rows[0]
        self.assertIsNone(row["اسم الحوض"])
        self.assertEqual(row["كود الحوض"], "2")
        self.assertIsNone(row["نوع المحصول"])
        self.assertIsNone(row["مراحل النمو"])
        self.assertIsNone(row["ملاحظات على الحيازة (من فريق العمل)"])
        self.assertEqual(report.rule_changes, 2)

    def test_numeric_crop_type_uses_domain_description_when_available(self):
        base = WorkbookTable(
            Path("base.xlsx"),
            ["كود القطعة", "نوع المحصول"],
            [{"كود القطعة": "41", "نوع المحصول": "قديم"}],
        )
        gis = WorkbookTable(
            Path("gis.xlsx"),
            ["كودالقطعة", "Crop_type", "d_Crop_type"],
            [{"كودالقطعة": "41", "Crop_type": 7, "d_Crop_type": "ذرة صفراء"}],
        )

        report = self.service.update(
            base,
            gis,
            "كود القطعة",
            "كودالقطعة",
            {"نوع المحصول": "Crop_type"},
        )

        self.assertEqual(base.rows[0]["نوع المحصول"], "ذرة صفراء")
        self.assertEqual(report.unresolved_crop_codes, [])
        self.assertIn("وصف Domain", report.details[0][4])

    def test_numeric_crop_type_without_description_is_reported_and_not_written(self):
        base = WorkbookTable(
            Path("base.xlsx"),
            ["كود القطعة", "نوع المحصول"],
            [{"كود القطعة": "42", "نوع المحصول": "قمح"}],
        )
        gis = WorkbookTable(
            Path("gis.xlsx"),
            ["كودالقطعة", "Crop_type"],
            [{"كودالقطعة": "42", "Crop_type": 7}],
        )

        report = self.service.update(
            base,
            gis,
            "كود القطعة",
            "كودالقطعة",
            {"نوع المحصول": "Crop_type"},
        )

        self.assertEqual(base.rows[0]["نوع المحصول"], "قمح")
        self.assertEqual(report.unresolved_crop_codes, [("42", "7")])
        self.assertEqual(report.changed_cells, 0)

    def test_other_crop_type_fills_crop_type_when_primary_is_blank(self):
        base = WorkbookTable(
            Path("base.xlsx"),
            ["كود القطعة", "نوع المحصول"],
            [{"كود القطعة": "31", "نوع المحصول": "قديم"}],
        )
        gis = WorkbookTable(
            Path("gis.xlsx"),
            ["كودالقطعة", "Crop_type", "other_crop_type"],
            [{"كودالقطعة": "31", "Crop_type": "NULL", "other_crop_type": "بنجر السكر"}],
        )

        report = self.service.update(
            base,
            gis,
            "كود القطعة",
            "كودالقطعة",
            {"نوع المحصول": "Crop_type"},
        )

        self.assertEqual(base.rows[0]["نوع المحصول"], "بنجر السكر")
        self.assertEqual(report.details[0][4], "تحديث من other_crop_type (بديل Crop_type)")

    def test_crop_type_has_priority_over_other_crop_type(self):
        base = WorkbookTable(
            Path("base.xlsx"),
            ["كود القطعة", "نوع المحصول"],
            [{"كود القطعة": "32", "نوع المحصول": "قديم"}],
        )
        gis = WorkbookTable(
            Path("gis.xlsx"),
            ["كودالقطعه", "Crop_type", "other_crop_type"],
            [{"كودالقطعه": "32", "Crop_type": "ذرة", "other_crop_type": "بنجر السكر"}],
        )

        self.service.update(
            base,
            gis,
            "كود القطعة",
            "كودالقطعه",
            {"نوع المحصول": "Crop_type"},
        )

        self.assertEqual(base.rows[0]["نوع المحصول"], "ذرة")

    def test_duplicate_code_is_reported_and_left_unchanged(self):
        base = WorkbookTable(Path("base.xlsx"), ["كود القطعة", "اسم الحوض"], [
            {"كود القطعة": "4", "اسم الحوض": "أ", "__excel_row__": 2},
            {"كود القطعة": "4", "اسم الحوض": "ب", "__excel_row__": 3},
        ])
        gis = WorkbookTable(Path("gis.xlsx"), ["كود القطعة", "Basin_Name"], [
            {"كود القطعة": "4", "Basin_Name": "ج", "__excel_row__": 2},
        ])

        report = self.service.update(base, gis, "كود القطعة", "كود القطعة", {"اسم الحوض": "Basin_Name"})

        self.assertEqual(report.base_duplicates, ["4"])
        self.assertEqual(base.rows[0]["اسم الحوض"], "أ")

    def test_preview_clone_does_not_mutate_the_original_file_data(self):
        base = WorkbookTable(Path("base.xlsx"), ["كود القطعة", "اسم الحوض"], [
            {"كود القطعة": "1", "اسم الحوض": "قديم", "__excel_row__": 2},
        ])
        gis = WorkbookTable(Path("gis.xlsx"), ["كود القطعة", "Basin_Name"], [
            {"كود القطعة": "1", "Basin_Name": "جديد", "__excel_row__": 2},
        ])

        self.service.update(base.clone_for_preview(), gis.clone_for_preview(), "كود القطعة", "كود القطعة", {"اسم الحوض": "Basin_Name"})

        self.assertEqual(base.rows[0]["اسم الحوض"], "قديم")

    def test_notes_header_detects_parentheses_and_ali_variant(self):
        resolver = MappingResolver()
        actual = resolver.find_header(
            ["ملاحظات علي الحيازة (من فريق العمل)"],
            "ملاحظات على الحيازة (من فريق العمل)",
        )
        self.assertEqual(actual, "ملاحظات علي الحيازة (من فريق العمل)")

    def test_gis_parcel_code_without_spaces_is_detected_automatically(self):
        resolver = MappingResolver()
        actual = resolver.find_header(["كودالقطعة"], "كود القطعة")
        self.assertEqual(actual, "كودالقطعة")

    def test_gis_parcel_code_with_heh_is_detected_automatically(self):
        resolver = MappingResolver()
        actual = resolver.find_header(["كودالقطعه"], "كود القطعة")
        self.assertEqual(actual, "كودالقطعه")

    def test_numeric_basin_column_is_blocked_as_a_name_source(self):
        base = WorkbookTable(Path("base.xlsx"), ["كود القطعة", "اسم الحوض"], [])
        gis = WorkbookTable(Path("gis.xlsx"), ["Basin_Code"], [
            {"Basin_Code": 101},
            {"Basin_Code": 102},
        ])
        validator = MappingContentValidator(MappingResolver())

        errors = validator.validate(base, gis, {"اسم الحوض": "Basin_Code"})

        self.assertEqual(len(errors), 1)
        self.assertIn("اسم الحوض", errors[0])

    def test_automatic_mapping_keeps_basin_name_separate_from_basin_code(self):
        base = WorkbookTable(
            Path("base.xlsx"),
            ["كود القطعة", "اسم الحوض", "كود الحوض", "نوع المحصول"],
            [{"كود القطعة": "21", "اسم الحوض": "قديم", "كود الحوض": "1", "نوع المحصول": "قديم"}],
        )
        gis = WorkbookTable(
            Path("gis.xlsx"),
            ["كودالقطعه", "Basin_Code", "Basin_Name", "Crop_type"],
            [{"كودالقطعه": "21", "Basin_Code": 55, "Basin_Name": "حوض النخيل", "Crop_type": "قمح"}],
        )
        resolver = MappingResolver()
        choices, base_key, gis_key = resolver.resolve_defaults(base, gis)
        mapping = {choice.base_header: choice.suggested_gis_header for choice in choices if choice.base_header}
        report = self.service.update(base, gis, base_key, gis_key, mapping)

        self.assertEqual(base.rows[0]["اسم الحوض"], "حوض النخيل")
        self.assertEqual(base.rows[0]["كود الحوض"], 55)
        self.assertEqual(base.rows[0]["نوع المحصول"], "قمح")
        self.assertEqual(report.updated_rows, 1)


if __name__ == "__main__":
    unittest.main()
