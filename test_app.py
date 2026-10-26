import unittest
from pathlib import Path
from app import BASE_TO_GIS, Table, apply_update, key


class UpdateRulesTest(unittest.TestCase):
    def base_row(self):
        return {
            "كود القطعة": "10",
            "اسم الحوض": "قديم",
            "كود الحوض": "1",
            "فدان": 1,
            "قراط": 0,
            "سهم": 0,
            "نوع الاستخدام": "زراعي",
            "نوع المحصول": "قمح",
            "مراحل النمو": "نمو",
            "ملاحظات على الحيازة من فريق العمل": "ملاحظة",
            "__excel_row__": 2,
        }

    def test_empty_gis_and_boor_rule(self):
        base = Table(Path("base.xlsx"), ["كود القطعة", *BASE_TO_GIS.keys()], [self.base_row()])
        gis = Table(Path("gis.xlsx"), ["كود القطعة", *BASE_TO_GIS.values()], [{
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
        }])
        report = apply_update(base, gis, "كود القطعة", "كود القطعة", dict(BASE_TO_GIS))
        row = base.rows[0]
        self.assertIsNone(row["اسم الحوض"])
        self.assertEqual(row["كود الحوض"], "2")
        self.assertIsNone(row["نوع المحصول"])
        self.assertIsNone(row["مراحل النمو"])
        self.assertIsNone(row["ملاحظات على الحيازة من فريق العمل"])
        self.assertEqual(report.rule_changes, 2)

    def test_duplicates_are_not_updated(self):
        base = Table(Path("base.xlsx"), ["كود القطعة", "اسم الحوض"], [
            {"كود القطعة": "4", "اسم الحوض": "أ", "__excel_row__": 2},
            {"كود القطعة": "4", "اسم الحوض": "ب", "__excel_row__": 3},
        ])
        gis = Table(Path("gis.xlsx"), ["كود القطعة", "Basin_Name"], [
            {"كود القطعة": "4", "Basin_Name": "ج", "__excel_row__": 2},
        ])
        report = apply_update(base, gis, "كود القطعة", "كود القطعة", {"اسم الحوض": "Basin_Name"})
        self.assertEqual(report.base_duplicates, ["4"])
        self.assertEqual(base.rows[0]["اسم الحوض"], "أ")

    def test_key_normalization(self):
        self.assertEqual(key("0015"), "0015")
        self.assertEqual(key(15.0), "15")
        self.assertEqual(key("١٥"), "15")


if __name__ == "__main__":
    unittest.main()
