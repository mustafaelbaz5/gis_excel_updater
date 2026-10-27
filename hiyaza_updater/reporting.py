from .models import UpdateReport


def format_summary(report: UpdateReport) -> str:
    lines = [
        "ملخص التحديث",
        f"إجمالي صفوف الملف الأساسي: {report.base_total}",
        f"إجمالي صفوف ملف GIS: {report.gis_total}",
        f"القطع المرتبطة والمؤهلة للتحديث: {report.eligible_matches}",
        f"القطع التي تغيّر بها حقل واحد على الأقل: {report.updated_rows}",
        f"إجمالي الخلايا التي تغيّرت: {report.changed_cells}",
        f"تعديلات بسبب قاعدة بور/مباني: {report.rule_changes}",
        "",
        "التغييرات حسب العمود:",
    ]
    if report.changes_by_field:
        lines.extend(
            f"- {field}: {count} تحديث، منها {report.blank_changes_by_field[field]} تحولت إلى فراغ"
            for field, count in report.changes_by_field.items()
        )
    else:
        lines.append("- لا توجد تغييرات")
    lines.extend([
        "",
        f"أكواد الأساسي غير المرتبطة: {len(report.base_without_match)}",
        f"أكواد GIS غير المرتبطة: {len(report.gis_without_match)}",
        f"أكواد مكررة في الأساسي: {len(report.base_duplicates)}",
        f"أكواد مكررة في GIS: {len(report.gis_duplicates)}",
        f"صفوف بكود فارغ في الأساسي: {report.base_blank_keys}",
        f"صفوف بكود فارغ في GIS: {report.gis_blank_keys}",
        f"أكواد نوع المحصول بلا وصف Domain: {len(report.unresolved_crop_codes)}",
    ])
    return "\n".join(lines)
