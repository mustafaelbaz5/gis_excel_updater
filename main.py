from hiyaza_updater.mapping import MappingResolver
from hiyaza_updater.ui import HiyazaUpdaterUI
from hiyaza_updater.update_service import ParcelUpdateService
from hiyaza_updater.workbook_store import WorkbookStore


def main() -> None:
    workbook_store = WorkbookStore()
    mapping_resolver = MappingResolver()
    update_service = ParcelUpdateService(workbook_store, mapping_resolver)
    application = HiyazaUpdaterUI(workbook_store, mapping_resolver, update_service)
    application.mainloop()


if __name__ == "__main__":
    main()
