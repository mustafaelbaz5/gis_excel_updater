from __future__ import annotations

from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk

from .constants import SKIP_UPDATE_LABEL
from .exceptions import SpreadsheetValidationError
from .mapping import MappingResolver
from .models import MappingChoice, UpdateReport, WorkbookTable
from .reporting import format_summary
from .update_service import ParcelUpdateService
from .validation import MappingContentValidator
from .workbook_store import WorkbookStore


PALETTE = {
    "page": "#F4F7FA",
    "card": "#FFFFFF",
    "ink": "#16233A",
    "muted": "#617083",
    "primary": "#176B87",
    "primary_hover": "#0F526B",
    "success": "#14804A",
    "danger": "#B42318",
    "warning": "#B54708",
    "soft": "#E8F1F4",
}


class ScrollablePage(ttk.Frame):
    """Single scroll surface for the entire application content."""

    def __init__(self, master: tk.Misc):
        super().__init__(master)
        self.canvas = tk.Canvas(self, bg=PALETTE["page"], highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.content = ttk.Frame(self.canvas, style="Page.TFrame")
        self.window_id = self.canvas.create_window((0, 0), window=self.content, anchor="nw")

        self.content.bind("<Configure>", self._sync_scroll_region)
        self.canvas.bind("<Configure>", self._sync_width)
        self.canvas.bind("<Enter>", self._activate_wheel)
        self.canvas.bind("<Leave>", self._deactivate_wheel)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

    def _sync_scroll_region(self, _event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _sync_width(self, event):
        self.canvas.itemconfigure(self.window_id, width=event.width)

    def _activate_wheel(self, _event=None):
        self.canvas.bind_all("<MouseWheel>", self._scroll_mousewheel)

    def _deactivate_wheel(self, _event=None):
        self.canvas.unbind_all("<MouseWheel>")

    def _scroll_mousewheel(self, event):
        self.canvas.yview_scroll(int(-event.delta / 120), "units")


class HiyazaUpdaterUI(tk.Tk):
    def __init__(
        self,
        workbook_store: WorkbookStore,
        mapping_resolver: MappingResolver,
        update_service: ParcelUpdateService,
        mapping_validator: MappingContentValidator,
    ):
        super().__init__()
        self._store = workbook_store
        self._mapping_resolver = mapping_resolver
        self._update_service = update_service
        self._mapping_validator = mapping_validator
        self._base: WorkbookTable | None = None
        self._gis: WorkbookTable | None = None
        self._mapping_choices: list[MappingChoice] = []
        self._mapping_values: dict[str, tk.StringVar] = {}
        self._base_key = tk.StringVar()
        self._gis_key = tk.StringVar()

        self.title("مُحدِّث بيانات الحيازات من GIS")
        self.geometry("1170x820")
        self.minsize(980, 680)
        self.configure(bg=PALETTE["page"])
        self.option_add("*Font", ("Tahoma", 10))
        self._configure_style()
        self._build_screen()

    def _configure_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Page.TFrame", background=PALETTE["page"])
        style.configure("Card.TFrame", background=PALETTE["card"])
        style.configure("Title.TLabel", background=PALETTE["page"], foreground=PALETTE["ink"], font=("Tahoma", 20, "bold"))
        style.configure("PageText.TLabel", background=PALETTE["page"], foreground=PALETTE["muted"])
        style.configure("CardTitle.TLabel", background=PALETTE["card"], foreground=PALETTE["ink"], font=("Tahoma", 11, "bold"))
        style.configure("CardText.TLabel", background=PALETTE["card"], foreground=PALETTE["muted"])
        style.configure("Step.TLabel", background=PALETTE["soft"], foreground=PALETTE["primary"], font=("Tahoma", 9, "bold"), padding=(10, 5))
        style.configure("TButton", padding=(13, 9), font=("Tahoma", 10, "bold"))
        style.configure("Primary.TButton", background=PALETTE["primary"], foreground="white")
        style.map("Primary.TButton", background=[("active", PALETTE["primary_hover"]), ("disabled", "#A9C5CF")])
        style.configure("Success.TButton", background=PALETTE["success"], foreground="white")
        style.map("Success.TButton", background=[("active", "#0D633A"), ("disabled", "#A9D0B4")])
        style.configure("Treeview", rowheight=29, font=("Tahoma", 9))
        style.configure("Treeview.Heading", background=PALETTE["soft"], foreground=PALETTE["ink"], font=("Tahoma", 9, "bold"))
        style.configure("TNotebook", background=PALETTE["card"])
        style.configure("TNotebook.Tab", padding=(16, 8), font=("Tahoma", 9, "bold"))

    def _card(self, parent: tk.Misc, padding=18) -> ttk.Frame:
        return ttk.Frame(parent, style="Card.TFrame", padding=padding)

    def _build_screen(self):
        page = ScrollablePage(self)
        page.pack(fill="both", expand=True)
        content = page.content
        content.columnconfigure(0, weight=1)

        header = ttk.Frame(content, style="Page.TFrame", padding=(26, 22, 26, 12))
        header.grid(row=0, column=0, sticky="ew")
        ttk.Label(header, text="مُحدِّث بيانات الحيازات من GIS", style="Title.TLabel", anchor="e").pack(fill="x")
        ttk.Label(
            header,
            text="حمّل الملفين، راجع التركيب التلقائي، ثم أنشئ نسخة جديدة آمنة من الملف الأساسي.",
            style="PageText.TLabel",
            anchor="e",
        ).pack(fill="x", pady=(4, 0))

        steps = ttk.Frame(content, style="Page.TFrame", padding=(26, 0, 26, 14))
        steps.grid(row=1, column=0, sticky="ew")
        for title in ("1  اختيار الملفات", "2  مراجعة الأعمدة", "3  فحص وتحديث"):
            ttk.Label(steps, text=title, style="Step.TLabel").pack(side="right", padx=(8, 0))

        self._build_files_card(content).grid(row=2, column=0, sticky="ew", padx=26)
        self._build_mapping_card(content).grid(row=3, column=0, sticky="nsew", padx=26, pady=(14, 0))
        self._build_actions(content).grid(row=4, column=0, sticky="ew", padx=26, pady=(14, 0))
        self._build_report_card(content).grid(row=5, column=0, sticky="nsew", padx=26, pady=(14, 26))

    def _build_files_card(self, parent):
        card = self._card(parent)
        card.columnconfigure(1, weight=1)
        ttk.Label(card, text="اختيار الملفات", style="CardTitle.TLabel", anchor="e").grid(row=0, column=0, columnspan=3, sticky="ew")
        ttk.Label(
            card,
            text="يُقرأ شيت البيانات الوحيد تلقائيًا. لا تحتاج لاختيار Sheet من داخل الأداة.",
            style="CardText.TLabel",
            anchor="e",
        ).grid(row=1, column=0, columnspan=3, sticky="ew", pady=(4, 12))

        self._base_status = ttk.Label(card, text="لم يتم اختيار الملف الأساسي", style="CardText.TLabel", anchor="e")
        self._gis_status = ttk.Label(card, text="لم يتم اختيار ملف GIS", style="CardText.TLabel", anchor="e")
        self._file_row(card, 2, "الملف الأساسي", self._base_status, "اختيار الملف الأساسي", self._choose_base)
        self._file_row(card, 3, "ملف GIS", self._gis_status, "اختيار ملف GIS", self._choose_gis)
        return card

    def _file_row(self, parent, row, label, status, button_text, command):
        ttk.Label(parent, text=label, style="CardTitle.TLabel", anchor="e").grid(row=row, column=0, sticky="e", padx=(0, 14), pady=5)
        status.grid(row=row, column=1, sticky="ew", pady=5)
        ttk.Button(parent, text=button_text, command=command, style="Primary.TButton").grid(row=row, column=2, pady=5)

    def _build_mapping_card(self, parent):
        self._mapping_card = self._card(parent)
        ttk.Label(self._mapping_card, text="مراجعة ربط الأعمدة", style="CardTitle.TLabel", anchor="e").pack(fill="x")
        self._mapping_note = ttk.Label(
            self._mapping_card,
            text="بعد تحميل الملفين ستظهر هنا خانات الملف الأساسي وما يقابلها تلقائيًا في ملف GIS.",
            style="CardText.TLabel",
            anchor="e",
        )
        self._mapping_note.pack(fill="x", pady=(4, 12))
        self._mapping_body = ttk.Frame(self._mapping_card, style="Card.TFrame")
        self._mapping_body.pack(fill="x")
        return self._mapping_card

    def _build_actions(self, parent):
        card = self._card(parent, padding=14)
        self._update_button = ttk.Button(
            card,
            text="تحديث وإنشاء نسخة جديدة",
            command=self._update_and_save,
            style="Success.TButton",
            state="disabled",
        )
        self._update_button.pack(side="right")
        self._inspect_button = ttk.Button(
            card,
            text="فحص البيانات أولًا",
            command=self._inspect,
            style="Primary.TButton",
            state="disabled",
        )
        self._inspect_button.pack(side="right", padx=(0, 8))
        self._action_status = ttk.Label(card, text="الملف الأصلي محمي؛ الأداة تنشئ نسخة جديدة فقط.", style="CardText.TLabel", anchor="e")
        self._action_status.pack(side="right", fill="x", expand=True, padx=(14, 0))
        return card

    def _build_report_card(self, parent):
        card = self._card(parent)
        ttk.Label(card, text="التقرير", style="CardTitle.TLabel", anchor="e").pack(fill="x")
        self._report_note = ttk.Label(card, text="نفّذ الفحص لإظهار نتائج الربط قبل إنشاء النسخة الجديدة.", style="CardText.TLabel", anchor="e")
        self._report_note.pack(fill="x", pady=(4, 9))
        self._tabs = ttk.Notebook(card)
        self._tabs.pack(fill="both", expand=True)

        self._summary_text = self._text_tab("الملخص")
        self._base_unmatched_text = self._text_tab("غير مرتبطة في الأساسي")
        self._gis_unmatched_text = self._text_tab("غير مرتبطة في GIS")
        self._duplicates_text = self._text_tab("الأكواد المكررة")
        self._unresolved_crop_text = self._text_tab("أكواد محاصيل تحتاج مراجعة")
        self._changes_tree = self._changes_tab()
        return card

    def _text_tab(self, title):
        tab = ttk.Frame(self._tabs, style="Card.TFrame", padding=8)
        self._tabs.add(tab, text=title)
        tool_row = ttk.Frame(tab, style="Card.TFrame")
        tool_row.pack(fill="x", pady=(0, 5))
        content = tk.Text(tab, height=9, wrap="word", relief="flat", bg="#FBFCFE", fg=PALETTE["ink"], font=("Tahoma", 10))
        ttk.Button(tool_row, text="نسخ المحتوى", command=lambda: self._copy_text(content)).pack(side="right")
        content.pack(fill="both", expand=True)
        content.configure(state="disabled")
        return content

    def _changes_tab(self):
        tab = ttk.Frame(self._tabs, style="Card.TFrame", padding=8)
        self._tabs.add(tab, text="تفاصيل التغييرات")
        columns = ("code", "field", "old", "new", "reason")
        tree = ttk.Treeview(tab, columns=columns, show="headings")
        labels = ("كود القطعة", "الحقل", "القيمة القديمة", "القيمة الجديدة", "السبب")
        widths = (130, 210, 180, 180, 150)
        for column, label, width in zip(columns, labels, widths):
            tree.heading(column, text=label, anchor="e")
            tree.column(column, width=width, anchor="e")
        scrollbar = ttk.Scrollbar(tab, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        return tree

    def _choose_file(self):
        return filedialog.askopenfilename(
            title="اختيار ملف بيانات",
            filetypes=[
                ("ملفات Excel", "*.xlsx *.xls *.xlsm *.xlsb"),
                ("ملفات نصية", "*.csv *.tsv"),
                ("كل الملفات", "*.*"),
            ],
        )

    def _choose_base(self):
        path = self._choose_file()
        if path:
            self._load_table(path, is_base=True)

    def _choose_gis(self):
        path = self._choose_file()
        if path:
            self._load_table(path, is_base=False)

    def _load_table(self, path: str, is_base: bool):
        try:
            self._action_status.configure(text="جارِ قراءة الملف...")
            self.update_idletasks()
            table = self._store.read(path)
        except SpreadsheetValidationError as error:
            messagebox.showerror("تعذر قراءة الملف", str(error))
            self._action_status.configure(text="اختر ملفًا صالحًا ثم حاول مرة أخرى.")
            return

        if is_base:
            self._base = table
            self._base_status.configure(text=f"{Path(path).name} — {len(table.rows):,} صف")
        else:
            self._gis = table
            self._gis_status.configure(text=f"{Path(path).name} — {len(table.rows):,} صف")
        self._action_status.configure(text="تم تحميل الملف بنجاح.")
        if self._base and self._gis:
            self._show_mapping()

    def _show_mapping(self):
        for child in self._mapping_body.winfo_children():
            child.destroy()
        self._mapping_values.clear()
        self._mapping_choices, default_base_key, default_gis_key = self._mapping_resolver.resolve_defaults(self._base, self._gis)
        self._base_key.set(default_base_key or "")
        self._gis_key.set(default_gis_key or "")

        self._mapping_body.columnconfigure(0, weight=2)
        self._mapping_body.columnconfigure(1, weight=2)
        self._mapping_body.columnconfigure(2, weight=1)
        for column, label in enumerate(("الحقل في الأساسي", "العمود المقابل من GIS", "الحالة / مثال")):
            ttk.Label(self._mapping_body, text=label, style="CardTitle.TLabel", anchor="e").grid(row=0, column=column, sticky="ew", padx=8, pady=(0, 6))

        key_card = ttk.Frame(self._mapping_body, style="Card.TFrame")
        key_card.grid(row=1, column=0, columnspan=3, sticky="ew", padx=8, pady=(0, 8))
        key_card.columnconfigure(1, weight=1)
        key_card.columnconfigure(3, weight=1)
        ttk.Label(key_card, text="كود القطعة في الملف الأساسي", style="CardText.TLabel").grid(row=0, column=0, padx=(0, 6))
        ttk.Combobox(key_card, textvariable=self._base_key, values=self._base.headers, state="readonly").grid(row=0, column=1, sticky="ew", padx=(0, 16))
        ttk.Label(key_card, text="كود القطعة في ملف GIS", style="CardText.TLabel").grid(row=0, column=2, padx=(0, 6))
        ttk.Combobox(key_card, textvariable=self._gis_key, values=self._gis.headers, state="readonly").grid(row=0, column=3, sticky="ew")

        options = [SKIP_UPDATE_LABEL, *self._gis.headers]
        for row_number, choice in enumerate(self._mapping_choices, start=2):
            self._add_mapping_row(row_number, choice, options)

        self._mapping_note.configure(
            text="تم اقتراح الربط تلقائيًا. لو وجدت عمود GIS غير صحيح، بدّله من القائمة قبل الفحص."
        )
        self._inspect_button.configure(state="normal")
        self._update_button.configure(state="normal")
        self._clear_report()

    def _add_mapping_row(self, row_number, choice, options):
        value = tk.StringVar(value=choice.suggested_gis_header or SKIP_UPDATE_LABEL)
        self._mapping_values[choice.target_label] = value
        ttk.Label(self._mapping_body, text=choice.target_label, style="CardText.TLabel", anchor="e").grid(row=row_number, column=0, sticky="ew", padx=8, pady=4)
        ttk.Combobox(self._mapping_body, textvariable=value, values=options, state="readonly").grid(row=row_number, column=1, sticky="ew", padx=8, pady=4)
        state = ttk.Label(self._mapping_body, style="CardText.TLabel", anchor="e")
        state.grid(row=row_number, column=2, sticky="ew", padx=8, pady=4)

        def refresh_state(*_args):
            if not choice.base_header:
                state.configure(text="غير موجود في الأساسي", foreground=PALETTE["danger"])
            elif value.get() == SKIP_UPDATE_LABEL:
                state.configure(text="لن يتم تحديثه", foreground=PALETTE["warning"])
            else:
                example = self._mapping_validator.example_value(self._gis, value.get())
                state.configure(text=f"جاهز — مثال: {example}", foreground=PALETTE["success"])

        value.trace_add("write", refresh_state)
        refresh_state()

    def _selected_mapping(self):
        mapped = {}
        by_label = {choice.target_label: choice for choice in self._mapping_choices}
        for label, selected in self._mapping_values.items():
            choice = by_label[label]
            if choice.base_header:
                mapped[choice.base_header] = None if selected.get() == SKIP_UPDATE_LABEL else selected.get()
        return mapped

    def _has_valid_key_selection(self):
        if self._base_key.get() and self._gis_key.get():
            return True
        messagebox.showerror("كود القطعة", "اختر عمود كود القطعة في الملفين أولًا.")
        return False

    def _has_safe_mapping(self) -> bool:
        errors = self._mapping_validator.validate(self._base, self._gis, self._selected_mapping())
        if not errors:
            return True
        messagebox.showerror("راجع ربط الأعمدة", "\n\n".join(errors))
        return False

    def _inspect(self):
        if not self._has_valid_key_selection() or not self._has_safe_mapping():
            return
        try:
            report = self._update_service.update(
                self._base.clone_for_preview(),
                self._gis.clone_for_preview(),
                self._base_key.get(),
                self._gis_key.get(),
                self._selected_mapping(),
            )
        except SpreadsheetValidationError as error:
            messagebox.showerror("تعذر الفحص", str(error))
            return

        self._show_report(report, "تقرير فحص قبل التنفيذ")
        self._action_status.configure(text="اكتمل الفحص. يمكنك الآن إنشاء النسخة الجديدة.")
        
    def _update_and_save(self):
        if not self._has_valid_key_selection() or not self._has_safe_mapping():
            return
        if not messagebox.askyesno(
            "تأكيد إنشاء النسخة",
            "سيتم تحديث نسخة جديدة فقط. الملف الأساسي الأصلي لن يتم تعديله.",
        ):
            return

        try:
            self._action_status.configure(text="جارِ تحديث البيانات وحفظ النسخة الجديدة...")
            self.update_idletasks()
            report = self._update_service.update(
                self._base,
                self._gis,
                self._base_key.get(),
                self._gis_key.get(),
                self._selected_mapping(),
            )
            output = self._store.save_as_new_file(self._base)
        except (SpreadsheetValidationError, OSError) as error:
            messagebox.showerror("تعذر إنشاء النسخة", str(error))
            return

        self._show_report(report, "تقرير التحديث والحفظ")
        self._action_status.configure(text=f"تم إنشاء: {output.name}")
        messagebox.showinfo("تم التحديث", f"تم إنشاء النسخة الجديدة بنجاح:\n{output}")

    def _clear_report(self):
        for widget in (self._summary_text, self._base_unmatched_text, self._gis_unmatched_text, self._duplicates_text, self._unresolved_crop_text):
            widget.configure(state="normal")
            widget.delete("1.0", "end")
            widget.configure(state="disabled")
        for item in self._changes_tree.get_children():
            self._changes_tree.delete(item)
        self._report_note.configure(text="نفّذ الفحص لإظهار نتائج الربط قبل إنشاء النسخة الجديدة.")

    def _show_report(self, report: UpdateReport, title: str):
        self._report_note.configure(text=title)
        self._set_text(self._summary_text, format_summary(report))
        self._set_text(self._base_unmatched_text, "\n".join(report.base_without_match))
        self._set_text(self._gis_unmatched_text, "\n".join(report.gis_without_match))
        self._set_text(
            self._duplicates_text,
            "\n".join(("مكررة في الملف الأساسي:", *report.base_duplicates, "", "مكررة في ملف GIS:", *report.gis_duplicates)),
        )
        self._set_text(
            self._unresolved_crop_text,
            "\n".join(f"القطعة {parcel}: كود المحصول {code}" for parcel, code in report.unresolved_crop_codes),
        )
        for item in self._changes_tree.get_children():
            self._changes_tree.delete(item)
        for parcel, field, old, new, reason in report.details:
            self._changes_tree.insert("", "end", values=(parcel, field, self._value_for_display(old), self._value_for_display(new), reason))
        self._tabs.select(0)

    @staticmethod
    def _value_for_display(value):
        return "فارغ" if value is None or str(value).strip().casefold() == "null" else str(value)

    @staticmethod
    def _set_text(widget, content: str):
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("1.0", content or "لا توجد بيانات.")
        widget.configure(state="disabled")

    def _copy_text(self, widget):
        self.clipboard_clear()
        self.clipboard_append(widget.get("1.0", "end-1c"))
        self.update_idletasks()
        self._action_status.configure(text="تم نسخ المحتوى إلى الحافظة.")
