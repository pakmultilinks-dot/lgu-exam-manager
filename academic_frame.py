from academic_identity import student_count
"""Academic Course Records screen: preview, filter, validate, draft exports."""
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk

from academic_reports import generate_draft

GREEN = "#0d542b"
PALE = "#f5f7fa"
DARK = "#1a1a1a"
MUTED = "#555555"


class AcademicFrame(ctk.CTkFrame):
    def __init__(self, parent, app):
        super().__init__(parent, fg_color=PALE)
        self.app = app
        ctk.CTkLabel(self, text="Academic Course Records", font=("Arial", 23, "bold"),
                     text_color=DARK).pack(anchor="w", padx=24, pady=(8, 0))
        ctk.CTkLabel(self, text="One course attempt per row. Registration number identifies each student. "
                      "GPA = grade points / credit hours; best repeated course counted once.",
                     font=("Arial", 12), text_color=MUTED).pack(anchor="w", padx=24, pady=(4, 12))

        filter_bar = ctk.CTkFrame(self, fg_color="white", corner_radius=9)
        filter_bar.pack(fill="x", padx=24, pady=(0, 10))
        selectors = ctk.CTkFrame(filter_bar, fg_color="transparent")
        selectors.pack(fill="x", padx=16, pady=10)
        self.terms = ctk.CTkOptionMenu(selectors, values=["All Terms"], width=120,
                                        command=lambda _: self.reset_page())
        self.degrees = ctk.CTkOptionMenu(selectors, values=["All Degrees"], width=145,
                                          command=lambda _: self.reset_page())
        self.cohorts = ctk.CTkOptionMenu(selectors, values=["All Cohorts"], width=155,
                                          command=lambda _: self.reset_page())
        for name, menu in (("Attempt term", self.terms), ("Degree", self.degrees),
                           ("Admission session", self.cohorts)):
            ctk.CTkLabel(selectors, text=name, text_color=DARK).pack(side="left", padx=(4, 6))
            menu.pack(side="left", padx=(0, 15))

        self.search_var = ctk.StringVar(value="")
        self._search_after = None
        self.search_var.trace_add("write", self.debounce_search)
        search_bar = ctk.CTkFrame(self, fg_color="transparent")
        search_bar.pack(fill="x", padx=24, pady=(0, 8))
        ctk.CTkLabel(search_bar, text="Search", text_color=DARK).pack(side="left", padx=(0, 8))
        ctk.CTkEntry(search_bar, textvariable=self.search_var, width=430,
                     placeholder_text="Registration, student, course name or course code").pack(side="left")

        self.summary = ctk.CTkLabel(self, text="Import a course-level workbook to begin.",
                                     font=("Arial", 12, "bold"), text_color=GREEN)
        self.summary.pack(anchor="w", padx=26, pady=(0, 8))
        pages=ctk.CTkFrame(self,fg_color="transparent")
        pages.pack(fill="x",padx=24,pady=(0,6))
        self.prev_button=ctk.CTkButton(pages,text="Previous 150",width=110,command=lambda:self.move_page(-1))
        self.prev_button.pack(side="left",padx=4)
        self.next_button=ctk.CTkButton(pages,text="Next 150",width=110,command=lambda:self.move_page(1))
        self.next_button.pack(side="left",padx=4)
        ctk.CTkButton(pages,text="Review data / source rows",width=190,command=self.show_validation).pack(side="left",padx=12)

        preview_box = ctk.CTkFrame(self, fg_color="white")
        preview_box.pack(fill="both", expand=True, padx=24, pady=(0, 10))
        cols = ("registration", "name", "term", "degree", "code", "course", "crh", "marks", "grade", "gp")
        table_holder = ttk.Frame(preview_box)
        table_holder.pack(fill="both", expand=True, padx=10, pady=10)
        table_holder.rowconfigure(0, weight=1)
        table_holder.columnconfigure(0, weight=1)
        self.tree = ttk.Treeview(table_holder, columns=cols, show="headings", height=15)
        titles = ["Registration", "Student", "Term", "Degree", "Course Code", "Course Name", "CH", "Marks", "Grade", "Source GP"]
        widths = [170, 190, 62, 95, 105, 245, 54, 72, 65, 83]
        for col, title, width in zip(cols, titles, widths, strict=True):
            self.tree.heading(col, text=title)
            self.tree.column(col, width=width, minwidth=40)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vbar = ttk.Scrollbar(table_holder, orient="vertical", command=self.tree.yview)
        hbar = ttk.Scrollbar(table_holder, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)
        vbar.grid(row=0, column=1, sticky="ns")
        hbar.grid(row=1, column=0, sticky="ew")

        actions = ctk.CTkFrame(self, fg_color="transparent")
        actions.pack(fill="x", padx=24, pady=(0, 14))
        self.validate_button = ctk.CTkButton(actions, text="Review data / source rows",
                                             fg_color=GREEN, command=self.show_validation)
        self.validate_button.pack(side="left", padx=(0, 10))
        self.export_buttons = []
        export_row=ctk.CTkFrame(self,fg_color="transparent")
        export_row.pack(fill="x",padx=24,pady=(0,6))
        for label, fmt, color in (("Export Draft PDF", "pdf", "#b23b3b"),
                                  ("Export Draft Word", "word", "#2465a5"),
                                  ("Export Draft Excel", "excel", "#237742")):
            b = ctk.CTkButton(export_row, text=label, fg_color=color,
                              command=lambda f=fmt: self.export(f))
            b.pack(side="left", padx=5)
            self.export_buttons.append(b)
        ctk.CTkButton(actions, text="Mapping", width=85, command=self.show_mapping).pack(side="left",padx=4)
        self.notice = ctk.CTkLabel(self, text="GPA uses imported history. Excel includes exact totals and repeat audit. Official standing remains unassigned.",
                                   text_color="#962a25", font=("Arial", 11, "bold"))
        self.notice.pack(anchor="w", padx=25, pady=(0, 6))
        self.page = 0
        self.refresh()

    def reset_page(self):
        self.page = 0
        self.refresh_rows()

    def move_page(self, delta):
        self.page = max(0, self.page + delta)
        self.refresh_rows()

    def debounce_search(self, *_):
        if self._search_after is not None:
            self.after_cancel(self._search_after)
        self._search_after = self.after(220, self.reset_page)

    def on_show(self):
        self.refresh()

    def refresh(self):
        ds = self.app.academic_dataset
        if ds is None:
            for item in self.tree.get_children():
                self.tree.delete(item)
            self.summary.configure(text="No course-level Excel workbook loaded yet.")
            for b in self.export_buttons:
                b.configure(state="disabled")
            self.validate_button.configure(state="disabled")
            return
        self.validate_button.configure(state="normal")
        for menu, placeholder, items in ((self.terms, "All Terms", ds.terms),
                                         (self.degrees, "All Degrees", ds.degrees),
                                         (self.cohorts, "All Cohorts", ds.cohorts)):
            old = menu.get()
            menu.configure(values=[placeholder] + items)
            menu.set(old if old in items else placeholder)
        self.refresh_rows()

    def chosen_courses(self, *, ignore_term=False):
        ds = self.app.academic_dataset
        if not ds:
            return []
        rows = ds.filtered(term='All Terms' if ignore_term else self.terms.get(), degree=self.degrees.get(), cohort=self.cohorts.get())
        q = self.search_var.get().strip().casefold()
        if q:
            rows = [c for c in rows if q in (c.registration + " " + c.name + " " +
                                               c.course_code + " " + c.course_name).casefold()]
        return rows

    def refresh_rows(self):
        ds = self.app.academic_dataset
        for item in self.tree.get_children():
            self.tree.delete(item)
        if ds is None:
            return
        records = self.chosen_courses()
        # Preview 300 records to keep a large workbook responsive. Exports use all.
        self.page = min(self.page, max(0, (len(records)-1)//150))
        self.prev_button.configure(state="normal" if self.page else "disabled")
        self.next_button.configure(state="normal" if (self.page+1)*150 < len(records) else "disabled")
        for c in records[self.page*150:(self.page+1)*150]:
            self.tree.insert("", "end", values=(c.registration, c.name, c.term, c.degree,
                                                   c.course_code, c.course_name, c.credits,
                                                   c.marks, c.grade, c.quality_points))
        counts = ds.counts()
        from validation_scope import relevant_errors
        errors=relevant_errors(ds,records,history=True)
        self.summary.configure(text=(
            f"{student_count(records)} students | {len(records)} course attempts | "
            f"page {self.page+1}/{max(1,(len(records)+149)//150)} (150 rows/page)"))
        info=sum(i.level=='INFO' for i in ds.issues)
        self.notice.configure(text=f"Workbook review: {counts['errors']} errors, {counts['warnings']} warnings, {info} information notices. "
                                  f"{len(errors)} errors affect this selection/history. {counts['placeholders']} empty rows skipped.")
        can_export = len(records)>0 and not errors
        for b in self.export_buttons:
            b.configure(state="normal" if can_export else "disabled")

    def show_mapping(self):
        self.app.frames['import'].review_workbook(self.app.source_path)

    def show_validation(self):
        from review_ui import open_review
        open_review(self,self.app.academic_dataset,self.chosen_courses())

    def export(self, fmt):
        from course_ui import operation_available,dataset_snapshot
        if not operation_available(self.app):return
        ds = dataset_snapshot(self.app.academic_dataset)
        if ds is None:
            return
        records = self.chosen_courses(ignore_term=True)
        if not records:
            messagebox.showwarning("No records", "No records match these filters.")
            return
        ext = {"pdf": ".pdf", "word": ".docx", "excel": ".xlsx"}[fmt]
        filetypes = {"pdf": "PDF", "word": "Word", "excel": "Excel"}
        chosen = filedialog.asksaveasfilename(
            title="Save draft register - NOT official gazette",
            defaultextension=ext, initialfile="LGU_DRAFT_Result_Register" + ext,
            filetypes=[(filetypes[fmt], "*" + ext)])
        if not chosen:
            return
        settings=dict(self.app.report_settings)
        self.app.run_job("Exporting draft register",
                         lambda: generate_draft(fmt, ds, records, chosen, template_key=settings.get("gazette_template","auto"),settings=settings),
                         lambda _: messagebox.showinfo("Draft exported", "Draft saved: " + chosen))
