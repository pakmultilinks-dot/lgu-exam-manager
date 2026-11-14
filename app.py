from user_errors import friendly_error
from academic_identity import student_count,program_key
"""LGU Examination Department - Result Management System (desktop app)."""
import queue
import threading
import time
import logging
import os
import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
import pandas as pd

import importer
import reports
from dashboard import compute_stats, pass_percentage_label
from sample_data import save_sample
from academic_engine import detect_layouts, load_course_workbook
from academic_frame import AcademicFrame

# ---- Premium light theme: official university software ----
LGU_GREEN = "#0d542b"
LGU_GREEN_DARK = "#0a3d20"
LGU_GREEN_LIGHT = "#e8f5e9"
LGU_GOLD = "#b8860b"
LGU_GOLD_DARK = "#9a6f0a"
LGU_GOLD_LIGHT = "#fef6e0"
BG = "#f5f7fa"              # main background: soft light gray
PANEL = "#ffffff"           # cards: pure white
CARD_BG = "#ffffff"
SIDEBAR_BG = "#0d542b"      # deep green sidebar
TEXT_DARK = "#1a1a1a"
TEXT_MID = "#555555"
TEXT_LIGHT = "#999999"
BORDER = "#e0e0e0"
SUCCESS = "#1b8c3e"
DANGER = "#c0392b"
WARNING = "#b8860b"
# Format button colors
PDF_RED = "#c0392b"
PDF_RED_HOVER = "#96281b"
WORD_BLUE = "#2b6cb0"
WORD_BLUE_HOVER = "#1f4e8c"
EXCEL_GREEN = "#1e8e3e"
EXCEL_GREEN_HOVER = "#146c2e"

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("green")

# Navigation with icons
NAV_ITEMS = [
    ("import", "  \u25c8  Import Excel", "Upload & map result sheets"),
    ("data", "  \u2630  Data Grid", "Browse student records"),
    ("academic", "  \u25a6  Academic Records", "Course attempts and draft registers"),
    ("dashboard", "  \u25a3  Dashboard", "Results at a glance"),
    ("reports", "  \u25a4  Reports", "Legacy report center"),
    ("settings", "  \u2699  Settings", "App preferences"),
]

# One-line descriptions for the Report Center cards
REPORT_DESCRIPTIONS = {
    "gazette": "Legacy report for one-row-per-student spreadsheets; manual review required.",
    "lgu_official": "Locked until the university-approved gazette template and grading rules arrive.",
    "cgpa_analysis": "Grade distribution, averages and performance bands.",
    "pass_fail": "Overall pass / fail counts and success percentage.",
    "merit": "Top 20 position holders ranked by grade.",
    "failed": "Students with failing status and their grades.",
    "cards": "Printable individual result card for every student.",
    "subjects": "Average, highest and lowest marks per subject.",
    "notification": "Formal pass notification for official circulation.",
    "stats_breakdown": "Mean, median, quartiles and standard deviation per subject.",
    "subject_toppers": "Top 3 scorers in each subject.",
    "certificates": "Merit certificates for the position holders.",
    "transcript": "Detailed marks certificate (DMC) per student.",
    "notice": "One-page class summary for the notice board.",
    "failing_subjects": "Failed students' subjects below your Settings passing marks.",
    "supplementary": "Student counts and list by status, as per source data.",
    "probation": "CGPA ranking, or probation list when you set a cutoff in Settings.",
    "subject_pass_pct": "Per-subject pass % with your Settings rule, or marks summary.",
    "data_validation": "Find data-entry issues before results are published.",
    "subject_difficulty": "Per-subject averages, spread and charts. No judgments.",
    "student_profile": "One page per student: marks table, chart and class rank.",
    "session_comparison": "Compare with a previous session file you load below.",
}


def enable_trackpad_scroll(scrollable_frame):
    """Make CTkScrollableFrame respond to laptop trackpad two-finger scroll
    AND regular mouse wheel on all platforms.

    Trackpads send the same MouseWheel events as a mouse, but the events
    get swallowed by child widgets. This binds scrolling at the frame level
    and forwards it to the internal canvas.
    """
    canvas = getattr(scrollable_frame, "_parent_canvas",
                     getattr(scrollable_frame, "_canvas", None))
    if canvas is None:
        return  # unknown customtkinter internals; leave default behavior

    def _on_mousewheel(event):
        # Windows: event.delta in multiples of 120. Trackpads send smaller deltas.
        # Normalize so both mouse wheel and trackpad feel right.
        delta = event.delta
        if delta == 0:
            return
        # One "click" = 120 units; scroll proportionally for smooth trackpad feel
        steps = int(-delta / 60)  # 60 = half-click for smoother trackpad
        if steps == 0:
            steps = -1 if delta > 0 else 1
        canvas.yview_scroll(steps, "units")
        return "break"  # stop the event from propagating further

    def _bind_recursive(widget):
        widget.bind("<MouseWheel>", _on_mousewheel, add="+")
        widget.bind("<Button-4>", lambda e: (canvas.yview_scroll(-1, "units"), "break")[1], add="+")
        widget.bind("<Button-5>", lambda e: (canvas.yview_scroll(1, "units"), "break")[1], add="+")
        for child in widget.winfo_children():
            _bind_recursive(child)

    _bind_recursive(scrollable_frame)
    # Re-bind when new children are added (e.g. mapping rows, report cards)
    scrollable_frame.bind("<Configure>", lambda e: _bind_recursive(scrollable_frame), add="+")


class App(ctk.CTk):
    def __init__(self,restore_autosave=False):
        super().__init__()
        self.title("LGU Exam Manager - Gazette v4.5.2")
        self.geometry("1280x800")
        self.minsize(1100, 700)

        # State
        self.raw_df = None          # original imported dataframe
        self.df = None              # normalized working frame
        self.mapping = {}           # role -> column(s)
        self.source_path = None
        self.stats = {}
        self.auto_save_imports = restore_autosave
        self.import_pending = None
        self.academic_dataset = None  # course-level mode; cannot use legacy reports
        # Multi-department support: when the imported file has a Department
        # column, the Data tab shows a dropdown to filter the grid, the
        # dashboard and every report to a single department. "All Departments"
        # keeps the previous single-department behaviour.
        self.selected_department = "All Departments"
        # Session Comparison (v1.4): the previous-session file the office
        # loads in the Report Center. Never invented; always user-chosen.
        self.compare_df = None
        self.compare_mapping = None
        self.compare_label = None
        # User-configured report setting (Settings tab). The app NEVER
        # invents thresholds; the exam department sets its own rules.
        # Only topper_count remains: there is no fixed passing mark per
        # subject (relative marking), so pass_marks was removed.
        # department/semester/session/declaration_date feed the official
        # LGU gazette cover page (user-configured, never invented).
        self.report_settings = {
            "gazette_template": "auto",
            "topper_count": 3,      # position holders for certificates
            "department": "",
            "semester": "",
            "session": "",
            "declaration_date": "",
            # Operator-set rules (empty = no rule; the app never invents
            # thresholds). Used by Failing Subjects Detail, Subject-wise
            # Pass Percentage (passing_marks) and Academic Probation List
            # (probation_cgpa). LGU uses relative marking.
            "passing_marks": "",
            "probation_cgpa": "",
        }

        self._default_report_settings=dict(self.report_settings)

        self._jobs = queue.Queue()
        self.job_busy = False
        log_dir=os.path.join(os.path.expanduser("~"),".lgu-exam-manager")
        os.makedirs(log_dir,exist_ok=True)
        from logging.handlers import RotatingFileHandler
        handler=RotatingFileHandler(os.path.join(log_dir,"application.log"),maxBytes=1000000,backupCount=2,encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | build=4.5.2 | %(message)s"))
        logger=logging.getLogger("lgu")
        for previous in list(logger.handlers):
            logger.removeHandler(previous);previous.close()
        logger.addHandler(handler)
        logger.setLevel(logging.INFO)
        logger.info("Application starting")
        self._build_shell()
        self.after(75, self._poll_jobs)
        self.show_frame("import")
        if restore_autosave:self._restore_autosave()
        self.bind("<Control-o>", lambda e:self.frames["import"].browse())
        self.bind("<Control-d>", lambda e:self.show_frame("dashboard"))
        self.bind("<Control-r>", lambda e:self.show_frame("reports"))
        self.bind("<Control-s>", lambda e:self.save_session())
        self.bind("<Control-f>", lambda e:self.show_frame("academic"))
        for key,fmt in (("1","pdf"),("2","word"),("3","excel")):
            self.bind("<Control-Key-"+key+">",lambda e,f=fmt:self.frames["academic"].export(f))

    def run_job(self, label, work, done):
        if self.job_busy:
            messagebox.showinfo("Operation in progress", "Wait for the current operation to finish.")
            return False
        self.job_busy = True
        self.title("LGU Exam Manager - Gazette v4.5.2 - " + label)
        def worker():
            started = time.perf_counter()
            try:
                result = work()
                self._jobs.put((label, done, result, None, time.perf_counter()-started))
            except Exception as exc:
                # Log operation and exception type only; source cells may be in exception text.
                frames=[(os.path.basename(f.filename),f.lineno,f.name) for f in __import__("traceback").extract_tb(exc.__traceback__)]
                logging.getLogger("lgu").error("Operation failed: %s / %s; frames=%s", label, type(exc).__name__,frames)
                error=f"{type(exc).__name__}: {str(exc).strip() or 'No error details were supplied.'}"
                self._jobs.put((label, done, None, error, time.perf_counter()-started))
        threading.Thread(target=worker, daemon=True).start()
        return True

    def _poll_jobs(self):
        try:
            while True:
                label, done, result, error, elapsed = self._jobs.get_nowait()
                self.job_busy = False
                self.title("LGU Exam Manager - Gazette v4.5.2")
                if error is not None:
                    from user_errors import friendly_error
                    display_error=friendly_error(error)
                    if getattr(self,'import_pending',None):
                        self.frames['import'].status_lbl.configure(text='Import failed. The previous workspace is still active. '+display_error)
                        self.import_pending=None;self.update_header_status()
                    logging.getLogger("lgu").error("Operation result: %s; elapsed=%.3fs; success=False",label,elapsed)
                    messagebox.showerror("Operation failed", label+"\n\n"+display_error)
                else:
                    try:
                        done(result)
                        logging.getLogger("lgu").info("Operation completed: %s; elapsed=%.3fs; success=True",label,elapsed)
                    except Exception as exc:
                        logging.getLogger("lgu").error("Operation completion failed: %s; success=False",label)
                        self.report_callback_exception(type(exc), exc, exc.__traceback__)
        except queue.Empty:
            pass
        self.after(75, self._poll_jobs)

    def report_callback_exception(self, exc_type, exc, tb):
        logging.getLogger("lgu").error("UI callback failed: %s; frames=%s",exc_type.__name__,[(os.path.basename(f.filename),f.lineno,f.name) for f in __import__("traceback").extract_tb(tb)])
        messagebox.showerror("Application error", "The operation could not complete. Your saved files are preserved. Details are available in the support log.")

    # ---------------- shell ----------------
    def _build_shell(self):
        # --- Top header bar (64px, LGU green) ---
        self.header = ctk.CTkFrame(self, height=64, corner_radius=0, fg_color=LGU_GREEN)
        self.header.pack(side="top", fill="x")
        self.header.pack_propagate(False)

        header_inner = ctk.CTkFrame(self.header, fg_color="transparent")
        header_inner.pack(fill="both", expand=True, padx=24, pady=10)

        # Brand block: official LGU logo (falls back to gold "LGU" badge)
        brand = ctk.CTkFrame(header_inner, fg_color="transparent")
        brand.pack(side="left", fill="y")
        try:
            from PIL import Image
            _logo_file = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                      "lgu-logo.jpg")
            _pil = Image.open(_logo_file).convert("RGB")
            self._logo_ctk = ctk.CTkImage(light_image=_pil, size=(44, 44))
            ctk.CTkLabel(brand, text="", image=self._logo_ctk,
                         width=48, height=44).pack(side="left", padx=(0, 14))
        except Exception:
            badge = ctk.CTkLabel(brand, text="LGU", font=("Arial", 20, "bold"),
                                 width=56, height=40, corner_radius=8,
                                 fg_color=LGU_GOLD, text_color="white")
            badge.pack(side="left", padx=(0, 14))
        brand_text = ctk.CTkFrame(brand, fg_color="transparent")
        brand_text.pack(side="left")
        ctk.CTkLabel(brand_text, text="Lahore Garrison University",
                     font=("Arial", 14, "bold"), text_color="white").pack(anchor="w")
        ctk.CTkLabel(brand_text, text="Examination Department  \u00b7  Result Management System",
                     font=("Arial", 10), text_color="#c8e6c9").pack(anchor="w")

        # Right side: one-click save + live record-count status pill
        header_right = ctk.CTkFrame(header_inner, fg_color="transparent")
        header_right.pack(side="right")
        self.save_btn = ctk.CTkButton(
            header_right, text="Save", fg_color=LGU_GOLD,
            hover_color=LGU_GOLD_DARK, text_color="white",
            font=("Arial", 12, "bold"), height=34, width=90, corner_radius=8,
            command=self.save_session)
        self.save_btn.pack(side="left", padx=(0, 12), pady=4)
        status_pill = ctk.CTkFrame(header_right, fg_color=LGU_GREEN_DARK,
                                   corner_radius=16)
        status_pill.pack(side="left", pady=4)
        self.header_status = ctk.CTkLabel(status_pill, text="No file loaded",
                                          font=("Arial", 11, "bold"),
                                          text_color="#c8e6c9")
        self.header_status.pack(padx=16, pady=6)

        # --- Body: sidebar + main ---
        self.body = ctk.CTkFrame(self, fg_color=BG, corner_radius=0)
        self.body.pack(side="top", fill="both", expand=True)

        # --- Sidebar (240px, deep green) ---
        self.sidebar = ctk.CTkFrame(self.body, width=240, corner_radius=0, fg_color=SIDEBAR_BG)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        ctk.CTkLabel(self.sidebar, text="NAVIGATION", font=("Arial", 10, "bold"),
                     text_color="#81c784").pack(pady=(24, 10), padx=24, anchor="w")

        self.nav_buttons = {}
        for key, label, _desc in NAV_ITEMS:
            b = ctk.CTkButton(self.sidebar, text=label, fg_color="transparent",
                              hover_color=LGU_GREEN_DARK, anchor="w",
                              font=("Arial", 13), height=46, corner_radius=8,
                              text_color="white",
                              command=lambda k=key: self.show_frame(k))
            b.pack(fill="x", padx=14, pady=3)
            self.nav_buttons[key] = b

        # Sidebar bottom spacer (keeps nav at the top)
        ctk.CTkLabel(self.sidebar, text="", font=("Arial", 4)).pack(expand=True, fill="both")

        self.main = ctk.CTkFrame(self.body, fg_color=BG, corner_radius=0)
        self.main.pack(side="right", fill="both", expand=True)

        self.frames = {}
        for key, cls in [("import", ImportFrame), ("data", DataFrame_),
                         ("academic", AcademicFrame), ("dashboard", DashboardFrame), ("reports", ReportsFrame),
                         ("settings", SettingsFrame)]:
            f = cls(self.main, self)
            self.frames[key] = f

    def show_frame(self, key):
        if key == "data" and self.academic_dataset is not None:
            key = "academic"
        for k, f in self.frames.items():
            f.pack_forget()
            self.nav_buttons[k].configure(fg_color="transparent")
        self.frames[key].pack(fill="both", expand=True, padx=24, pady=20)
        self.nav_buttons[key].configure(fg_color=LGU_GREEN_DARK)
        self.frames[key].on_show()
        self.update_header_status()

    def update_header_status(self):
        """Three-state status pill:
        gray  - no file chosen yet
        amber - file chosen, mapping not finished / data not loaded yet
        green - data loaded into the app
        """
        imp = self.frames.get("import")
        file_chosen = bool(imp is not None and imp.path)
        if self.import_pending:
            self.header_status.configure(text='Loading '+self.import_pending+'...',text_color='#ffe082')
            return
        if self.academic_dataset is not None:
            n = self.academic_dataset.counts()
            self.header_status.configure(
                text=f"{self.academic_dataset.source_filename}\n{n['students']} students | {n['course_attempts']} course attempts",
                text_color="#a5d6a7")
        elif self.raw_df is not None:
            self.header_status.configure(
                text=f"\u25cf {len(self.raw_df)} records loaded",
                text_color="#a5d6a7")
        elif file_chosen:
            self.header_status.configure(
                text="\u25d0 File selected",
                text_color=LGU_GOLD)
        else:
            self.header_status.configure(
                text="\u25cb No file selected",
                text_color="#9e9e9e")

    # ---------------- shared helpers ----------------
    def get_filtered_df(self):
        """Return the working dataframe scoped to the selected department.

        Falls back to the full frame when no department data exists (files
        without a Department column behave exactly as before).
        """
        if self.df is None:
            return None
        sel = getattr(self, "selected_department", "All Departments")
        if (sel and sel != "All Departments"
                and "_department" in self.df.columns):
            fdf = self.df[self.df["_department"] == sel].copy()
            # attrs may not survive filtering on older pandas; re-attach them.
            for k, v in self.df.attrs.items():
                if k not in fdf.attrs:
                    fdf.attrs[k] = v
            return fdf
        return self.df

    def department_values(self):
        """Sorted unique non-empty department names, or [] if none."""
        if self.df is None or "_department" not in self.df.columns:
            return []
        vals = (self.df["_department"].dropna().astype(str).str.strip()
                .replace("", pd.NA).dropna().unique().tolist())
        return sorted(vals)

    def clear_loaded_data(self):
        """Never leave previous students visible after clearing or changing source."""
        self.raw_df = None
        self.df = None
        self.academic_dataset = None
        self.mapping = {}
        self.source_path = None
        self.stats = {}
        self.selected_department = "All Departments"
        self.compare_df=None;self.compare_mapping=None;self.compare_label=None
        self.import_pending=None
        for key in ("academic", "data", "dashboard", "reports", "settings"):
            self.frames[key].refresh()
        self.update_header_status()

    def set_academic_data(self, dataset, source_path, reset_selection=False):
        """Course-level data must never flow into legacy student-row reports."""
        self.academic_dataset = dataset
        self.raw_df = None
        self.df = None
        self.stats = {}
        self.mapping = {}
        self.source_path = source_path
        self.selected_department = "All Departments"
        if reset_selection:
            academic=self.frames['academic']
            academic.terms.set('All Terms');academic.degrees.set('All Degrees');academic.cohorts.set('All Cohorts')
            if academic.search_var.get():academic.search_var.set('')
            if academic._search_after is not None:
                academic.after_cancel(academic._search_after);academic._search_after=None
            academic.page=0
            panel=getattr(self.frames['reports'],'course_panel',None)
            if panel:
                panel.program_menu.set('All programs / degrees')
                from course_reports import ROUTES
                panel.route_menu.set(ROUTES['register'])
        self.frames["academic"].refresh()
        self.update_header_status()

    def set_data(self, raw_df, mapping, source_path):
        self.academic_dataset = None
        self.frames["academic"].refresh()
        self.raw_df = raw_df
        self.mapping = mapping
        self.source_path = source_path
        self.df = importer.apply_mapping(raw_df, mapping)
        # Attach user-configured thresholds (never invented by the app)
        self.df.attrs["settings"] = dict(self.report_settings)
        self.stats = compute_stats(self.df)
        self.selected_department = "All Departments"
        for key in ("data", "dashboard", "reports"):
            self.frames[key].refresh()

    def ask_pdf_path(self, default):
        return filedialog.asksaveasfilename(defaultextension=".pdf",
                                            initialfile=default,
                                            filetypes=[("PDF", "*.pdf")])

    # ---------------- session save / auto-restore ----------------
    # One-click Save in the header bar. Saved as a .lguexam ZIP containing
    # data.xlsx + mapping.json (deliberately NOT pickle: pickle can execute
    # arbitrary code on load). Loaded only when explicitly requested.
    def _autosave_path(self):
        d = os.path.join(os.path.expanduser("~"), ".lgu-exam-manager")
        os.makedirs(d, exist_ok=True)
        return os.path.join(d, "autosave.lguexam")

    def save_session(self):
        if self.raw_df is None and self.academic_dataset is None:
            messagebox.showinfo("Nothing to save", "Import a workbook first.")
            return
        import session_store
        path=self._autosave_path()
        mode="course" if self.academic_dataset is not None else "wide"
        source=self.source_path; raw=self.raw_df; mapping=dict(self.mapping); settings=dict(self.report_settings)
        dataset=self.academic_dataset
        self.run_job("Saving session",lambda:session_store.save(path,mode,source,raw,mapping,settings,dataset),
                     lambda _:messagebox.showinfo("Saved", "Workspace saved. New launches start empty."))

    def open_saved_workspace(self):
        if self.job_busy:
            messagebox.showinfo("Operation in progress","Wait for the current operation to finish.")
            return False
        if not os.path.isfile(self._autosave_path()):
            messagebox.showinfo("No saved workspace","Save an imported workbook first. Startup remains empty until you import or open saved records.")
            return False
        return self._restore_autosave()

    def _restore_autosave(self):
        path=self._autosave_path()
        if not os.path.isfile(path):return
        import session_store
        def done(result):
            meta,data=result
            self.frames['import'].path=None
            self.frames['import']._sheet_override=None
            self.frames['import']._header_override=None
            self.report_settings=dict(self._default_report_settings)
            if isinstance(meta.get("settings"),dict):self.report_settings.update(meta["settings"])
            if meta.get("mode")=="course":
                from gazette_options import resolve_template_groups
                try:resolve_template_groups(self.report_settings.get('gazette_template','auto'),data.courses)
                except ValueError:self.report_settings['gazette_template']='auto'
                self.set_academic_data(data,None,reset_selection=True)
                self.frames['import'].file_lbl.configure(text=data.source_filename)
                self.frames['import'].status_lbl.configure(text=f'Restored saved workspace: {data.source_filename} ({len(data.courses)} attempts). Import Excel to replace it.')
                self.show_frame("import")
            else:
                self.set_data(data,meta.get("mapping",{}),None)
                self.show_frame("data")
        return self.run_job("Opening saved workspace",lambda:session_store.restore(path),done)


# ---------------- Import frame: Smart Import (fully automatic) ----------------
class ImportFrame(ctk.CTkFrame):
    """Recognize course tables or review their structure before committing records."""
    def __init__(self, parent, app):
        super().__init__(parent, fg_color=BG)
        self.app = app
        self.path = None
        self.sheets = []
        self.df = None
        # Manual overrides (None = auto-detect)
        self._sheet_override = None
        self._header_override = None
        self._adv_open = False

        # Page title
        ctk.CTkLabel(self, text="Import Result File",
                     font=("Arial", 22, "bold"),
                     text_color=TEXT_DARK).pack(anchor="w", padx=24, pady=(4, 2))
        ctk.CTkLabel(self,
                     text="Choose an Excel result workbook. Recognized tables load automatically; "
                          "unfamiliar columns open the import review screen.",
                     font=("Arial", 12), text_color=TEXT_MID).pack(anchor="w",
                                                                  padx=24, pady=(0, 16))

        # Main card
        card = ctk.CTkFrame(self, fg_color=PANEL, corner_radius=12,
                            border_width=1, border_color=BORDER)
        card.pack(fill="x", padx=24, pady=8)
        picker = ctk.CTkFrame(card, fg_color="#f0f4f8", corner_radius=10)
        picker.pack(fill="x", padx=20, pady=20)
        inner = ctk.CTkFrame(picker, fg_color="transparent")
        inner.pack(fill="x", padx=16, pady=18)
        ctk.CTkLabel(inner, text="\u21e7", font=("Arial", 30),
                     text_color=LGU_GREEN).pack(side="left", padx=(4, 16))
        btn_col = ctk.CTkFrame(inner, fg_color="transparent")
        btn_col.pack(side="left")
        ctk.CTkButton(btn_col, text="Browse Excel File", fg_color=LGU_GREEN,
                      hover_color=LGU_GREEN_DARK, height=44, corner_radius=8,
                      font=("Arial", 13, "bold"),
                      command=self.browse).pack(anchor="w")
        self.file_lbl = ctk.CTkLabel(inner, text="No file selected yet",
                                     font=("Arial", 11, "italic"),
                                     text_color=TEXT_LIGHT)
        self.file_lbl.pack(side="left", padx=24)

        # Status line (what the app detected / any problem)
        self.status_lbl = ctk.CTkLabel(card, text="",
                                       font=("Arial", 12),
                                       text_color=TEXT_MID, wraplength=900,
                                       justify="left")
        self.status_lbl.pack(anchor="w", padx=20, pady=(0, 4))

        actions=ctk.CTkFrame(card,fg_color="transparent")
        actions.pack(fill="x",padx=20,pady=(8,16))
        ctk.CTkButton(actions,text="Review sheets and columns",command=self.review_workbook,fg_color=LGU_GREEN).pack(side="left",padx=(0,10))
        self.open_saved_button=ctk.CTkButton(actions,text="Open saved workspace",command=self.app.open_saved_workspace,fg_color=LGU_GREEN)
        self.open_saved_button.pack(side="left",padx=(0,10))
        self.clear_button=ctk.CTkButton(actions,text="Clear current data",command=self.clear_all,fg_color=LGU_GREEN)
        self.clear_button.pack(side="left")

    def on_show(self):
        pass

    def clear_all(self):
        """Reset the import screen to its initial state."""
        if self.app.job_busy:
            messagebox.showinfo("Operation in progress","Wait for the current operation to finish before clearing records.")
            return False
        self.path = None
        self.sheets = []
        self.df = None
        self._sheet_override = None
        self._header_override = None
        self.file_lbl.configure(text="No file selected yet",
                                font=("Arial", 11, "italic"),
                                text_color=TEXT_LIGHT)
        self.status_lbl.configure(text="")
        self.app.clear_loaded_data()

    # ---------------- Smart Import pipeline ----------------
    def browse(self):
        if self.app.job_busy:
            messagebox.showinfo("Operation running", "Wait for the current operation to finish.")
            return
        paths = filedialog.askopenfilenames(filetypes=[("Excel workbooks", "*.xlsx *.xlsm *.xls *.xlsb")])
        if not paths:return
        self._sheet_override = None
        self._header_override = None
        if len(paths)==1:
            self._smart_load(paths[0])
            return
        from academic_engine import load_course_workbooks
        from workspace_state import import_settings,persist_import
        from workbook_schema import inspect_workbook,needs_review
        self.app.import_pending=f'{len(paths)} workbooks';self.app.update_header_status()
        self.status_lbl.configure(text="Reading workbook structures...")
        previous=dict(self.app.report_settings)
        autosave=self.app._autosave_path() if self.app.auto_save_imports else None
        def load(plans):
            dataset=load_course_workbooks(paths,plans);settings=import_settings(dataset,previous);error=None
            if autosave:
                try:persist_import(autosave,None,dataset,settings)
                except Exception as exc:error=f'{type(exc).__name__}: {exc}'
            return dataset,settings,error
        def done(result):
            dataset,settings,error=result
            self.app.import_pending=None;self.app.report_settings=settings
            self.app.set_academic_data(dataset,None,reset_selection=True)
            self.file_lbl.configure(text=f"{len(paths)} workbooks imported")
            self.status_lbl.configure(text=f'Loaded {len(dataset.courses)} attempts from {len(paths)} workbooks.'+(' Workspace saved automatically.' if autosave and not error else ''))
            if error:messagebox.showwarning('Workspace not saved','The workbooks are loaded; use Save before closing.')
            self.app.show_frame("academic")
        def prepare():
            plans={};pending=[]
            for path in paths:
                try:layouts=detect_layouts(path)
                except ValueError as exc:
                    if 'Ambiguous duplicate headings' not in str(exc):raise
                    layouts=[]
                inspection=inspect_workbook(path,layouts)
                if not inspection:raise ValueError('One of the selected workbooks has no populated worksheets.')
                if needs_review(inspection,layouts):pending.append((path,inspection))
                else:plans[str(path)]=layouts
            return plans,pending
        def prepared(result):
            plans,pending=result
            self.app.import_pending=None;self.app.update_header_status()
            def next_file(index):
                if index==len(pending):
                    self.app.import_pending=f'{len(paths)} workbooks';self.app.update_header_status()
                    self.app.run_job('Importing workbooks',lambda:load(plans),done);return
                from import_review import WorkbookImportReview
                path,inspection=pending[index]
                self.status_lbl.configure(text=f'Review workbook {index+1} of {len(pending)}. No active records have been replaced.')
                def accepted(layouts):plans[str(path)]=layouts;next_file(index+1)
                self.import_review=WorkbookImportReview(self.app,path,inspection,accepted)
            next_file(0)
        self.app.run_job('Reading workbook structures',prepare,prepared)

    def make_sample(self):
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                            "sample_result.xlsx")
        save_sample(path)
        self._sheet_override = None
        self._header_override = None
        self.app.clear_loaded_data()
        self._smart_load(path)

    def _smart_load(self, path, layouts_override=None, reviewed=False):
        if self.app.job_busy:
            messagebox.showinfo('Operation in progress','Wait for the current operation to finish before importing another file.')
            return False
        self.app.import_pending=os.path.basename(path);self.app.update_header_status()
        self.path = path
        self.file_lbl.configure(text=os.path.basename(path))
        self.status_lbl.configure(text="Reading and validating workbook in background...")
        selected = self._sheet_override
        header_override = self._header_override
        previous_settings=dict(self.app.report_settings)
        autosave_path=self.app._autosave_path() if self.app.auto_save_imports else None
        def completed(result):
            from workspace_state import import_settings,persist_import
            dataset=result[1] if result[0]=='course' else None
            settings=import_settings(dataset,previous_settings);save_error=None
            if autosave_path:
                try:persist_import(autosave_path,path,dataset,settings,raw=result[1] if dataset is None else None,mapping=result[2] if dataset is None else None)
                except Exception as exc:save_error=friendly_error(f'{type(exc).__name__}: {exc}')
            return result,settings,save_error
        def work():
            from workbook_schema import inspect_workbook,needs_review
            try:layouts=layouts_override if layouts_override is not None else detect_layouts(path,selected_sheet=selected)
            except ValueError as exc:
                if 'Ambiguous duplicate headings' not in str(exc):raise
                layouts=[]
            if not reviewed:
                inspection=inspect_workbook(path,layouts)
                if not inspection:raise ValueError('This workbook contains no populated worksheets.')
                if needs_review(inspection,layouts):
                    # Only recognize a summary workbook from an explicit GPA field.
                    legacy=None
                    if not layouts and len(inspection)==1:
                        item=inspection[0]
                        df=importer.read_sheet(path,item['sheet'],item['header_row'])
                        mapping=importer.suggest_mapping(df)
                        label=str(mapping.get('cgpa','')).lower()
                        if mapping.get('roll') and mapping.get('name') and 'gpa' in label and not any(role in item['mapping'] for role in ('course_code','course_name','credits')):legacy=(df,mapping)
                    if legacy:return completed(('wide',*legacy))
                    return (('review',inspection),previous_settings,None)
            if layouts:
                return completed(("course", load_course_workbook(path, layouts_override=layouts)))
            sheets = importer.load_workbook(path)
            sheet = selected or self._auto_sheet(path, sheets)
            header = header_override if header_override is not None else self._auto_header_row(path, sheet)
            df = importer.read_sheet(path, sheet, header_row=header)
            mapping = importer.suggest_mapping(df)
            if not mapping.get("roll") or not mapping.get("name"):
                raise ValueError("Needs mapping: could not identify registration and student name. Review File options and source headings.")
            return completed(("wide", df, mapping))
        def done(payload):
            result,settings,save_error=payload
            self.app.import_pending=None
            if result[0]=='review':
                from import_review import WorkbookImportReview
                self.app.update_header_status()
                self.status_lbl.configure(text='Review this workbook structure before importing. The active workspace has not changed.')
                self.import_review=WorkbookImportReview(self.app,path,result[1],lambda layouts:self._smart_load(path,layouts,True))
                return
            self.app.report_settings=settings
            if result[0] == "course":
                ds = result[1]
                self.app.set_academic_data(ds, path,reset_selection=True)
                self.status_lbl.configure(text=f"Loaded {ds.source_filename}: {len(ds.students)} students / {len(ds.courses)} attempts. "+('Workspace saved automatically.' if autosave_path and not save_error else ''))
                self.app.show_frame("academic")
            else:
                self.app.set_data(result[1], result[2], path)
                self.status_lbl.configure(text=f"Imported {len(result[1])} student rows.")
                self.app.show_frame("data")
            if save_error:
                self.status_lbl.configure(text='Workbook loaded, but automatic workspace save failed: '+save_error)
                messagebox.showwarning('Workspace not saved','The workbook is loaded. Use Save before closing.\n\n'+save_error)
        return self.app.run_job("Importing workbook", work, done)

    def review_workbook(self,path=None):
        if self.app.job_busy:return
        path=path or self.path
        if not path:path=filedialog.askopenfilename(filetypes=[('Excel workbooks','*.xlsx *.xlsm *.xls *.xlsb')])
        if not path:return
        from workbook_schema import inspect_workbook
        from import_review import WorkbookImportReview
        def done(inspection):
            if not inspection:
                messagebox.showwarning('No worksheet data','This workbook has no nonempty worksheets to map.');return
            self.import_review=WorkbookImportReview(self.app,path,inspection,lambda layouts:self._smart_load(path,layouts,True))
        self.app.run_job('Reading workbook structure',lambda:inspect_workbook(path),done)

    def _fail(self, msg):
        self.status_lbl.configure(text=msg, text_color=DANGER)
        messagebox.showwarning("Import", msg)

    def _auto_sheet(self, path, sheets):
        """Pick the worksheet holding the most data (rows x columns,
        measured after header-row detection so title rows don't win)."""
        if len(sheets) == 1:
            return sheets[0]
        best, best_score = sheets[0], -1
        for s in sheets[:8]:
            try:
                hr = self._auto_header_row(path, s)
                df = importer.read_sheet(path, s, header_row=hr)
                score = len(df) * len(df.columns)
            except Exception:
                continue
            if score > best_score:
                best, best_score = s, score
        return best

    @staticmethod
    def _looks_numeric(text):
        try:
            float(str(text).replace(",", "").strip())
            return True
        except (ValueError, TypeError):
            return False

    def _auto_header_row(self, path, sheet):
        """Row (of rows 1-4) that looks like column titles: mostly
        non-numeric text cells. Prefers the candidate with the most
        columns (a lone title cell like "My University" loses to the real
        header); ties go to the earliest row."""
        candidates = []
        for hr in range(0, 4):
            try:
                df = importer.read_sheet(path, sheet, header_row=hr)
            except Exception:
                continue
            cols = [str(c).strip() for c in df.columns if str(c).strip()]
            if not cols:
                continue
            text_like = sum(1 for c in cols if not self._looks_numeric(c))
            if text_like >= max(1, int(len(cols) * 0.5)):
                candidates.append((hr, len(cols)))
        if not candidates:
            return 0
        multi = [c for c in candidates if c[1] >= 2]
        pool = multi or candidates
        pool.sort(key=lambda c: (-c[1], c[0]))
        return pool[0][0]

    # ---------------- File options (manual override) ----------------
    def _toggle_advanced(self):
        self.review_workbook()

    def _on_option_change(self,*_):
        self.review_workbook()


# ---------------- Data grid frame ----------------
class DataFrame_(ctk.CTkFrame):
    def __init__(self, parent, app):
        super().__init__(parent, fg_color=BG)
        self.app = app
        self.sort_col = None
        self.sort_rev = False

        # Toolbar card: title + search + status filter
        toolbar = ctk.CTkFrame(self, fg_color=PANEL, corner_radius=12,
                               border_width=1, border_color=BORDER)
        toolbar.pack(fill="x", padx=24, pady=(4, 8))
        title_row = ctk.CTkFrame(toolbar, fg_color="transparent")
        title_row.pack(fill="x", padx=20, pady=(14, 8))
        ctk.CTkLabel(title_row, text="Student Records",
                     font=("Arial", 20, "bold"),
                     text_color=TEXT_DARK).pack(side="left")
        self.count_lbl = ctk.CTkLabel(title_row, text="", font=("Arial", 12),
                                      text_color=TEXT_MID)
        self.count_lbl.pack(side="left", padx=14)

        filter_row = ctk.CTkFrame(toolbar, fg_color="transparent")
        filter_row.pack(fill="x", padx=20, pady=(0, 14))
        ctk.CTkLabel(filter_row, text="Search",
                     font=("Arial", 12, "bold"),
                     text_color=TEXT_DARK).pack(side="left", padx=(0, 8))
        self.search_var = ctk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.refresh())
        ctk.CTkEntry(filter_row, textvariable=self.search_var, width=340,
                     height=36, corner_radius=8,
                     placeholder_text="Search any column\u2026",
                     border_width=1, border_color=BORDER,
                     fg_color="white", text_color=TEXT_DARK).pack(side="left",
                                                                 padx=4)
        ctk.CTkLabel(filter_row, text="Status",
                     font=("Arial", 12, "bold"),
                     text_color=TEXT_DARK).pack(side="left", padx=(24, 8))
        self.status_var = ctk.StringVar(value="All")
        ctk.CTkOptionMenu(filter_row, variable=self.status_var,
                          values=["All", "Pass", "Fail"], width=130,
                          height=36, corner_radius=8,
                          fg_color="white", text_color=TEXT_DARK,
                          button_color=LGU_GREEN,
                          command=lambda _: self.refresh()).pack(side="left",
                                                                padx=4)
        # Department filter: only shown when the data has a Department column
        self.dept_lbl = ctk.CTkLabel(filter_row, text="Department",
                                     font=("Arial", 12, "bold"),
                                     text_color=TEXT_DARK)
        self.dept_var = ctk.StringVar(value="All Departments")
        self.dept_menu = ctk.CTkOptionMenu(filter_row, variable=self.dept_var,
                                           values=["All Departments"],
                                           width=200, height=36,
                                           corner_radius=8,
                                           fg_color="white",
                                           text_color=TEXT_DARK,
                                           button_color=LGU_GREEN,
                                           command=self._on_dept_change)

        # Light professional table styling
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background="white", foreground=TEXT_DARK,
                        fieldbackground="white", rowheight=30,
                        font=("Arial", 11), borderwidth=0)
        style.configure("Treeview.Heading", background=LGU_GREEN,
                        foreground="white", font=("Arial", 11, "bold"),
                        relief="flat")
        style.map("Treeview",
                  background=[("selected", "#fff3cd")],
                  foreground=[("selected", TEXT_DARK)])

        table_card = ctk.CTkFrame(self, fg_color=PANEL, corner_radius=12,
                                  border_width=1, border_color=BORDER)
        table_card.pack(fill="both", expand=True, padx=24, pady=(0, 4))
        table_inner = ctk.CTkFrame(table_card, fg_color="transparent")
        table_inner.pack(fill="both", expand=True, padx=12, pady=12)
        self.tree = ttk.Treeview(table_inner, show="headings")
        self.tree.tag_configure("even", background="white")
        self.tree.tag_configure("odd", background="#f0f6f1")
        vsb = ttk.Scrollbar(table_inner, orient="vertical",
                            command=self.tree.yview)
        hsb = ttk.Scrollbar(table_inner, orient="horizontal",
                            command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

    def on_show(self):
        self.refresh()

    def _on_dept_change(self, _choice):
        """Department dropdown changed: scope grid, dashboard and reports."""
        self.app.selected_department = self.dept_var.get()
        self.refresh()
        # Keep the dashboard in sync if it is the visible tab
        try:
            self.app.frames["dashboard"].refresh()
        except Exception:
            pass

    def _view(self):
        """Raw dataframe with search + status + department filters applied."""
        df = self.app.raw_df
        q = self.search_var.get().strip().lower()
        view = df
        if q:
            mask = df.apply(lambda r: r.astype(str).str.lower().str.contains(q, na=False, regex=False)).any(axis=1)
            view = df[mask]
        if self.status_var.get() != "All" and self.app.mapping.get("status"):
            scol = self.app.mapping["status"]
            norm = importer.normalize_status(df[scol])
            view = view[norm.loc[view.index] == self.status_var.get()]
        sel = self.dept_var.get()
        dcol = self.app.mapping.get("department") if self.app.mapping else None
        if dcol and sel and sel != "All Departments" and dcol in df.columns:
            dnorm = df[dcol].astype(str).str.strip()
            view = view[dnorm.loc[view.index] == sel]
        return view

    def _sync_dept_menu(self):
        """Show the department dropdown only when department data exists."""
        depts = self.app.department_values()
        if depts:
            self.dept_lbl.pack(side="left", padx=(24, 8))
            self.dept_menu.pack(side="left", padx=4)
            values = ["All Departments"] + depts
            self.dept_menu.configure(values=values)
            if self.dept_var.get() not in values:
                self.dept_var.set("All Departments")
                self.app.selected_department = "All Departments"
        else:
            self.dept_lbl.pack_forget()
            self.dept_menu.pack_forget()
            self.dept_var.set("All Departments")
            self.app.selected_department = "All Departments"

    def _insert_rows(self, view):
        for idx, (_, r) in enumerate(view.iterrows()):
            tag = "odd" if idx % 2 else "even"
            self.tree.insert("", "end",
                             values=[("" if pd.isna(v) else v) for v in r],
                             tags=(tag,))

    def refresh(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        if self.app.raw_df is None:
            self.count_lbl.configure(text="No data imported yet")
            return
        self._sync_dept_menu()
        df = self.app.raw_df
        view = self._view()
        cols = list(view.columns)
        self.tree["columns"] = cols
        for c in cols:
            self.tree.heading(c, text=c,
                              command=lambda cc=c: self.sort_by(cc))
            self.tree.column(c, width=130, anchor="center")
        self._insert_rows(view)
        sel = self.dept_var.get()
        scope = f" ({sel})" if sel != "All Departments" else ""
        self.count_lbl.configure(
            text=f"Showing {len(view)} of {len(df)} rows{scope}")

    def sort_by(self, col):
        if self.sort_col == col:
            self.sort_rev = not self.sort_rev
        else:
            self.sort_col, self.sort_rev = col, False
        df = self.app.raw_df
        # re-apply current filters then sort
        view = self._view()
        try:
            view = view.sort_values(col, ascending=not self.sort_rev,
                                    kind="mergesort")
        except Exception:
            pass
        for i in self.tree.get_children():
            self.tree.delete(i)
        self._insert_rows(view)
        sel = self.dept_var.get()
        scope = f" ({sel})" if sel != "All Departments" else ""
        self.count_lbl.configure(
            text=f"Showing {len(view)} of {len(df)} rows{scope}")


# ---------------- Dashboard frame ----------------
class DashboardFrame(ctk.CTkFrame):
    # Icon + tint per card type
    CARD_ICONS = [
        ("Total Students", "\u25c8", LGU_GREEN, LGU_GREEN_LIGHT),
        ("Passed", "\u2713", SUCCESS, "#e6f4ea"),
        ("Failed", "\u2717", DANGER, "#fdecea"),
        ("Pass %", "%", LGU_GOLD, LGU_GOLD_LIGHT),
        ("Avg", "\u25b2", LGU_GREEN, LGU_GREEN_LIGHT),
        ("Highest", "\u2605", LGU_GOLD, LGU_GOLD_LIGHT),
        ("Lowest", "\u25bc", TEXT_MID, "#eef0f3"),
    ]

    def __init__(self, parent, app):
        super().__init__(parent, fg_color=BG)
        self.app = app
        ctk.CTkLabel(self, text="Result Dashboard",
                     font=("Arial", 22, "bold"),
                     text_color=TEXT_DARK).pack(anchor="w", padx=24, pady=(4, 2))
        ctk.CTkLabel(self, text="Key performance indicators for the loaded result set.",
                     font=("Arial", 12),
                     text_color=TEXT_MID).pack(anchor="w", padx=24, pady=(0, 8))
        self.cards = ctk.CTkFrame(self, fg_color="transparent")
        self.cards.pack(fill="x", padx=16, pady=6)

        # Merit list as a clean styled table inside a white card
        self.merit_box = ctk.CTkFrame(self, fg_color=PANEL, corner_radius=12,
                                     border_width=1, border_color=BORDER)
        self.merit_box.pack(fill="both", expand=True, padx=24, pady=8)
        ctk.CTkLabel(self.merit_box, text="Top 10 Merit List",
                     font=("Arial", 15, "bold"),
                     text_color=TEXT_DARK).pack(anchor="w", padx=20,
                                               pady=(16, 8))
        merit_inner = ctk.CTkFrame(self.merit_box, fg_color="transparent")
        merit_inner.pack(fill="both", expand=True, padx=20, pady=(0, 16))
        self.merit_tree = ttk.Treeview(merit_inner, show="headings",
                                       columns=("rank", "roll", "name", "grade"),
                                       height=10)
        self.merit_tree.heading("rank", text="#")
        self.merit_tree.heading("roll", text="Roll Number")
        self.merit_tree.heading("name", text="Student Name")
        self.merit_tree.heading("grade", text="Grade")
        self.merit_tree.column("rank", width=50, anchor="center")
        self.merit_tree.column("roll", width=150, anchor="center")
        self.merit_tree.column("name", width=280, anchor="w")
        self.merit_tree.column("grade", width=100, anchor="center")
        self.merit_tree.tag_configure("even", background="white")
        self.merit_tree.tag_configure("odd", background="#f0f6f1")
        self.merit_tree.pack(fill="both", expand=True)

    def on_show(self):
        self.refresh()

    def _icon_for(self, title):
        for key, icon, color, tint in self.CARD_ICONS:
            if title.startswith(key):
                return icon, color, tint
        return "\u25c8", TEXT_MID, "#eef0f3"

    def _card(self, title, value, color=LGU_GREEN):
        icon, icon_color, tint = self._icon_for(title)
        if title in ("Passed",):
            icon_color, tint = SUCCESS, "#e6f4ea"
        elif title in ("Failed",):
            icon_color, tint = DANGER, "#fdecea"
        elif title.startswith("Pass %"):
            icon_color, tint = LGU_GOLD, LGU_GOLD_LIGHT
        f = ctk.CTkFrame(self.cards, fg_color=PANEL, corner_radius=12,
                         border_width=1, border_color=BORDER)
        f.pack(side="left", padx=8, pady=6, fill="x", expand=True)
        icon_lbl = ctk.CTkLabel(f, text=icon, font=("Arial", 16, "bold"),
                                width=42, height=42, corner_radius=21,
                                fg_color=tint, text_color=icon_color)
        icon_lbl.pack(pady=(16, 8))
        ctk.CTkLabel(f, text=str(value), font=("Arial", 26, "bold"),
                     text_color=color).pack()
        ctk.CTkLabel(f, text=title, font=("Arial", 11),
                     text_color=TEXT_MID).pack(pady=(2, 6))
        return f

    def refresh(self):
        for w in self.cards.winfo_children():
            w.destroy()
        for i in self.merit_tree.get_children():
            self.merit_tree.delete(i)
        ds = self.app.academic_dataset
        if ds is not None:
            rows = self.app.frames["academic"].chosen_courses()
            self.merit_box.pack_forget()
            metrics=(("Distinct students",student_count(rows)),
                     ("Course attempts",len(rows)),("Attempt terms",len({c.term for c in rows})),
                     ("Admission cohorts",len({c.admission_session for c in rows})),
                     ("Programs",len({program_key(c.degree) for c in rows})),("Errors requiring review",ds.counts()["errors"]))
            self.metric_canvas=tk.Canvas(self.cards,height=330,bg=BG,highlightthickness=0)
            self.metric_canvas.pack(fill="x",expand=True)
            def draw_metrics(event=None):
                canvas=self.metric_canvas
                width=max(600,canvas.winfo_width());cell=(width-36)/3
                canvas.delete("all")
                for i,(label,value) in enumerate(metrics):
                    x=6+(i%3)*(cell+12);y=8+(i//3)*156
                    canvas.create_rectangle(x,y,x+cell,y+140,fill=PANEL,outline=BORDER,tags="metric")
                    canvas.create_text(x+cell/2,y+50,text=str(value),font=("Arial",28,"bold"),fill=LGU_GREEN,tags="metric")
                    canvas.create_text(x+cell/2,y+103,text=label,font=("Arial",12),fill=TEXT_MID,tags="metric")
            self.metric_canvas.bind("<Configure>",draw_metrics)
            self.metric_canvas.tag_bind("metric","<Button-1>",lambda e:self.app.show_frame("academic"))
            draw_metrics()
            ctk.CTkButton(self.cards,text="Filter / search / inspect",command=lambda:self.app.show_frame("academic")).pack(anchor="w",padx=8,pady=10)
            counts=ds.counts()
            info=sum(issue.level=="INFO" for issue in ds.issues)
            ctk.CTkLabel(self.cards,text=f"Workbook review: {counts['errors']} errors, {counts['warnings']} warnings, {info} information notices. Only relevant errors block a selected result.",wraplength=850,text_color=TEXT_MID).pack(anchor="w",padx=8,pady=8)
            return
        if not self.merit_box.winfo_manager():
            self.merit_box.pack(fill="both",expand=True,padx=24,pady=8)
        if self.app.raw_df is None:
            ctk.CTkLabel(self.cards, text="Import data to see statistics.",
                         font=("Arial", 12),
                         text_color=TEXT_MID).pack(pady=20)
            return
        # Department-aware: stats follow the department filter in the Data tab
        fdf = self.app.get_filtered_df()
        s = compute_stats(fdf if fdf is not None else self.app.df)
        self._card("Total Students", s["total"], TEXT_DARK)
        if s.get("has_passfail"):
            self._card("Passed", s["passes"], SUCCESS)
            self._card("Failed", s["fails"], DANGER)
            pct_card = self._card(pass_percentage_label(s), f"{s['pass_pct']}%", LGU_GOLD)
            bar = ctk.CTkProgressBar(pct_card, width=150, height=8,
                                     corner_radius=4, fg_color="#eef0f3",
                                     progress_color=LGU_GOLD)
            bar.set(max(0.0, min(1.0, s["pass_pct"] / 100.0)))
            bar.pack(pady=(0, 14))
        else:
            self._card("Passed", "N/A", TEXT_LIGHT)
            self._card("Failed", "N/A", TEXT_LIGHT)
            self._card("Pass %", "N/A", TEXT_LIGHT)
        if s.get("has_cgpa"):
            gl = s.get("grade_label", "CGPA")
            self._card(f"Avg {gl}", s["avg_cgpa"] if s["avg_cgpa"] is not None else "-",
                       TEXT_DARK)
            self._card(f"Highest {gl}", s["max_cgpa"] if s["max_cgpa"] is not None else "-",
                       LGU_GOLD)
            self._card(f"Lowest {gl}", s["min_cgpa"] if s["min_cgpa"] is not None else "-",
                       TEXT_MID)
        else:
            self._card("Avg Grade", "N/A", TEXT_LIGHT)
            self._card("Highest Grade", "N/A", TEXT_LIGHT)
            self._card("Lowest Grade", "N/A", TEXT_LIGHT)
        for i, m in enumerate(s["merit"], 1):
            tag = "odd" if i % 2 == 0 else "even"
            self.merit_tree.insert("", "end",
                                   values=(i, m["roll"], m["name"], m["cgpa"]),
                                   tags=(tag,))


# ---------------- Reports frame: Report Center ----------------
class ReportsFrame(ctk.CTkFrame):
    # Format button colors: PDF red, Word blue, Excel green
    FORMAT_STYLES = {
        "pdf": (PDF_RED, PDF_RED_HOVER),
        "word": (WORD_BLUE, WORD_BLUE_HOVER),
        "excel": (EXCEL_GREEN, EXCEL_GREEN_HOVER),
    }

    def __init__(self, parent, app):
        super().__init__(parent, fg_color=BG)
        self.app = app
        self.legacy_built=False

    def _build_legacy(self):
        ctk.CTkLabel(self, text="Report Center",
                     font=("Arial", 22, "bold"),
                     text_color=TEXT_DARK).pack(anchor="w", padx=24, pady=(4, 2))
        ctk.CTkLabel(self,
                     text="Available reports export in PDF, Word and Excel. "
                          "Publication and academic decisions require approved rules.",
                     font=("Arial", 12), text_color=TEXT_MID).pack(anchor="w",
                                                                  padx=24,
                                                                  pady=(0, 8))
        self.grid_box = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.grid_box.pack(fill="both", expand=True, padx=14)
        enable_trackpad_scroll(self.grid_box)

        # ---- Batch Export All (v1.4): one click, every report, all formats
        batch_panel = ctk.CTkFrame(self.grid_box, fg_color=CARD_BG,
                                   corner_radius=12, border_width=1,
                                   border_color=BORDER)
        batch_panel.grid(row=0, column=0, columnspan=2, padx=10, pady=8,
                         sticky="ew")
        ctk.CTkLabel(batch_panel, text="Export Everything at Once",
                     font=("Arial", 14, "bold"),
                     text_color=TEXT_DARK).pack(anchor="w", padx=18,
                                                pady=(14, 2))
        ctk.CTkLabel(batch_panel,
                     text="Generates every available report in PDF, Word and "
                          "Excel into one dated folder, with a log of what was "
                          "created. The office's biggest time saver.",
                     font=("Arial", 11), text_color=TEXT_MID,
                     wraplength=640, anchor="w",
                     justify="left").pack(anchor="w", padx=18, pady=(0, 10))
        brow = ctk.CTkFrame(batch_panel, fg_color="transparent")
        brow.pack(anchor="w", padx=14, pady=(0, 8))
        self.batch_btn = ctk.CTkButton(
            brow, text="\u2b07 Export All Reports (PDF + Word + Excel)",
            fg_color=LGU_GREEN, hover_color=LGU_GREEN_DARK,
            width=340, height=40, corner_radius=10,
            font=("Arial", 12, "bold"), text_color="white",
            command=self.batch_export)
        self.batch_btn.pack(side="left", padx=4)
        self.batch_status = ctk.CTkLabel(brow, text="", font=("Arial", 11),
                                         text_color=TEXT_MID)
        self.batch_status.pack(side="left", padx=12)

        # ---- Session Comparison loader (v1.4)
        cmp_panel = ctk.CTkFrame(self.grid_box, fg_color=CARD_BG,
                                 corner_radius=12, border_width=1,
                                 border_color=BORDER)
        cmp_panel.grid(row=1, column=0, columnspan=2, padx=10, pady=8,
                       sticky="ew")
        ctk.CTkLabel(cmp_panel, text="Compare with a Previous Session",
                     font=("Arial", 14, "bold"),
                     text_color=TEXT_DARK).pack(anchor="w", padx=18,
                                                pady=(14, 2))
        ctk.CTkLabel(cmp_panel,
                     text="Load last semester's result file to unlock the "
                          "Session Comparison Report below. Subjects are "
                          "matched by name; anything that does not match is "
                          "listed, never guessed.",
                     font=("Arial", 11), text_color=TEXT_MID,
                     wraplength=640, anchor="w",
                     justify="left").pack(anchor="w", padx=18, pady=(0, 10))
        crow = ctk.CTkFrame(cmp_panel, fg_color="transparent")
        crow.pack(anchor="w", padx=14, pady=(0, 8))
        ctk.CTkButton(crow, text="Choose previous session file...",
                      fg_color=WORD_BLUE, hover_color=WORD_BLUE_HOVER,
                      width=240, height=36, corner_radius=8,
                      font=("Arial", 11, "bold"), text_color="white",
                      command=self.load_compare_file).pack(side="left", padx=4)
        ctk.CTkButton(crow, text="Clear",
                      fg_color="transparent", border_width=1,
                      border_color=BORDER, text_color=TEXT_MID,
                      width=90, height=36, corner_radius=8,
                      font=("Arial", 11),
                      command=self.clear_compare_file).pack(side="left", padx=4)
        self.compare_status = ctk.CTkLabel(crow, text="No previous session loaded.",
                                           font=("Arial", 11),
                                           text_color=TEXT_MID)
        self.compare_status.pack(side="left", padx=12)

        # One card per report: title + description + 3 format buttons
        self.cards = []  # (key, req, {fmt: button})
        for i, (key, (label, _handlers, req)) in enumerate(reports.REPORTS.items()):
            card = ctk.CTkFrame(self.grid_box, fg_color=CARD_BG, corner_radius=12,
                                border_width=1, border_color=BORDER)
            card.grid(row=(i // 2) + 2, column=i % 2, padx=10, pady=8,
                      sticky="nsew")
            ctk.CTkLabel(card, text=label, font=("Arial", 13, "bold"),
                         text_color=TEXT_DARK, wraplength=300,
                         anchor="w", justify="left").pack(anchor="w", padx=18,
                                                          pady=(16, 2))
            ctk.CTkLabel(card, text=REPORT_DESCRIPTIONS.get(key, ""),
                         font=("Arial", 11), text_color=TEXT_MID,
                         wraplength=300, anchor="w",
                         justify="left").pack(anchor="w", padx=18, pady=(0, 12))
            btn_row = ctk.CTkFrame(card, fg_color="transparent")
            btn_row.pack(anchor="w", padx=14, pady=(0, 16))
            fmt_btns = {}
            for fmt_key, (fmt_label, _ext) in reports.FORMATS.items():
                fg, hover = self.FORMAT_STYLES[fmt_key]
                b = ctk.CTkButton(btn_row, text=f"\u2b07 {fmt_label}",
                                  fg_color=fg, hover_color=hover,
                                  width=96, height=34, corner_radius=8,
                                  font=("Arial", 11, "bold"),
                                  text_color="white",
                                  command=lambda k=key, f=fmt_key: self.run_report(k, f))
                b.pack(side="left", padx=4)
                fmt_btns[fmt_key] = b
            self.cards.append((key, req, fmt_btns))
        self.grid_box.grid_columnconfigure(0, weight=1)
        self.grid_box.grid_columnconfigure(1, weight=1)
        self.hint_lbl = ctk.CTkLabel(self, text="", font=("Arial", 11),
                                     text_color=TEXT_MID, wraplength=700)
        self.hint_lbl.pack(pady=(4, 4))

    def on_show(self):
        self.refresh()

    def _req_status(self, key, req):
        """(ok, reason) for a report given the current data/mapping."""
        if key in ("lgu_official", "gazette", "notification"):
            return False, "Official gazette is locked pending approved template and grading rules."
        if self.app.academic_dataset is not None:
            from course_reports import capability
            return capability(key, self.app.academic_dataset, self.app.frames["academic"].chosen_courses())
        m = self.app.mapping or {}
        has_data = self.app.df is not None
        has_cgpa = bool(m.get("cgpa"))
        has_status = bool(m.get("status"))
        has_subjects = bool(m.get("subject"))
        if not has_data:
            return False, "Import data first."
        if key in ("failing_subjects", "subject_pass_pct") and not self.app.report_settings.get("passing_marks"):
            return False, "Needs policy: configured approved passing marks per subject."
        if key == "probation" and not self.app.report_settings.get("probation_cgpa"):
            return False, "Needs policy: configured approved probation cutoff."
        if req == "cgpa" and not has_cgpa:
            return False, "Needs a mapped grade column."
        if req == "passfail" and not (has_status or has_cgpa):
            return False, "Needs a mapped Status or grade column."
        if req == "subjects" and not has_subjects:
            return False, "Needs mapped subject columns."
        if key == "session_comparison" and self.app.compare_df is None:
            return False, "Load a previous session file first."
        return True, ""

    def refresh(self):
        """Enable/disable report buttons based on the confirmed column mapping.
        Reports are generated ONLY from user-confirmed mappings; missing
        mappings disable the button with a clear reason instead of crashing."""
        if self.app.academic_dataset is not None:
            if not hasattr(self, 'course_panel'):
                from course_ui import CourseReportPanel
                self.legacy_pack = [(w, w.pack_info()) for w in self.winfo_children() if w.winfo_manager() == 'pack']
                self.course_panel = CourseReportPanel(self, self.app)
            for w, info in self.legacy_pack:
                w.pack_forget()
            self.course_panel.pack(fill="both", expand=True)
            self.course_panel.refresh()
            return
        if not self.legacy_built:
            self._build_legacy()
            self.legacy_built=True
            self.legacy_pack=[(w,w.pack_info()) for w in self.winfo_children() if w.winfo_manager()=='pack' and w is not getattr(self,'course_panel',None)]
        if hasattr(self, 'course_panel'):
            self.course_panel.pack_forget()
            for w, info in self.legacy_pack:
                w.pack(**info)
        missing = []
        for key, req, fmt_btns in self.cards:
            ok, reason = self._req_status(key, req)
            for b in fmt_btns.values():
                b.configure(state="normal" if ok else "disabled")
            if not ok and reason not in missing:
                missing.append(reason)
        self.hint_lbl.configure(
            text=("Unavailable: " + "; ".join(missing)) if missing else "")

    def _report_path_for(self, key, fmt):
        """Build the dated LGU_* filename used by single and batch exports."""
        label = reports.REPORTS[key][0]
        fmt_label, ext = reports.FORMATS[fmt]
        safe = "".join(c if c.isalnum() else "_" for c in label)[:40]
        dept = self.app.selected_department
        if dept and dept != "All Departments":
            safe = f"{safe}_{''.join(c if c.isalnum() else '_' for c in dept)[:20]}"
        return f"LGU_{safe}{ext}"

    def _df_for_report(self, key):
        """Working dataframe for a report, with comparison data attached
        when the Session Comparison report needs it."""
        fdf = self.app.get_filtered_df()
        if key == "session_comparison" and self.app.compare_df is not None:
            fdf = fdf.copy()
            fdf.attrs["compare_df"] = self.app.compare_df
            fdf.attrs["compare_mapping"] = self.app.compare_mapping
            fdf.attrs["compare_label"] = self.app.compare_label or "Previous session"
            src = self.app.source_path or "Current session"
            fdf.attrs["current_label"] = os.path.basename(src)
        return fdf

    def run_report(self, key, fmt):
        if self.app.academic_dataset is not None:
            self.refresh()
            from course_reports import ROUTES
            if key in ROUTES:self.course_panel.export(key,fmt)
            else:messagebox.showwarning("Report unavailable", self._req_status(key,None)[1])
            return
        ok, reason = self._req_status(key, reports.REPORTS[key][2])
        if not ok:
            messagebox.showwarning("Report unavailable", reason)
            return
        if self.app.df is None:
            messagebox.showwarning("No data", "Import an Excel file first.")
            return
        label = reports.REPORTS[key][0]
        fmt_label, ext = reports.FORMATS[fmt]
        p = filedialog.asksaveasfilename(
            defaultextension=ext,
            initialfile=self._report_path_for(key, fmt),
            filetypes=[(fmt_label, f"*{ext}")])
        if p:
            if self.app.source_path and os.path.normcase(os.path.abspath(p))==os.path.normcase(os.path.abspath(self.app.source_path)):
                messagebox.showwarning("Choose another filename","The original input workbook cannot be overwritten.")
                return
            frame=self._df_for_report(key).copy()
            mapping=dict(self.app.mapping)
            self.app.run_job("Generating " + label,
                             lambda:reports.generate(key,fmt,frame,mapping,p),
                             lambda _:messagebox.showinfo("Report generated", f"{label} saved to:\n{p}"))

    # ---------------- Batch Export All (v1.4) ----------------
    def batch_export(self):
        if self.app.academic_dataset is not None:
            self.refresh();self.course_panel.batch();return
        if self.app.df is None:
            messagebox.showwarning("No data", "Import a workbook first.");return
        base=filedialog.askdirectory(title="Choose report destination")
        if not base:return
        folder=os.path.join(base,"LGU_Reports_"+datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))
        jobs=[];log=[]
        for key,(label,handlers,req) in reports.REPORTS.items():
            ok,reason=self._req_status(key,req)
            if not ok:
                log.append(f"SKIPPED {key}: {reason}");continue
            frame=self._df_for_report(key).copy()
            for fmt in reports.FORMATS:
                jobs.append((key,fmt,frame,self._report_path_for(key,fmt)))
        mapping=dict(self.app.mapping)
        def work():
            os.makedirs(folder,exist_ok=False);success=0
            for key,fmt,frame,filename in jobs:
                try:
                    reports.generate(key,fmt,frame,mapping,os.path.join(folder,filename))
                    log.append("OK "+filename);success+=1
                except Exception as exc:log.append(f"FAILED {filename}: {type(exc).__name__}")
            with open(os.path.join(folder,"export-log.txt"),"w",encoding="utf-8") as f:f.write("\n".join(log))
            return f"{success}/{len(jobs)} files saved; inspect export-log.txt in {folder}"
        self.app.run_job("Batch reports",work,lambda result:messagebox.showinfo("Batch completed",result))

    # ---------------- Session Comparison file loading (v1.4) ----------------
    def load_compare_file(self):
        """Let the office pick a previous session's Excel file and map it
        with the same automatic detection the Import tab uses."""
        p = filedialog.askopenfilename(
            title="Choose the previous session's Excel file",
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm *.xls *.xlsb")])
        if not p:
            return
        try:
            sheets = importer.load_workbook(p)
        except Exception as e:
            messagebox.showerror("Cannot read file", str(e))
            return
        sheet = sheets[0]
        try:
            raw = importer.read_sheet(p, sheet, 0)
        except Exception as e:
            messagebox.showerror("Cannot read sheet", str(e))
            return
        mapping = importer.suggest_mapping(raw)
        if not mapping.get("roll") or not mapping.get("name"):
            messagebox.showerror(
                "Columns not detected",
                "Could not detect roll number and student name columns "
                f"in sheet '{sheet}'. The file may need different headers.")
            return
        self.app.compare_df = importer.apply_mapping(raw, mapping)
        self.app.compare_mapping = mapping
        self.app.compare_label = (f"{os.path.basename(p)}"
                                  + (f" [{sheet}]" if len(sheets) > 1 else ""))
        n = len(self.app.compare_df)
        self.compare_status.configure(
            text=f"Loaded: {self.app.compare_label} ({n} students).")
        self.refresh()

    def clear_compare_file(self):
        self.app.compare_df = None
        self.app.compare_mapping = None
        self.app.compare_label = None
        self.compare_status.configure(text="No previous session loaded.")
        self.refresh()


# ---------------- Settings frame ----------------
class SettingsFrame(ctk.CTkFrame):
    """User-configured app preference. The app never invents thresholds;
    the exam department sets its own rules here."""
    def __init__(self, parent, app):
        super().__init__(parent, fg_color=BG)
        self.app = app
        from course_ui import GazetteSettingsPanel
        self.course_panel=GazetteSettingsPanel(self,app)
        self.legacy_panel=ctk.CTkFrame(self,fg_color='transparent')
        ctk.CTkLabel(self, text="Settings",
                     font=("Arial", 22, "bold"),
                     text_color=TEXT_DARK).pack(anchor="w", padx=24, pady=(4, 2))
        ctk.CTkLabel(self.legacy_panel,
                     text="Control gazette details, position-holder certificates, "
                          "and the optional rules used by judgment reports. "
                          "Empty rule = no judgment applied (relative marking).",
                     font=("Arial", 12),
                     text_color=TEXT_MID).pack(anchor="w", padx=24, pady=(0, 8))

        form = ctk.CTkFrame(self.legacy_panel, fg_color=PANEL, corner_radius=12,
                            border_width=1, border_color=BORDER)
        form.pack(fill="x", padx=24, pady=8)

        self.entries = {}
        fields = [
            ("department", "Department (for official gazette)",
             "e.g. COMPUTER SCIENCE. Shown on the gazette cover page."),
            ("semester", "Semester (for official gazette)",
             "e.g. SPRING or FALL."),
            ("session", "Session year (for official gazette)",
             "e.g. 2026."),
            ("declaration_date", "Result declaration date (YYYY-MM-DD)",
             "Shown on the gazette cover page."),
            ("topper_count", "Position holders for certificates",
             "Top N students get merit certificates."),
            ("passing_marks", "Passing marks per subject (optional)",
             "Empty = no pass/fail judgment. Used by Failing Subjects Detail "
             "and Subject-wise Pass Percentage."),
            ("probation_cgpa", "CGPA probation cutoff (optional)",
             "Empty = no probation judgment. Used by Academic Probation List."),
        ]
        for key, label, hint in fields:
            block = ctk.CTkFrame(form, fg_color="transparent")
            block.pack(fill="x", padx=24, pady=12)
            ctk.CTkLabel(block, text=label, font=("Arial", 13, "bold"),
                         text_color=TEXT_DARK, anchor="w").pack(anchor="w")
            e = ctk.CTkEntry(block, width=180, height=38, corner_radius=8,
                             font=("Arial", 13), border_width=1,
                             border_color=BORDER, fg_color="white",
                             text_color=TEXT_DARK)
            e.insert(0, str(self.app.report_settings[key]))
            e.pack(anchor="w", pady=(6, 4))
            self.entries[key] = e
            ctk.CTkLabel(block, text=hint, font=("Arial", 11),
                         text_color=TEXT_MID, anchor="w").pack(anchor="w")

        ctk.CTkButton(self.legacy_panel, text="Save Settings", fg_color=LGU_GREEN,
                      hover_color=LGU_GREEN_DARK, width=200, height=44,
                      corner_radius=10, font=("Arial", 13, "bold"),
                      command=self.save).pack(anchor="w", padx=24, pady=16)
        self.msg = ctk.CTkLabel(self.legacy_panel, text="", font=("Arial", 11, "bold"))
        self.msg.pack(anchor="w", padx=24)

    def on_show(self):
        self.refresh()

    def refresh(self):
        if self.app.academic_dataset is not None:
            self.legacy_panel.pack_forget()
            self.course_panel.pack(fill='both',expand=True,padx=24,pady=12)
            self.course_panel.refresh()
            return
        self.course_panel.pack_forget()
        self.legacy_panel.pack(fill='both',expand=True)
        for key, e in self.entries.items():
            e.delete(0, "end")
            e.insert(0, str(self.app.report_settings[key]))

    def save(self):
        try:
            new = {
                "department": self.entries["department"].get().strip(),
                "semester": self.entries["semester"].get().strip(),
                "session": self.entries["session"].get().strip(),
                "declaration_date": self.entries["declaration_date"].get().strip(),
                "topper_count": int(self.entries["topper_count"].get()),
            }
        except ValueError:
            self.msg.configure(text="Position holders must be a valid number.", text_color=DANGER)
            return
        if new["topper_count"] < 1:
            self.msg.configure(text="Value must be at least 1.", text_color=DANGER)
            return
        if new["topper_count"] > 10000:
            self.msg.configure(text="Position holders looks too large (max 10000).",
                               text_color=DANGER)
            return
        # Declaration date must be a real calendar date in YYYY-MM-DD form
        # (it is printed on the official gazette cover page).
        try:
            if new["declaration_date"]:datetime.date.fromisoformat(new["declaration_date"])
        except ValueError:
            self.msg.configure(
                text="Declaration date must be a real date in YYYY-MM-DD form.",
                text_color=DANGER)
            return
        # Optional operator rules: empty means "no rule" (never invented).
        for key, label in (("passing_marks", "Passing marks"),
                           ("probation_cgpa", "CGPA probation cutoff")):
            raw = self.entries[key].get().strip()
            if raw == "":
                new[key] = ""
            else:
                try:
                    v = float(raw)
                except ValueError:
                    self.msg.configure(
                        text=f"{label} must be a number or left empty.",
                        text_color=DANGER)
                    return
                if v < 0:
                    self.msg.configure(text=f"{label} cannot be negative.",
                                       text_color=DANGER)
                    return
                new[key] = raw
        self.app.report_settings.update(new)
        # Re-attach to current data so reports use the new setting
        if self.app.df is not None:
            self.app.df.attrs["settings"] = dict(self.app.report_settings)
        self.msg.configure(text="\u2713 Settings saved.", text_color=SUCCESS)


def main():
    if '--self-check' in __import__('sys').argv:
        import argparse
        from pathlib import Path
        parser=argparse.ArgumentParser()
        parser.add_argument('--self-check',action='store_true')
        parser.add_argument('--self-check-output',default=str(Path.home()/'.lgu-exam-manager'/'self-check.json'))
        args=parser.parse_args()
        from self_check import run_self_check
        return run_self_check(App,args.self_check_output)
    app = App()
    app.mainloop()


if __name__ == "__main__":
    raise SystemExit(main())
