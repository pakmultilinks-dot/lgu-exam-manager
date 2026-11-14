## v3.1 identity and import performance

`academic_identity.py` supplies bounded normalization caches shared by counting, history, summaries, fee matching and gazette grouping. Registration keys retain punctuation and ignore capitalization; program keys normalize typography. Source fields remain unmodified. Tuple iteration avoids allocating a pandas Series per source row. Incomplete non-final status records remain in the audit and do not receive invented credits or points.

# Gazette v3 architecture

academic_engine.py parses source attempts and fee records, with registration identity, validation, provenance, optional semester number 1-14, multi-file support and retained raw source rows. session_store.py saves bounded non-executable JSON/Excel archives including fee records and settings.

calculation_policy.py implements the explicit user-confirmed formula and best-repeat rule using Decimal inputs and exact Fraction ratios. AcademicIndex caches per-student/term summaries over full imported history. Source conflicts, incomplete grade states, zero credits and explicit missing-semester history require review. Display truncation uses exact integer arithmetic.

gazette_options.py reads template choices from the catalog, resolves programs conservatively and selects the current print term independently of Academic Records' term preview filter. Cover/semester settings are snapshotted before background export.

gazette_layout.py renders measured reference grids and current-term courses, with current GP/CH and cumulative T.GP/T.CH/CGPA. Editable cover regions and source templates live in gazette_assets/reference_layouts.json. semester_reports.py creates all-semester roll-wise summaries and a companion CGPA table for reference variants lacking CGPA. fee_register.py validates selected fee statuses, withholds verified defaulters from result body pages and paginates the final fee appendix.

course_ui.py provides current-term and all-semester workflows, shared settings, fee import, preview and export. academic_frame.py retains independent source-record filtering; gazettes/CGPA lists use its student/program/search selection while Settings controls the gazette term. app.py orchestrates Tk/main-thread updates and queued background work. Legacy reports remain separate from course-attempt calculations.

No sample student count, registration, marks, semester date, program or defaulter list drives production results. Reference geometry and confirmed arithmetic rules are configuration/policy; they are not inferred student data. The parser does not claim to infer arbitrary documents or unseen enrollment history.
