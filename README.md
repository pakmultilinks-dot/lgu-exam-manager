# LGU Exam Manager — Gazette v4.3

Extract the complete ZIP into a new folder, close older versions, and run `START_LGU_Exam_Manager.bat`. Keep the executable and its `_internal` folder together.

Import Excel workbooks (`.xlsx`, `.xlsm`, `.xls`, `.xlsb`). Recognized course tables load automatically. Unfamiliar columns, ambiguous points fields, or extra result tables open **Review workbook import**: choose sheets, heading row, one or two heading rows, and map the actual columns. **Courses across columns** expands student rows into course attempts. Supply sheet-level details only when they apply to that sheet. Cancel preserves the active workspace.

The import review screen is the single place to choose sheets, heading rows and column meanings. Remembered mappings contain column structure only. Entered programs, terms, course codes, titles and credits are never reused from a previous workbook. Dummy files are test fixtures, never startup data or a schema requirement. Macros are not executed.

Choose the current exam term in Settings or Reports. Academic pages show that term; CGPA includes all imported earlier history through that term. Source department metadata can route an unfamiliar program to an existing department layout. Without enough department evidence, choose a supported layout explicitly. Cover text is editable. Fee-defaulter appendices follow academic sections.

The configured LGU grade scale includes F=0.00. Weighted source points take priority, followed by explicitly mapped source points per credit, then the editable grade scale. GPA/CGPA use exact points / credits and the highest-rate repeated course once. Display uses two decimals without rounding. Unknown statuses, invalid credits and conflicting identities remain visible for review; missing academic rules are not guessed.

The app starts empty on every normal launch, even if a saved workspace exists. Import Excel to load records. Save preserves a workspace backup without automatically loading it on startup. Use **Open saved workspace** to load it explicitly, or **Clear current data** to return to an empty workspace without deleting saved records. No result workbook is bundled as default data.

Run `CHECK_LGU_Exam_Manager.bat` for opt-in diagnostics using temporary fictional records. See `GENERAL_IMPORT_REPORT.md` for current verification and limits. Older test reports document earlier releases.
