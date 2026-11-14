# Gazette v4.3 large fictional dataset

149 regression tests and 19 packaged self-checks passed. Large fixture: 2,400 students, 139,200 attempts, 13 programs, semesters 1–14. All 76,800 GPA/CGPA scopes and 38,400 exported semester rows matched an independent exact-arithmetic oracle. Save/clear/reopen, 1,193-page PDF, Word, audit Excel, and preview passed. See evidence/big-data-TEST_RESULTS.md for scope, timing and limitations.

# Gazette v4.2 workflow and rendering checks

147 automated tests passed. The independent exact-arithmetic oracle matched 15,830 numeric GPA/CGPA scopes across 42,331 attempts with zero mismatches; 1,614 unresolved scopes remained reviews. See SYSTEM_WORKFLOW_REPAIR.md for the new explicit workspace actions, unified import review, actual-path source protection and PDF frontmatter/rotated-label repairs. Current regression and packaged diagnostic results are included in evidence.

# Gazette v4.1: empty startup

144 automated tests passed. A normal launch leaves student records, source filename and report data empty even when a saved workspace exists. Regression checks preserve the saved file while preventing automatic restore. The Windows self-check includes the same startup check against a real saved diagnostic workspace.

# Gazette v4.0 workbook-independent import verification

143 automated tests passed. Fictional fixtures cover XLSX, XLSM and XLS imports, unfamiliar headings, two-row headings, deep title sections, wide course groups, original row provenance, structural mapping reuse without value reuse, source-point precedence, older-format session recovery, and blank legacy cover metadata. The desktop self-check has 16 checks, including the actual mapping controls and export readback. See GENERAL_IMPORT_REPORT.md and evidence for details.

# Gazette v3.9 workspace and supplied-output regression



132 automated tests passed. Real Browse-command replacement of REHMAN.xlsx (2,721 attempts) with FALL 2020.xlsx (42,331 attempts) passed, including automatic save, previous-workspace backup, reset filters/settings, visible full/current counts, mismatched-template blocking and PDF/Word/Excel output. Six current-term CGPA reviews were traced to their historical source rows. See WORKSPACE_REPAIR_REPORT.md.



# Gazette v3.8 expanded verification



127 automated tests passed. The desktop matrix passed all 85 visible export commands across 27 programs and 13 department layouts, with 81 fee defaulters checked at the end of the combined document. An independent Fraction-based oracle matched 15,830 numeric GPA/CGPA scopes across the complete 42,331-attempt workbook and correctly retained 1,614 review scopes. Six historical program routing gaps, Urdu department headings, and missing reference university titles were fixed. See CROSS_PROGRAM_TEST_REPORT.md for coverage, data limitations, and source references.



# Gazette v3.7 verification



The real FALL 2020 session was checked using the exact reported selection: SU26, All programs, Automatic. All 13 current-term attempts across seven students and four department sections exported through the desktop commands to PDF, Word and Excel; preview passed. PDF cumulative values were checked against the academic calculation index, and fee appendices occur after all academic sections. Unknown academic statuses remain REVIEW rather than invented values. Routing aliases and preferred templates are editable in gazette_assets/program_routing.json.



The rebuilt Windows executable passed its own opt-in self-check: background import, all navigation sections, exact repeat/GPA/CGPA assertions, PDF/Word/Excel output readback, preview generation, all 33 report dispatcher outputs, zero-failure batch export, and dataset-only session reload/resave. This run included F=0.00 in the bundled profile and verified a zero-GPA case.





After F=0 was confirmed, independent exact checks matched 41,836 derived attempt-point values and 1,000 latest cumulative student scopes. Original course records stayed unchanged. Of 42,331 attempts, 41,836 have derived points, two U grades lack a rate, 481 are non-final/incomplete, and 12 have nonpositive credit hours. The remaining 425 latest cumulative scopes require further source/status rules; these are not silently guessed.





76 regression tests passed after the fixes, including F=0.00 and managed-profile migration that preserves custom scales. The standalone-launcher negative control also passed: a launcher copied without its app files exits with extraction instructions and does not create an environment. Seven targeted failure controls passed, including repeated dataset-only saves, corrupted archives, serialization failure, and blank-message/callback exceptions; Word withholding has an additional independent check.



A missing PDF-reader association is also tested: the generated PDF remains available and the app displays its location.



New negative controls cover blank-message exceptions, failed UI callbacks, repeated dataset-only session saves, failed serialization preserving the old archive, malformed/incomplete sessions remaining untouched, and Word defaulter withholding with a final appendix.



# Gazette v3.5 supplied-scale verification



Desktop acceptance passed automatic profile loading, custom-scale edits, restoring the supplied scale, PDF/Word/Excel exports, and source-cell preservation. A fictional no-GP workbook rendered GPA 3.30 and CGPA 3.17 in the gazette; the body page was visually inspected and had no out-of-page text.



64 regression tests passed, covering every supplied grade rate, exact weighted points, repeats, two-place truncation, source-point priority, session migration, custom-scale preservation, and point-derivation audit cells.



Independent exact-arithmetic comparison matched 38,091 derived attempt-point values and 457 latest student cumulative scopes in the large saved dataset. 968 latest scopes still require missing grade/status/credit rules. Original course attempts remained unchanged. F is not mapped, per explicit user instruction; U and incomplete/non-final cases remain unresolved. Seven retained attempts have nonpositive credits.



The coverage of 42,331 attempts reconciles to 38,091 derived points, 3,752 missing grade rates (3,750 F and two U), 481 non-final/incomplete attempts, and seven nonpositive-credit attempts. Source errors and identity conflicts remain separately enforced during report validation.



# Gazette v3.4 verification



52 regression tests passed. The desktop workflow passed background import, six navigation sections, filters/search, PDF/Word/Excel exports, session save/reload and a 30-file batch with zero failed outputs.



The expanded suite contains 52 tests, including failed-export preservation, Windows case-insensitive source protection, mismatched extensions, missing later-term CGPA history, current-term export scope, invalid term choices, busy cancellation state, grade-scale limits, immutable export inputs, cover paste persistence and continued review access. Legacy dispatcher: 63 successful outputs, zero failures; locked official-layout route excluded.



# Gazette v3.3 verification



The actual large-workbook selected-report desktop path passed: SU26, BSCS, five course attempts; direct program selection; PDF/Word/Excel export; source-row review dialog; preview. All four current-term program selections are export-ready with their explicitly selected reference layouts. Existing numeric GPA inputs remain absent, so those exported cells are marked REVIEW. This is a successful source-record export, not evidence of calculated numeric GPA for that workbook.



The grade-scale Settings UI was also exercised with a clearly fictional no-points workbook: applying explicit test rates produced CGPA 3.17 and left original GP cells blank.



43 regression tests passed, including unrelated errors, blocking missing historical rows, explicit scale validation, exact totals, source preservation and scale/session/Excel-audit round trips. The 29 errors in the large workbook do not affect its 13 SU26 attempts or their imported student histories. This corrects the v3.2 global blocking behavior; prior evidence that all result exports were blocked is superseded. Information notices are shown separately from errors and warnings. No grading scale was inferred from marks or inserted as a default.



# Gazette v3.2 end-to-end verification



Large-workbook desktop acceptance passed all navigation, paging/search, session persistence, source-error blocking, validation PDF/Word/Excel exports and PDF preview. Under concurrent Windows-build load, import took 29.817 seconds and maximum observed UI heartbeat gap was 5.858 seconds.



40 automated regression tests passed after the batch-filter fix. The test suite covers imports, schema mapping, repeat handling, exact arithmetic, two-place truncation, current-term scope with historical CGPA, semester 1ï¿½14, fees, sessions and desktop callbacks.



The legacy dispatcher generated 63 reports successfully (the locked official-layout route was excluded). After the fix, desktop acceptance with the smaller supplied workbook passed navigation, paging, filters, PDF/Word/Excel exports, save/reload and a 30-file batch. Independent arithmetic matched 1,189 scopes; 35 scopes correctly require source review. Current-term output was checked against selected course counts and page boundaries. All 13 reference families rendered with zero out-of-page words; the generated contact sheet and three-defaulter example were visually inspected.



The batch fix keeps the academic term filter for ordinary reports and complete selected-student history for gazette/CGPA reports. This does not resolve missing source values or confer publication approval. Word files were reopened and structurally checked; Microsoft Word pagination was not visually checked.



## v3.1 large workbook validation



39 regression tests passed. Desktop acceptance passed background import, navigation, 150-row paging, program filtering, case-insensitive search and complete session save/reload. Desktop import: 16.079 seconds; save: 8.239 seconds; restore: 1.569 seconds.



FALL 2020.xlsx imported in 14.385 seconds, compared with 38.584 seconds before optimization on this machine. Peak process working set was 147.43 MB. All 42,430 original source rows and all 41,894 previously recognized attempts remained unchanged; 437 additional status records were retained. Twenty rows require reconciliation, 77 are placeholders, and two are repeated headings. There are nine student-name conflicts for review. Numeric grade points and fee status are absent in this source; neither is inferred. Registration capitalization resolves 1,483 original spellings to 1,425 identities. This validates this workbook size; it is not an unlimited-scale guarantee.



# Gazette v3 verification



- 36 automated tests passed (latest complete run: 116.50 seconds). Coverage includes imports/mapping/provenance, separate student identity, exact arithmetic/truncation, repeat selection, current-term scope, all-semester summaries, 14-semester progression, missing explicit semester history, fee import/session persistence/conflicts, source protection, all template families and actual Tk mode/settings/report behavior.

- Reference audit: 13 supplied PDFs, 2,762 pages, 120 measured header/grid profiles; 2,734 compatible body pages, two additional fee-only pages and 26 front-matter pages; zero unclassified pages. All 120 header profiles render. Synthetic complete gazettes for all 13 families have zero out-of-page extracted words. Cover/body contact sheets were inspected; common cover artwork removed the old text-band overlap problem.

- Independent arithmetic over the supplied workbook previously matched 1,189 term/cumulative scopes; 35 scopes require source review because of 11 zero-credit attempts. Original workbook: 150 registrations, 2,721 attempts, 16 placeholders. No fee data supplied.

- Current-semester real-workbook output: latest supplied term SU22, 22 attempts, 16 students, 11 PDF pages; PDF course-code counts match exactly and no extracted word lies outside its page. The all-semester roll list has 612 student-term rows. Older/current data stay in the dataset and calculation history even though earlier course rows are not printed in the current gazette.

- GPA column test checks actual PDF cells: current GP 13.2, cumulative T.GP 22.2, current CH 4, cumulative T.CH 7, SGPA 3.30 and CGPA 3.17. Values use the same kept attempt's credits/points; no round-off is applied.

- Fictional demo: four students, 16 historical/current attempts, three current fee defaulters; one clear student in the academic body, three defaulters in the last appendix; eight semester-summary rows. Current GPA 3.50 and cumulative CGPA 3.00 are independently asserted. A provided semester-14 record without earlier source semesters gets review rather than fabricated history.

- Latest full Tk acceptance passed background import, all navigation, pagination, preview filters/search, visible PDF/Word/Excel commands, save/reload and 30 batch outputs. Import measured 6.59 seconds, filtered batch 16.20 seconds, max worker heartbeat gap 5.34 seconds under concurrent PDF/build load. These are measurements, not guarantees; background threads do not eliminate Python/GIL/system contention. The earlier quiet run was faster.

- A synthetic unseen 10,000-student workbook imports successfully. Limits remain 50 MB/workbook and 200,000 rows/sheet; a 200,000-row stress run was not performed.



## Scope and evidence



Source documents/workbooks are read only. Public release evidence is aggregate or fictional; no real student workbook, exports, autosave or reference PDFs are bundled. Source/math/GUI tests were executed with Python 3.11.9. Windows packaging code objects are checked against tested source after build, and ZIP contents are integrity-checked.



This is not a guarantee of error-free university outcomes. Source completeness, university golden cases, unresolved zero-credit/non-GPA cases, standing/promotion/merit rules and actual fee records still require appropriate evidence. Missing records are not inferred. Current results/roll lists are calculated from supplied history; explicit incomplete semester-number coverage is flagged. Word/Excel output is structurally checked; Word pagination was not visually validated in Microsoft Word.

