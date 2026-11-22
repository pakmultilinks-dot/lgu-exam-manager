# v4.6.9 (22 Nov 2026)

- Launcher fix: the app no longer installs all packages again in every new version folder. The Python environment now lives in one shared place outside the version folder (LGU-Exam-Manager in the user profile), so packages are installed once and every later version reuses them and opens without installing again. The first run of v4.6.9 installs once into the shared place; after that, no more installs. If the shared environment is ever broken, the launcher removes it and creates it again by itself.

# v4.6.8 (22 Nov 2026)

- Excel Gazette: the Adj. GP and Adj. CH cells on repeated courses are now center-aligned, matching the other numeric columns (SEM, COURSE CODE, GRADE). Values were already correct since v4.6.7; only the alignment was left-aligned.
- No .exe build: delivery still uses START_LGU_Exam_Manager.bat, not a standalone .exe. The LGU_Exam_Manager.spec file in the repo is not used for delivery, so there is no .exe to rebuild.

# v4.6.7

Three fixes, each verified on the real F20 file by rendering the actual
PDF and Excel output.

(1) Excel Adj. GP and Adj. CH now match the PDF on repeated courses. The
PDF already showed the replaced prior attempt (MUHAMMAD AYHAN FAROOQ
F22 BTCH311 Adj. 0.00 and 3), but the Excel export left both Adj.
columns blank on repeats. The Excel gazette now uses the same Adj. rule
as the PDF: blank on a first-time course, the replaced prior attempt's
GP and CH on a repeated course. First-time courses stay blank in Excel.

(2) PDF cover: the badge area no longer shows a bright-white rectangular
patch. The badge box is rebuilt from the cover's own dotted background
texture (a smooth inpaint of the surrounding paper, with the badge's
drop shadow removed and the green border frame restored), so the old
badge area blends with the rest of the cover. The semester is back on
the cover, as plain printed dark-green text in the top-right area where
the badge was, with no background, no gold and no ribbon: FALL
SEMESTER 2020, sized to fit on one line.

(3) PDF T.CH header: the label is shrunk by an extra 0.5 pt and centred
inside its narrow column, so it now clears the column's right border
instead of touching it.

# v4.6.6

Full audit fix, every item verified on the real F20 file by rendering the
actual PDF and Excel output.

PDF gazette. (1) Adj. GP and CH are now blank on first-time courses and
show only the replaced prior attempt on a repeated course. Before, every
course printed its own GP and CH into Adj., and a repeat showed the new
GP instead of the replaced one. AYESHA FIAZ F20 Adj. is now blank, and
MUHAMMAD AYHAN FAROOQ F22 BTCH311 shows Adj. 0.00 and 3, the replaced
F20 attempt. (2) Header labels no longer touch: REGISTRATION, Adj.,
T.CH and SGPA are shrunk until each has clear space. (3) The rotated
GRADE header now sits fully inside the header band instead of poking
above the top border. (4) Course codes no longer wrap mid-code:
BCHM311, BCHM312 and CHEM103 each fit on one line. (5) The NOTE area
now draws the 3 horizontal note boxes from the official gazettes.
(6) The closing FEE DEFAULTERS page now uses the same 17-column header
(T.CH and CGPA) as the body instead of switching back to 15 columns.
(7) The gold FALL SEMESTER badge is gone from the cover entirely: the
ribbon is painted out of the cover artwork and the border frame is
continuous again, and no gold box or period text is drawn. The cover
keeps the logo, LAHORE GARRISON UNIVERSITY, RESULT GAZETTE, department
band, campus photo, Prepared by, and footer. The semester stays in the
gazette body and contents.

Excel gazette. (8) Row 1 now has the lavender fill from the official
sheet. (9) Header rows 5-6 are vertically merged for SR through GRADE
and SGPA through COMMENTS. (10) The summary columns (GP, T.GP, CH,
T.CH, SGPA, Adj. GPA, COMMENTS) are merged vertically across each
10-row student block, centered, with Adj. GPA and COMMENTS bold.
(11) The NOTE area is 3 two-row boxes. (12) SEM, COURSE CODE, GRADE
and the totals are centered, and SR is bold. (13) Page breaks fall
between student blocks, so no block is split across printed pages.
(14) A Spring block now shows ONE totals row for the combined
Spring plus Summer window, not one per sub-term. (15) Gazette Contents
now counts only the students in the active publication term (F24
Maths now reports 2, not 15). (16) Course Results now sets the width
for all 15 columns.

Other reports. (17) Raw float noise (11.100000000000001 and friends)
no longer leaks into any report: register, term, cohort, degree and
repeats are clean in Excel, Word and PDF, using the existing clean
number logic. (18) The Roll-number-wise GPA and CGPA by semester PDF
headers no longer break mid-word. (19) The pack export no longer
crashes when every program is skipped; it reports the skip cleanly.

Numbers are unchanged. Full suite green and real-app verification
passed on F20.xlsx.

# v4.6.5

Look fixes, both verified against the university's own files. (1) The PDF
cover badge no longer looks like a sticker pasted on: the gold rectangle
drawn over the badge used (194,183,122), a different gold from the ribbon
artwork, so a lighter box showed around the FALL SEMESTER 2020 text. The
overlay now uses the artwork's exact ribbon gold (200,189,135), sampled
from the cover image, so the badge reads as one uniform ribbon like the
official cover. The department band overlay is likewise corrected to the
artwork's exact green (0,58,34). (2) The Excel title block now matches the
university's own sheet: the title was Cambria 18 and the department line
Cambria 13, both too small next to the official. The title is now Times New
Roman 24 bold navy, RESULT STATEMENT Times New Roman 16, the department
line Times New Roman 18 bold italic navy, and the session line Times New
Roman 12 italic, with row heights set so nothing is cut off. The numbers
are unchanged: all 12 Biology F20 students still match the source data
exactly.

# v4.6.4

Logic and error fixes found by a full audit of the F20 file, all verified on
the real file in the app. (1) Grade points that arrive as binary floats
(11.100000000000001 for 11.10, 9.899999999999999 for 9.90) no longer flip
half-up rounding: 32 SGPA/CGPA values across the PDF, Excel and Word printed
0.01 too low, for example BS BCH/003 S22 SGPA 2.72 instead of 2.73 and BS
BT/026 F20 SGPA and CGPA 2.32 instead of 2.33. The intended 2-decimal value
is now recovered from the float noise before the ratio is computed; the
source workbook is never changed. (2) The Excel T.GP and T.CH columns now
show the cumulative totals, not the semester totals. Before, a later term
showed the semester totals twice: SARA FAIZA's S21 section read T.GP 50.10
and T.CH 17 when the cumulative totals are 97.90 and 35. The PDF already
showed cumulative; the Excel now matches it. (3) One in-progress (IP)
course no longer poisons a student's whole SGPA and every later CGPA. SARA
FAIZA's F25 and SAWERA's F24/F25 used to show REVIEW with no GPA at all.
Non-GPA grades (IP, W, I, P, PASS, R, blank) are now skipped, so the GPA is
computed from the graded courses; a student with only non-GPA attempts
still gets a review. (4) The Excel Gazette sheet now shows only the
publication term window, like the PDF. Before, an S25 Excel printed the
student's full history (572 rows for S25 Maths) and a Fall gazette printed
the Summer sections that belong to the Spring gazette. Now an S25 Excel
shows S25 and SU25 only, and an F21 gazette excludes SU21. (5) The CGPA
probation cutoff can now be changed in Settings with a dataset loaded: the
field was only on the legacy Settings panel, so the v4.6.3 instruction to
change it in Settings could not be followed in the normal workflow. The
legacy hint is fixed (empty now means 2.00 for gazettes) and the Probation
List no longer blocks when the cutoff is empty. (6) The PDF Adj. GP now
matches the exact source row for a repeated course, instead of returning
the first (course, term) match for every row. (7) The default term now
counts a Spring cycle (S plus its Summer) together, so a big Spring cycle
cannot lose to a smaller Fall term. The Word register also prints the
standing per term now (G.A.ST / PROBATION), which v4.6.3 added to the PDF
and Excel only. Full suite: 173 passed, 1 skipped. Real-app verification
on F20.xlsx: all checks passed.

# v4.6.3

COMMENTS now prints the academic standing, in the Excel Gazette sheet and
in the PDF, like the official gazettes. Before, COMMENTS was empty, so the
sheet did not look like the university's own sheet. Now a student with CGPA
at or above the probation cutoff gets G.A.ST, and a student below it gets
PROBATION. The cutoff is the operator's probation_cgpa setting in Settings;
when it is empty the line is 2.00 CGPA, the usual university probation line,
and the READ FIRST sheet states the cutoff that was used, so a different
line is a settings change, not a guess. On the F20 file: BS BCH 11 G.A.ST
and 1 PROBATION, BS Maths 12 G.A.ST and 3 PROBATION (the all-F student is
PROBATION at 0.00). The window title also no longer says the old version:
it was hardcoded to v4.5.2 even after v4.6.2 shipped. It now comes from one
APP_VERSION constant, so it can never go stale again. Regression tests:
standing in the Excel (G.A.ST and PROBATION), standing in the PDF, and the
title version.

# v4.6.2

Excel Gazette sheet now matches the official LGU layout, the same table as
the PDF. Before, the Excel sheet had a different table: RESULT GAZETTE, a
PUBLICATION CYCLE line, GP RATE / COURSE GP / TOTAL GP columns, the student
identity repeated on every row, and raw float text like 11.100000000000001.
Now it has RESULT STATEMENT, a SESSION line, the 17 official columns (GP and
Chr each split in GP, Adj., T.GP and CH, Adj., T.CH, plus SGPA, Adj. GPA and
COMMENTS), one block per student with the identity once, totals on the last
course row, clean 2-decimal numbers (64.10, 3.77), blocks padded to 10 rows
with '-' rows, and the NOTE ( If any ) label. The title appears once at the
top only and is never repeated on printed pages. The audit sheets are
unchanged. Regression test: the Excel Gazette sheet matches the official
layout and the example student block.

# v4.6.1

Navigation fix: the Dashboard metric cards used to carry a hidden click
binding that jumped to Academic Records, so clicking anywhere on the
Dashboard looked like a navigation error. The cards are display only now.
The labelled "Filter / search / inspect" button stays as the way to open
Academic Records. Regression test: a click on a metric card leaves the
Dashboard shown.

# v4.6.0

New "All terms gazette pack": a file that spans several terms (for example Fall 2020 to Fall 2025) now gets one department gazette per term, in a subfolder named by the term (F20, S21, F21, ...). Each term's gazette holds only that term's programs and students, with CGPA from the full history through that term, and nothing else from the file. A pack-summary.txt at the root lists what was made for each term. The single-term pack is unchanged.

Default term fix: when no publication term is chosen, the app now uses the term with the most results instead of simply the latest term. A file ending with a nearly empty latest term (a new term with only a handful of rows) no longer produces a useless gazette. An explicitly chosen term still wins.

# v4.5.2

Validation scoping now uses mapped program fields for source errors even when required identity fields are missing, while retaining registration-level scoping for filtered student selections. Unassociated source errors fail closed. This avoids coupling export decisions to a specific workbook, filename, program list, or Excel row.

# v4.5.1

Reports now replaces the long multi-paragraph blocker wall with a concise list of blocked program names and a dedicated review action. Review opens scoped to blocked programs, shows each program's reason, and starts on findings affecting those programs so source workbook, sheet, row, and source values are easier to inspect.

# v4.5

Reports and Settings now offer Fall/Spring publication cycles only; same-year Summer results join Spring and Summer-only programs remain discoverable within that cycle. The Reports program list is export-ready: programs with validation or routing blockers are excluded, summarized with reasons, and linked to source review. Review data now gives workbook/sheet/row locations and can open the imported workbook for operator-led correction and reimport. The original workbook is never modified by the app.

# v4.4

Reports now offers gazette programs only when the selected exam term has source rows for them. The selected program view distinguishes current-term course attempts from the full-history rows included in publication, and explains when historical source findings pause an export. Gazette layout routing stays with the selected program in Reports instead of a separate generic automatic-layout setting. Fee status is resolved for the publication cycle, so an earlier semester's status does not conflict with a later one; contradictory records within the applicable term still block publication.

# v4.3

Large fictional workbook verification: 2,400 students and 139,200 course attempts. Excel summary counts retain numeric types. Session saves stream dataclass records without a full deep copy; save and restore enforce the same 512 MiB dataset limit, with atomic preservation of previous saves. Dummy data remains a separate optional fixture; normal startup remains empty.

# v4.2

Explicit Open saved workspace and Clear current data controls; unified import-review path; actual-path source overwrite protection; preserved reference cover artwork without added DRAFT frontmatter labels; source launcher checks the older Excel reader dependency. Startup remains empty.

# v4.1

Normal launches start empty. Saved workspaces no longer load automatically; import a workbook to populate records. Updated Save confirmation and startup regression coverage.

# v4.0

Workbook review for unfamiliar and wide structures; structural mapping profiles without literal data reuse; source grade-point rates; XLS/XLSB readers; source-department routing; clear operator errors and blank cover defaults. See GENERAL_IMPORT_REPORT.md.

## 3.9



Successful imports now persist the active workspace automatically with a previous-workspace backup. Fresh workbook imports reset the prior file's term, department cover, manual layout and hidden record filters. Startup identifies the restored workbook explicitly. The header always names the active file; Reports distinguishes imported totals from current-term printing and provides a direct term selector and full-grid action.



Explicit layouts cannot silently apply Computer Science headings to known programs from other departments. GPA/CGPA review now identifies the exact historical source row and unresolved grade in Review data and shows the count before export. Unresolved numeric PDF cells use a dash, with the review marker in Comments, to prevent split REVIEW text in narrow number columns. No academic inclusion policy was invented.



Regression checks cover replacing REHMAN with the complete FALL 2020 workbook, automatic persistence, backup recovery, reset filters, restart, scope display, department routing, three export formats, and original-file preservation.



## 3.8



The expanded degree/grade/credit-hour and desktop matrix found six missing historical program routes. Added data-driven department mappings for BS Biochemistry, BS Zoology, CERUF, MPhil Urdu, MSBA, and the MBA (2 Years) spelling alias, with official LGU source references in the editable routing file. Visual verification also caught Urdu headings inherited from shared International Relations grid geometry and 21 profiles whose university title was omitted during header extraction. Program headings now follow editable program metadata; the missing titles are restored from their original reference pages. Mixed Arts fee pages use an encompassing department heading. Original degree labels and academic values remain unchanged. Added independent arithmetic-oracle and semester-boundary regression tests, plus a 27-program desktop acceptance script.



## 3.7



Automatic gazette export now groups mixed selections into department sections. Editable program aliases recognize BBA (Hons); BS IT prefers its Information Technology reference. All fee appendices follow all academic sections. Added combined-program export regression coverage.



## Gazette v3.6 � reproducible failure repairs



- Updated the supplied scale to F=0.00 after the user's new confirmation. Automatically managed saved profiles are refreshed; legacy custom scales are marked as custom and preserved.



- A missing Windows PDF-reader association no longer turns a generated preview into a generic application failure; the app shows the saved PDF location.



- Fixed dataset-only/multi-workbook session resaving when no data.xlsx member exists. Session writes use unique atomic staging files.

- Damaged session archives produce an actionable error and remain unchanged.

- Exceptions with blank messages cannot enter the success path; callback failures are logged as failures. Logs now include time, version, operation and traceback locations without source-cell contents.

- Word now withholds fee-defaulter academic records and lists those students at the end, matching PDF behavior.

- Added an opt-in packaged executable self-check using fictional temporary records, all report dispatchers, batch exports and session resaving. It does not load the operator's autosave.



## Gazette v3.5 � supplied LGU grading profile



- Added the user-supplied D�A grade-point table as editable JSON data and an automatic default for new imports and older unconfigured sessions. F remains unresolved at the user's explicit request.

- Derive missing weighted points from source credit hours and the applied rate; never replace supplied numeric points or original cells. Preserve custom scales and intentional clearing in saved sessions.

- Settings shows coverage and missing rates, with a restore-profile action. Reports name missing rates affecting selected history.

- Excel includes per-attempt Point Derivation Audit and the applied scale/reference; CGPA lists identify the scale used.



## Gazette v3.4 � failure handling and consistent report scopes



- Failed exports preserve the previous destination file; source-name protection ignores capitalization on Windows. Export extensions must match their format.

- CGPA lists validate full selected-student history, including quarantined attempts in later terms. Current-term gazettes are scoped before validation in batch/API exports.

- Explicit unavailable exam terms produce an actionable Settings error; automatic selection ignores unrecognized terms without silently changing an explicit choice. Validation/mapping reports remain accessible.

- Busy actions cannot restart a batch, clear its cancellation state or modify the applied grade scale. Workers use a snapshot of scale, fee records and settings.

- Cover fields track pasted/programmatically inserted text; dashboard distinguishes source errors from information/warnings. Batch summaries report failures and skipped routes.

- Grade-scale text rejects excessive numeric precision/exponents before calculation; session metadata is size-checked before reading.



## Gazette v3.3 � selected report validation and usable review



- Errors block only affected student/program/term scopes; cumulative reports also check earlier history.

- Added source-row review with severity filters, selection relevance and full review-log export.

- Reports include a program selector; counters distinguish errors, warnings and information. Dashboard normalizes program identity.

- Settings accepts an explicit grade-point rate scale and its reference for missing GP inputs. Source values take priority, sessions retain the scale, and Excel records it separately. No default scale is assumed.



## Gazette v3.2 � end-to-end verification



- Batch exports now honor the selected term for ordinary reports; gazette and CGPA routes retain earlier history.

- PDF preview captures the selected report before starting its worker.

- Added a regression check for batch filter behavior.



## Gazette v3.1 — large workbook repair



- Import iteration uses tuples and bounded normalization caches.

- Registration capitalization and program typography no longer split student history. Original source spellings remain in records and audit.

- IP, W and PASS records with absent credit hours remain visible without invented credits.

- Validated FALL 2020.xlsx: 42,430 source rows, 42,331 retained attempts, 1,425 student identities, 18 terms and 26 normalized programs.



## Gazette v3 - current term and all-department support



- Audited 13 department PDFs / 2,762 pages and catalogued 120 reference profiles, including two separate fee-only pages.

- Current-term print selection in Settings; full earlier history remains available to cumulative calculations and all-semester roll lists.

- GPA / CGPA roll lists include name, registration, program, source semester number (1-14), exam term, semester/cumulative credits and subjects.

- Cumulative T.GP/T.CH are separate from current GP/CH; preserve Criminology's Adj. GPA and use a companion CGPA summary.

- Editable cover fields over one reusable reference artwork; fixes old text-band overlap across different source covers.

- Import and retain fee records from result rows or separate lists; validate identity/conflicts; defaulters appear only in final academic-publication appendix.

- Explicit missing semester-number history blocks misleading cumulative values. Session snapshots include fees and source semester numbers.

- 36 tests pass; fictional three-defaulter example and complete 14-semester case verified.



## Gazette-focused update



- GPA shows exactly two decimals using exact integer truncation; no rounding, exact calculations preserved.

- Removed the explanatory body-page footer from gazettes.

- Added automatic layout selection from the reference program catalog, with explicit review for ambiguous or mixed departments.

- Replaced unrelated course-mode settings with gazette layout and detected source metadata.

- Removed contradictory GPA status and red policy text; kept exact audits available in Excel.



## 2026-10-22 - Confirmed calculations and broader imports



- Connected user-confirmed points / credits and highest-rate repeat selection to reference PDF columns and Word statements.

- Added exact calculation totals and kept/excluded provenance sheets to Excel.

- Cumulative calculations use imported history through term regardless of display filters; unresolved source cases require review.

- Extended header search to 200 rows, heading aliases, explicit intake, semester normalization and repeated-heading recognition.

- Added varied-file and 10,000-student import verification.



# Changes — 2026-10-22



- Reproduced the blank dashboard on REHMAN.xlsx and confirmed the legacy/student-row gating defect.

- Connected the dashboard and a dedicated course Report Center to the typed course dataset without putting attempts into student-row statistics.

- Added student, term, cohort, degree and department registers; source-count course summaries; repeat and summer/regular audits; mapping and validation exports. Department reports require actual department data.

- Added exact source-row retention, source-file/sheet/Excel-row provenance, multi-workbook duplicate/conflict checks, editable mappings, and a manual heading/mapping preview.

- Added debounced search, 150-row pagination, honest capability explanations and route-specific legacy restrictions.

- Moved parsing, save/restore and report generation into queued worker jobs. Tk updates occur on the main thread. Added cancellation between batch files and private operation logging.

- Repaired session persistence so reviewed mappings, canonical records and quarantined source rows survive reopening; saves replace archives atomically.

- Added Unicode draft PDF fonts, pagination, repeated headers, student bookmark/contents index, configurable draft headings, and landscape Word registers.

- Fixed dashboard geometry-manager conflicts, a clipped contents heading, and eager construction of legacy report cards. Course reports use a compact report selector with shared format buttons.

- Added engine, real Tk callback, mode-switch, source-protection, session and policy-boundary regressions; ran authorized real-data acceptance and native Windows import/export/save actions.

- Added portable Windows build configuration and a launcher that selects the packaged application when present.





## Gazette format correction



- Replaced the generic warning-cover PDF with the actual supplied LGU cover artwork, session/degree contents and landscape statement layout.

- Matched the 612x792 cover/contents and 792x612 statement sizes, Cambria/Calibri typography on Windows, reference table coordinates, merged identity/note fields, GP/adjustment/CH/SGPA/comments columns and body page numbering.

- Added separate Computer Science and Management Sciences template choices in the Report Center and bundled the required cover assets.

- Removed repeated warning paragraphs from the cover. Draft/calculation status is confined to a small footer; source term coverage is displayed on the cover badge.

- Independently checked all 2,721 course rows in both generated PDFs; each template has 163 pages with no out-of-page text. Unsupported academic totals remain blank.





## Whole-gazette correction



- Audited every page in both supplied references and replaced the single guessed table with 43 measured header/grid profiles and cohort/program routing.

- Restored semester/summer block segmentation, per-row adjustment cells, CGPA/cumulative-credit variants, original header colors/origins and fee-section geometry.

- Added complete-reference-catalog and exam-window regressions. Preserved all source attempts; real-data PDFs now span 281 pages per selected department style.
