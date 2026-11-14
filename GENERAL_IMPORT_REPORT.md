# Gazette v4.0: workbook-independent import repair

The provided workbooks are regression fixtures only. Production import does not select a schema by filename, student identity, sample program, admission cohort, term, or row count.

## Changes

- Explicit worksheet review for unfamiliar structures; no silent fallback from an unrecognized course table to student summaries. Legacy summaries require an explicit GPA field.
- One- and two-row headings, arbitrary user-selected header rows, multiple sheets, multiple files, long course tables and wide course groups.
- Ambiguous GP fields require an explicit choice between weighted points and points per credit. Numeric ranges and marks are not used to guess their meaning.
- XLSX/XLSM use openpyxl; XLS/XLSB use the bundled calamine reader. Input macros are not run. Fee imports accept the same extensions.
- Literal NA text survives import. Original physical Excel rows remain auditable when a wide row expands into multiple attempts.
- Remembered column mappings omit literal metadata and course defaults. A workbook needing those defaults requires review again.
- New program labels can use an explicitly supplied department to choose an existing gazette geometry. Known program/department mismatches remain blocked.
- Blank initial cover metadata replaces sample department/year/date defaults. Operator error dialogs omit Python exception class names; support logs retain diagnostics.

## Verification

The automated regression suite covers original academic calculations, repeats, identity conflicts, fee appendices, selection scope, session recovery, worker failures and unfamiliar workbook structures. Additional desktop self-checks exercise the real mapping controls, cancellation, applying wide groups, original row references, non-reuse of literal metadata, unseen program routing, source rate precedence, the packaged XLS reader, PDF/Word/Excel readback, preview, navigation and batch export.

XLSX, XLSM and XLS course imports use independent fictional fixtures. XLSB reader verification uses the public pandas `test1.xlsb` fixture from its official GitHub repository; it verifies binary decoding, not university-specific academic interpretation. Missing columns or meaning still require explicit mapping. Automated checks and packaged self-check results are included in evidence.

## Limits

Supported formats and mappings cannot recover absent student identities, grades, credits or historical results. Formula cells require saved cached results from Excel when no literal value is stored. Non-final academic statuses need an explicit policy. No inferred marks-to-grade scale, fee status, standing or rank is substituted for missing rules. Scanned pages, password-encrypted workbooks and arbitrary report artwork are not Excel course tables.

Existing limits: 50 MB per workbook; 200,000 course attempts per imported workbook. Split larger files and use multi-file import. These are technical limits, not sample-data assumptions. All prior department geometry and confirmed grade rules remain externally configured.
