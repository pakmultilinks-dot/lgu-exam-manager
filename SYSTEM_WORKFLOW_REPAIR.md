# Gazette v4.2 workflow repair

The application starts empty. Import Excel to create active records; Save writes a local workspace; Open saved workspace restores it only after a user action. Clear current data empties active records and the header, removes stale comparison data, and preserves the saved file. Open and Clear are guarded while a worker is running.

Import review is the single visible path for sheet, heading-row and column mapping. The obsolete sheet/header dropdowns were removed because they did not reliably change course imports. Existing automatic recognition and structural mapping profiles remain available.

Exports protect actual original source paths, including resolved file aliases. An output with the same basename in a different folder is allowed for newly imported workbooks. Saved snapshots retain these paths. Older snapshots without source paths keep their conservative filename guard.

The reference cover artwork is no longer covered by an added white DRAFT footer, and the contents page no longer receives that label. The rotated NOTE label now uses direct text positioning: its former negative cell coordinate was treated as a position relative to the right margin and put the label outside the page. PDF regression checks require its presence and keep all extracted words within page bounds.

Word register labels and Excel audit metadata no longer claim that a grade scale/template is missing when it has been supplied. The source launcher now checks the binary Excel reader dependency as well as the original packages.

Verification uses fictional workbooks and explicit desktop control invocations. Saved files survive clear and reopen tests. The regression suite covers calculations, repeats, import mapping, source audits, recovery, worker errors and exports. The packaged diagnostics also exercise the new Open/Clear controls and confirm empty startup with a saved workspace present. A catalog recheck accounts for all 13 supplied department references, 2,762 pages. Poppler-rendered fictional cover/body pages were visually inspected after the footer and rotated-label repairs.

No marks-to-grade conversion, fee status, standing or promotion policy is inferred from missing data. Review requirements remain visible when source records or academic rules are incomplete.
