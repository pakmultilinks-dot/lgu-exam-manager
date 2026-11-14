# v3.9: active workspace and supplied PDF repair

The supplied screenshot showed a restored REHMAN workspace (150 students / 2,721 attempts). The application log also confirms a later successful import of the large workbook. Those are different workspace states; the import was not limited to the smaller sample.

The supplied LGU_Gazette.pdf used a Computer Science layout for BBA, IT and SE as well as CS. LGU_Gazette2.pdf used automatic department sections. Both print SU26, the latest term, which has only 13 attempts from seven students in the 42,331-attempt workbook. The earlier imported history feeds CGPA; it is not repeated as current-term course rows.

## Repairs

- A successful workbook import saves the active workspace automatically. The previous workspace is retained as autosave.previous.lguexam. Save failure leaves the imported data usable and reports the persistence problem explicitly.
- Fresh imports reset the previous file's term, department cover, manual template, search, degree and cohort filters. An import requested during another operation does not relabel the active file.
- Startup names the restored source on the Import screen. The header shows the active filename and complete dataset count on every page.
- Reports separates imported counts from current-term print counts, exposes the current term directly, and provides Show all imported records.
- A known program cannot be silently printed using another department's layout. Automatic still creates the appropriate department sections.
- Missing GPA/CGPA results are explained before export and in the Calculation issues review, with original source sheet and Excel row. Narrow numeric PDF cells display a dash; the review marker remains in Comments. No non-final-grade inclusion rule was invented.
- Workspace tests run in isolated files instead of overwriting the operator's workspace.

## Actual reproduction

132 automated tests passed. The real Browse callback imported REHMAN and then FALL 2020 with deliberately stale filters, term, cover and manual Computer Science template. The new workbook loaded all 42,331 attempts and 1,425 students; its previous-workspace backup retained REHMAN; the new active archive restored the large data correctly in a separate process. Original workbook hashes stayed unchanged.

The current SU26 selection contains 13 attempts, seven students and four departmental sections. PDF/Word/Excel button exports passed. Six CGPA reviews are explained by unresolved historical non-final grades; the application does not manufacture their values.

Private acceptance outputs are kept outside the release. Aggregate import/restart results are included in evidence/workspace-import.json and evidence/workspace-restart.json.
