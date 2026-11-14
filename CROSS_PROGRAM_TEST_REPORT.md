# LGU Exam Manager v3.8: expanded verification

Verified on 23 October 2026. This report distinguishes application failures from unresolved source records.

| Area | Coverage | Result |
| --- | --- | --- |
| Automated regression suite | 127 tests | PASS |
| Grade/credit matrix | 72 normalized reference/configured programs × 10 confirmed grades × 7 credit values = 5,040 cases | Exact weighted points, totals and two-place truncation matched |
| Independent randomized repeat oracle | 500 students, 7,500 attempts, 3,000 term/cumulative scopes | PASS |
| Actual desktop controls | 27 programs, 13 department layouts, 108 fictional students, 648 attempts | 85 export button commands passed |
| Defaulters | 81 fictional fee defaulters, including multiple fee pages | Withheld from academic body; all fee pages follow all academic sections |
| Academic semester numbers | All 1–14; invalid 0, -1, 15, fractional and text inputs | Valid numbers preserved; invalid numbers block publication |
| Large workbook | 1,425 students, 42,331 attempts, 26 programs, 18 terms | Import, navigation, paging, search, save/reload, review exports and preview passed |
| Complete real-workbook arithmetic | 8,253 GPA + 7,577 CGPA numeric scopes | Zero mismatches against independent exact Fraction calculations |
| Unresolved arithmetic scopes | 1,614 term/cumulative scopes | Both oracle and app require review; no invented numeric result |
| Every available report route | 11 routes × PDF/Word/Excel | 33 dispatcher outputs and background batch pass |
| Gazette headers | All 120 measured profile geometries preserved | Urdu identity corrected; 21 omitted university headings restored from original PDF spans |

## Failures found and corrected

Six historical program labels lacked automatic routing: BS BCH, BS ZOO, CERUF, M.Phil Urdu, MBA (2 Years), and MSBA. Their routes now live in editable JSON and original degree labels remain unchanged.

Visual verification found Urdu headings inherited from a shared International Relations grid profile. Program-specific department headings now apply to the cover, academic pages and fee pages. Mixed Arts/Urdu fee sections use an encompassing departmental heading. Missing university-title spans were restored from the exact original reference pages; table geometry and catalog routing were preserved.

## UI cases checked

The desktop matrix invokes the actual Tk export buttons, waits for background jobs, and reads back PDF, Word and Excel files. It checks every program filter, all navigation sections, every paging boundary, current-term selection in Settings, the interaction with Academic Records term filters, empty searches, unavailable terms, restoring valid selections, combined PDF preview, and session save/reload.

Each individual program PDF is checked for current-term course names, absence of earlier-only course names, independently expected GPA 3.36 and CGPA 3.16, and a final fee appendix. Urdu PDFs also reject the International Relations heading. The combined PDF is checked for every defaulter name, absence of those names before the fee section, and a contiguous fee-only tail.

## Data that still needs review

The original large workbook still contains 29 reported source-data errors (quarantined rows and identity/name conflicts). These are retained in its review audit and block affected selections. Passing the arithmetic audit does not reconcile contradictory source identities.

Non-final grades such as IP, I, W and PASS, unknown U rates, absent numeric values and zero/nonpositive-credit cases retain REVIEW where the calculation policy is insufficient. The app does not assign them an assumed grade-point rate or remove them silently. F uses the user-confirmed 0.00 rate. Original workbook cells and imported source attempts are preserved.

## Degree routing coverage

| Original workbook degree label | Automatic department template |
| --- | --- |
| BBA (Hons) | management-science |
| BS AF | management-science |
| BS AP | psychology |
| BS BCH | allied-health-biology |
| BS BT | allied-health-biology |
| BS CHEM. | chemistry |
| BS DFCS | management-science |
| BS Eng. | english |
| BS IR | art-and-humanities |
| BS IT | information-technology |
| BS Maths | mathematics |
| BS MB | allied-health-biology |
| BS MC | media-and-communication |
| BS PHYS | physics |
| BS SE | software-engineering |
| BS ZOO | allied-health-biology |
| BSCS | computer-science |
| CERUF | art-and-humanities |
| M.Phil CHEM. | chemistry |
| M.Phil MB | allied-health-biology |
| M.Phil Urdu | art-and-humanities |
| M.Phil ZOO | allied-health-biology |
| MBA (2 Years) | management-science |
| MSBA | management-science |
| MSCP | psychology |
| Ph. D URDU | art-and-humanities |

## Reference sources for additional historical routes

- [LGU faculty handbook: Biology and Management Sciences programs](https://lgu.edu.pk/wp-content/uploads/2021/07/FACULTY-EMPLOYEE-HAND-BOOK.pdf).
- [LGU Urdu programs](https://lgu.edu.pk/department-of-urdu-programs/).
- [LGU CERUF program](https://lgu.edu.pk/certificate-in-urdu-for-foreign-studentsceruf/).
- [LGU program fee structure: MBA (2 Years) naming](https://lgu.edu.pk/program-fee-structure/).

These sources establish program names and departmental associations only. Academic values and fee status come from the supplied workbooks. The department grid and header geometry come from the 13 supplied reference PDFs.

## Reproduce the checks

```powershell
.venv\Scripts\python.exe -m pytest -q
.venv\Scripts\python.exe tests/accept_matrix_local.py
```

From the extracted Windows release, run `CHECK_LGU_Exam_Manager.bat` for the packaged executable self-check. That check uses fictional records and does not load personal autosaves. The JSON evidence files contain aggregate counts; real workbooks and student sessions are excluded from the release.
