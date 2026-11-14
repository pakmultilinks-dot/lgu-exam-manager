from excel_io import excel_engine
"""Excel import engine: sheet detection, column mapping (fuzzy auto-suggest)."""
import difflib
import pandas as pd

# Logical roles the app understands. Nothing is hardcoded to real column names;
# the mapping screen lets the user confirm/override every role.
ROLES = [
    ("roll", "Roll Number"),
    ("name", "Student Name"),
    ("cgpa", "CGPA / GPA"),
    ("status", "Result Status (Pass/Fail)"),
    ("subject", "Subject / Marks columns (multi-select)"),
]

# Hints used for fuzzy auto-suggestion of column roles.
ROLE_HINTS = {
    "roll": ["roll", "reg no", "registration", "enroll", "id", "seat"],
    "name": ["name", "student", "candidate"],
    "cgpa": ["cgpa", "gpa", "cumulative gpa", "sgpa"],
    "status": ["status", "result", "pass", "fail", "remarks", "decision"],
    "department": ["department", "dept", "program", "faculty", "school",
                   "discipline"],
}


def load_workbook(path: str) -> list:
    """Return the list of sheet names in the workbook."""
    xl = pd.ExcelFile(path, engine=excel_engine(path))
    return xl.sheet_names


def read_sheet(path: str, sheet, header_row: int = 0) -> pd.DataFrame:
    """Read a sheet, drop fully-empty rows/cols, normalize column labels.

    header_row: 0-based index of the row that holds column headers
    (0 = first row). Lets the user handle files where headers sit in
    row 2 or 3 instead of row 1.
    """
    df = pd.read_excel(path, sheet_name=sheet, engine=excel_engine(path),
                       header=header_row)
    df = df.dropna(axis=0, how="all").dropna(axis=1, how="all")
    df.columns = [str(c).strip() for c in df.columns]
    # Drop duplicate column labels (keep first)
    df = df.loc[:, ~df.columns.duplicated()]
    # Drop "Unnamed" columns that are really blank
    df = df.loc[:, ~df.columns.str.match(r"^Unnamed")]
    return df.reset_index(drop=True)


def _best_role(col: str) -> str | None:
    """Fuzzy-match a column name to a role; None if no decent match."""
    c = col.lower()
    best, best_score = None, 0.0
    for role, hints in ROLE_HINTS.items():
        for hint in hints:
            score = difflib.SequenceMatcher(None, hint, c).ratio()
            # bonus for substring containment
            if hint in c:
                score = max(score, 0.75)
            if score > best_score:
                best, best_score = role, score
    return best if best_score >= 0.55 else None


def _name_score(col: str, role: str) -> float:
    """Score 0-1 of how well a column name matches a role (name hints only)."""
    c = col.lower()
    best = 0.0
    for hint in ROLE_HINTS.get(role, []):
        # Exact word match scores highest
        words = c.replace("_", " ").replace("-", " ").split()
        if hint in words:
            best = max(best, 1.0)
        elif hint == c.strip():
            best = max(best, 1.0)
        elif hint in c:
            # Substring but not a word: weaker (avoids "student" matching "StudentID" too strongly)
            best = max(best, 0.6)
        else:
            score = difflib.SequenceMatcher(None, hint, c).ratio()
            best = max(best, score * 0.8)
    return best


def suggest_mapping(df: pd.DataFrame) -> dict:
    """Return {role: column} auto-suggestions using name hints + data-kind
    intelligence. 'subject' gets a list. Takes the DataFrame (not just names)
    so it can inspect actual values.

    Strategy: combine name-score with data-kind evidence. A column must have
    BOTH reasonable name evidence AND matching data kind for id/name roles.
    """
    columns = list(df.columns)
    kinds = {col: detect_column_kind(df[col]) for col in columns}
    mapping = {}
    used = set()

    def claim(role, col):
        mapping[role] = col
        used.add(col)

    # ROLL: registration numbers can be long, repeated (one row per course),
    # or purely numeric. Header evidence outweighs value shape here.
    best_col, best_score = None, 0.0
    for col in columns:
        score = _name_score(col, "roll")
        normalized = " ".join(col.lower().replace("_", " ").replace("-", " ").split())
        if ("registration" in normalized or "roll" in normalized or
                "reg no" in normalized or "student id" in normalized):
            score = max(score, 1.0)
        if kinds[col] == "id_like":
            score += 0.05
        if score > best_score:
            best_col, best_score = col, score
    if best_col and best_score >= 0.65:
        claim("roll", best_col)

    # NAME: needs name_like data + name-ish column name
    best_col, best_score = None, 0.0
    for col in columns:
        if col in used:
            continue
        if kinds[col] not in ("name_like", "text"):
            continue
        score = _name_score(col, "name")
        if score > best_score:
            best_col, best_score = col, score
    if best_col and best_score >= 0.6:
        claim("name", best_col)

    # CGPA is academic evidence and must never be inferred just because
    # marks or an integer column falls inside 0-4. Require an explicit label.
    best_col, best_score = None, 0.0
    for col in columns:
        if col in used:
            continue
        label = " ".join(col.lower().replace("_", " ").replace("-", " ").split())
        if not any(h in label for h in ("cgpa", "sgpa", "gpa")):
            continue
        if detect_cgpa_scale(df[col]) not in ("cgpa4", "cgpa5", "cgpa10", "percent", "marks"):
            continue
        score = _name_score(col, "cgpa")
        if "cgpa" in label or "sgpa" in label:
            score = 1.0
        if score > best_score:
            best_col, best_score = col, score
    if best_col and best_score >= 0.6:
        claim("cgpa", best_col)

    # STATUS: never interpret a letter-grade column as student pass/fail.
    for col in columns:
        if col in used:
            continue
        label = " ".join(col.lower().replace("_", " ").replace("-", " ").split())
        if label in ("grade", "letter grade", "final grade", "grade point"):
            continue
        vals = df[col].astype(str).str.strip().str.lower().unique()
        # Single-letter grades such as F/P must not masquerade as Pass/Fail.
        pass_like = {"pass", "passed", "fail", "failed", "promoted",
                     "clear", "repeat", "reappear"}
        if any(v in pass_like for v in vals):
            claim("status", col)
            break
    if "status" not in mapping:
        # Fallback: name hint only
        best_col, best_score = None, 0.0
        for col in columns:
            if col in used:
                continue
            label = " ".join(col.lower().replace("_", " ").replace("-", " ").split())
            if not any(h in label for h in ("status", "result", "pass fail", "remarks", "decision")):
                continue
            score = _name_score(col, "status")
            if score > best_score:
                best_col, best_score = col, score
        if best_col and best_score >= 0.7:
            claim("status", best_col)

    # DEPARTMENT: text column with a department-like name (multi-department
    # files). Requires a strong name match AND text data so subject columns
    # with numeric marks are never claimed here.
    best_col, best_score = None, 0.0
    for col in columns:
        if col in used:
            continue
        if kinds[col] not in ("text", "name_like"):
            continue
        score = _name_score(col, "department")
        if score > best_score:
            best_col, best_score = col, score
    if best_col and best_score >= 0.7:
        claim("department", best_col)

    # Everything unclaimed becomes a subject candidate
    mapping["subject"] = [c for c in columns if c not in used]
    return mapping


def apply_mapping(df: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    """Return a normalized working copy with standard internal column names.

    mapping: {"roll": col, "name": col, "cgpa": col|None, "status": col|None,
              "department": col|None, "subject": [cols]}
    """
    out = pd.DataFrame()
    out["_roll"] = df[mapping["roll"]].astype(str) if mapping.get("roll") else ""
    out["_name"] = df[mapping["name"]].astype(str) if mapping.get("name") else ""
    if mapping.get("cgpa"):
        # Intelligent: auto-detects 4.0 / 5.0 / 10.0 / percent / letter grades
        out["_cgpa"] = normalize_cgpa(df[mapping["cgpa"]])
        out.attrs["cgpa_scale"] = detect_cgpa_scale(df[mapping["cgpa"]])
    else:
        out["_cgpa"] = pd.NA
        out.attrs["cgpa_scale"] = None
    if mapping.get("status"):
        out["_status"] = normalize_status(df[mapping["status"]])
    else:
        out["_status"] = ""
    if mapping.get("department"):
        out["_department"] = df[mapping["department"]].astype(str).str.strip()
    else:
        out["_department"] = ""
    out["_subjects"] = ""
    subj_cols = mapping.get("subject") or []
    out.attrs["subject_cols"] = list(subj_cols)
    out.attrs["raw"] = df  # keep original for filtered-view export
    return out


def normalize_status(series: pd.Series) -> pd.Series:
    """Best-effort Pass/Fail normalization from free-text status values."""
    s = series.astype(str).str.strip().str.lower()
    def conv(v):
        if v in ("pass", "passed", "p", "promoted", "clear"):
            return "Pass"
        if v in ("fail", "failed", "f", "r", "repeat", "reappear"):
            return "Fail"
        return v.title() if v and v != "nan" else ""
    return s.map(conv)


# ---------------------------------------------------------------------------
# Field-level intelligence: detect what KIND of data each column holds,
# regardless of its name. The app adapts to the data, not the other way around.
# ---------------------------------------------------------------------------

def detect_cgpa_scale(series: pd.Series) -> str:
    """Detect whether a numeric column is 4.0 CGPA, 10.0 GPA, %, or marks.

    Returns one of: 'cgpa4', 'cgpa5', 'cgpa10', 'percent', 'marks', 'grades', 'unknown'.

    Smart guards: all-integer small numbers (like roll numbers 1,2,3) are
    NOT treated as CGPA. Real CGPA columns have decimal values.
    """
    nums = pd.to_numeric(series, errors="coerce").dropna()
    if nums.empty:
        # Maybe letter grades?
        vals = series.astype(str).str.strip().str.upper().unique()
        grade_like = {"A+", "A", "A-", "B+", "B", "B-", "C+", "C", "C-",
                      "D+", "D", "F"}
        if any(v in grade_like for v in vals):
            return "grades"
        return "unknown"
    mx = nums.max()
    mn = nums.min()
    # Guard: if ALL values are integers and max <= 10, it's probably an ID
    # or count, not a CGPA. Real CGPAs have decimals.
    all_int = ((nums % 1) == 0).all()
    if mx <= 4.0 and mn >= 0:
        if all_int and len(nums) > 3:
            return "unknown"  # likely IDs like 1,2,3
        return "cgpa4"
    if mx <= 5.0 and mn >= 0:
        if all_int and len(nums) > 3:
            return "unknown"
        return "cgpa5"
    if mx <= 10.0 and mn >= 0:
        if all_int and len(nums) > 3:
            return "unknown"
        return "cgpa10"
    if mx <= 100.0 and mn >= 0:
        return "percent"
    return "marks"


_GRADE_TO_CGPA = {
    "A+": 4.0, "A": 4.0, "A-": 3.7,
    "B+": 3.3, "B": 3.0, "B-": 2.7,
    "C+": 2.3, "C": 2.0, "C-": 1.7,
    "D+": 1.3, "D": 1.0, "F": 0.0,
}


def normalize_cgpa(series: pd.Series) -> pd.Series:
    """Convert a CGPA/GPA column to pure numbers WITHOUT changing values.

    HARD RULE: never invent or approximate data. This only converts the
    dtype to numeric (e.g. "3.5" string -> 3.5 float). Values are NEVER
    altered, scaled, or converted between systems.

    Percentages stay as percentages. Letter grades are NOT converted
    (returned as NaN so reports show the original text instead).
    Use detect_cgpa_scale() to label what the numbers mean.
    """
    scale = detect_cgpa_scale(series)
    if scale in ("cgpa4", "cgpa5", "cgpa10", "percent", "marks"):
        return pd.to_numeric(series, errors="coerce")
    # grades / unknown: do not fabricate numbers, return NaN
    return pd.Series([pd.NA] * len(series), index=series.index,
                     dtype="Float64")


def detect_column_kind(series: pd.Series) -> str:
    """Classify a column: 'id_like', 'name_like', 'numeric', 'date', 'text'.

    Used to give smarter auto-suggestions and smarter per-column handling.
    """
    s = series.dropna()
    if s.empty:
        return "text"
    strs = s.astype(str)
    # Pure numbers? Check this BEFORE id_like, because numeric columns
    # (marks, CGPA) with few rows look "unique" too.
    nums = pd.to_numeric(s, errors="coerce")
    is_numeric = nums.notna().mean() > 0.8
    # ID-like: unique, NO spaces, short, and NOT purely numeric
    # (real IDs are alphanumeric codes like LGU-001, not plain numbers)
    uniq_ratio = strs.nunique() / max(len(strs), 1)
    avg_len = strs.str.len().mean()
    has_spaces = strs.str.contains(" ", na=False).mean()
    has_letters = strs.str.contains("[A-Za-z]", na=False).mean()
    if uniq_ratio > 0.9 and has_spaces < 0.1 and avg_len < 20:
        if not is_numeric or has_letters > 0.5:
            return "id_like"
    # Name-like: has spaces (full names), regardless of uniqueness
    if has_spaces > 0.5:
        return "name_like"
    # Date?
    try:
        if strs.str.match(r"\d{1,4}[-/]\d{1,2}[-/]\d{1,4}").mean() > 0.5:
            pd.to_datetime(s.iloc[:20], errors="raise")
            return "date"
    except Exception:
        pass
    # Numeric?
    if is_numeric:
        return "numeric"
    return "text"
