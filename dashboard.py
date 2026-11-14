"""Dashboard statistics computed from the mapped dataframe."""
import pandas as pd
from importer import normalize_status


def _score_distribution(scores: pd.Series, scale: str | None) -> dict:
    """Bucket only scales with meaningful, known ranges; never infer marks bands."""
    if scale == "percent":
        ranges = [
            ("90-100", 90), ("80-89", 80), ("70-79", 70),
            ("60-69", 60), ("50-59", 50), ("Below 50", None),
        ]
    elif scale == "cgpa4":
        ranges = [
            ("4.0", 4.0), ("3.5-3.99", 3.5), ("3.0-3.49", 3.0),
            ("2.5-2.99", 2.5), ("2.0-2.49", 2.0), ("Below 2.0", None),
        ]
    elif scale == "cgpa5":
        ranges = [
            ("4.5-5.0", 4.5), ("4.0-4.49", 4.0), ("3.5-3.99", 3.5),
            ("3.0-3.49", 3.0), ("2.5-2.99", 2.5), ("Below 2.5", None),
        ]
    elif scale == "cgpa10":
        ranges = [
            ("9.0-10.0", 9.0), ("8.0-8.99", 8.0), ("7.0-7.99", 7.0),
            ("6.0-6.99", 6.0), ("5.0-5.99", 5.0), ("Below 5.0", None),
        ]
    else:
        return {}

    buckets = {label: 0 for label, _ in ranges}
    for score in scores:
        for label, lower_bound in ranges:
            if lower_bound is None or score >= lower_bound:
                buckets[label] += 1
                break
    return buckets


def pass_percentage_label(stats: dict) -> str:
    """Make the denominator visible when some status values are unclassified."""
    count = stats["status_count"]
    if count < stats["total"]:
        return f"Pass % (of {count} classified statuses)"
    return "Pass %"


def compute_stats(df: pd.DataFrame) -> dict:
    """df: normalized frame from importer.apply_mapping.

    DATA INTEGRITY RULE: never invent data. Pass/fail comes ONLY from an
    actual status column. Averages are computed on original values only.
    Labels adapt to the detected scale (CGPA / Percentage / etc).
    """
    stats = {"total": len(df)}
    cgpa = pd.to_numeric(df["_cgpa"], errors="coerce")
    valid = cgpa.dropna()
    has_cgpa = not valid.empty
    scale = df.attrs.get("cgpa_scale")

    # Human label for the grade column based on detected scale
    scale_labels = {
        "cgpa4": "CGPA", "cgpa5": "CGPA (5.0)",
        "cgpa10": "GPA (10.0)", "percent": "Percentage",
        "marks": "Marks", "grades": "Grade",
    }
    stats["grade_label"] = scale_labels.get(scale, "Score")
    stats["numeric_score_count"] = len(valid)
    stats["numeric_score_coverage_pct"] = (
        round(100 * len(valid) / len(df), 1) if len(df) else None
    )

    status = normalize_status(df["_status"])
    passes = int((status == "Pass").sum())
    fails = int((status == "Fail").sum())
    classified_statuses = passes + fails
    status_missing = int(status.eq("").sum())
    status_unclassified = len(df) - classified_statuses - status_missing
    if not classified_statuses:
        # Different programs have different passing criteria; never derive status from scores.
        passes, fails = None, None
    stats["passes"] = passes
    stats["fails"] = fails
    stats["has_passfail"] = classified_statuses > 0
    stats["status_count"] = classified_statuses
    stats["status_missing"] = status_missing
    stats["status_unclassified"] = status_unclassified
    stats["status_coverage_pct"] = (
        round(100 * classified_statuses / len(df), 1) if len(df) else None
    )
    stats["pass_pct"] = (
        round(100 * passes / classified_statuses, 1)
        if passes is not None and classified_statuses else None
    )
    stats["has_cgpa"] = has_cgpa

    stats["avg_cgpa"] = round(float(valid.mean()), 2) if len(valid) else None
    stats["max_cgpa"] = round(float(valid.max()), 2) if len(valid) else None
    stats["min_cgpa"] = round(float(valid.min()), 2) if len(valid) else None

    # Missing grades cannot be ranked. Stable sorting preserves source order for ties.
    work = df.copy()
    work["_cgpa_num"] = cgpa
    work = work.loc[work["_cgpa_num"].notna()]
    merit = work.sort_values("_cgpa_num", ascending=False, kind="mergesort").head(10)
    stats["merit"] = [
        {"roll": r["_roll"], "name": r["_name"],
         "cgpa": (None if pd.isna(r["_cgpa_num"]) else round(float(r["_cgpa_num"]), 2)),
         "status": str(r["_status"])}
        for _, r in merit.iterrows()
    ]

    stats["distribution"] = _score_distribution(valid, scale)
    stats["_status_norm"] = status
    return stats


def subject_performance(df: pd.DataFrame, subject_cols: list) -> list:
    """Mean / max / min per subject column from the raw frame.

    NOTE: no pass-rate is computed here because the passing threshold
    varies by subject/program. The app never invents thresholds.
    """
    raw = df.attrs.get("raw")
    if raw is None:
        return []
    out = []
    for col in subject_cols:
        vals = pd.to_numeric(raw[col], errors="coerce").dropna()
        if len(vals) == 0:
            continue
        out.append({
            "subject": col,
            "avg": round(float(vals.mean()), 1),
            "max": round(float(vals.max()), 1),
            "min": round(float(vals.min()), 1),
            "count": len(vals),
        })
    return out
