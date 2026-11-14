"""Optional report-function smoke test. No GUI and no official approval implied."""
from pathlib import Path
import csv
import sys
import tempfile

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import importer
import reports
from sample_data import generate


def main():
    df = generate(8, seed=3)
    mapping = importer.suggest_mapping(df)
    working = importer.apply_mapping(df, mapping)
    working.attrs["settings"] = {
        "topper_count": 3, "department": "COMPUTER SCIENCE",
        "semester": "SPRING", "session": "2026",
        "declaration_date": "2026-05-20", "passing_marks": "", "probation_cgpa": "",
    }
    working.attrs["compare_df"] = working.copy()
    working.attrs["compare_mapping"] = mapping
    working.attrs["compare_label"] = "Fake previous term"
    working.attrs["current_label"] = "Fake current term"
    failures = []
    passed = []
    with tempfile.TemporaryDirectory() as d:
        for key in reports.REPORTS:
            for fmt, (_, ext) in reports.FORMATS.items():
                if key == "lgu_official":
                    continue  # locked in GUI pending official approval
                target = str(Path(d) / f"{key}{ext}")
                try:
                    reports.generate(key, fmt, working, mapping, target)
                    if not Path(target).is_file() or Path(target).stat().st_size < 100:
                        raise AssertionError("No usable output written")
                    passed.append((key, fmt))
                except Exception as e:
                    failures.append((key, fmt, type(e).__name__, str(e)[:150]))
    print(f"legacy report dispatcher: {len(passed)} succeeded, {len(failures)} failed; "
          "official-layout exporter excluded (locked in GUI)")
    for item in failures:
        print("FAILED", *item)
    return 0 if not failures else 1


if __name__ == "__main__":
    raise SystemExit(main())
