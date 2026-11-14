"""Diagnostic: check what's broken in LGU Exam Manager. Run: python diagnose.py"""
import sys

print("=== LGU Exam Manager Diagnostic ===")
print(f"Python: {sys.version.split()[0]}")
print()

deps = {
    "pandas": "pandas",
    "openpyxl": "openpyxl",
    "customtkinter": "customtkinter",
    "fpdf": "fpdf2",
    "docx": "python-docx",
    "matplotlib": "matplotlib",
}

missing = []
for name, pkg in deps.items():
    try:
        __import__(name)
        print(f"OK: {pkg}")
    except ImportError:
        print(f"MISSING: {pkg}")
        missing.append(pkg)

print()
if missing:
    print(f"Install missing: pip install {' '.join(missing)}")
else:
    print("All dependencies OK")

print()
for mod in ["importer", "dashboard", "reports"]:
    try:
        __import__(mod)
        print(f"OK: {mod}.py loads")
    except Exception as e:
        print(f"FAIL: {mod}.py - {e}")

print()
print("=== Test import ===")
try:
    import importer
    df = importer.read_sheet('sample_result.xlsx', 'Final Result', header_row=0)
    print(f"OK: Read {len(df)} rows")
    s = importer.suggest_mapping(df)
    m = {'roll': s.get('roll'), 'name': s.get('name'), 'cgpa': s.get('cgpa'),
         'status': s.get('status'), 'subject': list(s.get('subject', []))}
    out = importer.apply_mapping(df, m)
    print(f"OK: Mapped {len(out)} rows")
    print(f"OK: Sample: {out['_roll'].iloc[0]} / {out['_name'].iloc[0]} / CGPA {out['_cgpa'].iloc[0]}")
except Exception:
    import traceback
    print("FAIL:")
    traceback.print_exc()

print("\n=== Done ===")
