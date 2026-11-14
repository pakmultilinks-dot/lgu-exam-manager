"""Generate a realistic sample Excel result file for testing the LGU Exam App."""
import random
import pandas as pd

FIRST = ["Ahmed", "Ali", "Bilal", "Danish", "Fahad", "Hamza", "Imran", "Junaid",
         "Kashif", "Mohsin", "Nabeel", "Owais", "Rizwan", "Sajid", "Tariq",
         "Usman", "Waqas", "Yasir", "Zubair", "Adeel", "Asad", "Farhan",
         "Ayesha", "Fatima", "Hina", "Iqra", "Mahnoor", "Sana", "Zainab",
         "Rabia", "Sadia", "Tania", "Uzma", "Kiran", "Nadia", "Shazia"]
LAST = ["Khan", "Ahmed", "Ali", "Malik", "Raza", "Hussain", "Shah", "Butt",
        "Sheikh", "Qureshi", "Farooq", "Iqbal", "Akram", "Aslam", "Naseer"]

SUBJECTS = ["CS-101 Programming Fundamentals", "CS-102 Discrete Structures",
            "MTH-101 Calculus I", "ENG-101 English Composition",
            "PHY-101 Applied Physics"]


def _marks_to_gp(marks: float) -> float:
    if marks >= 85:
        return 4.0
    if marks >= 80:
        return 3.7
    if marks >= 75:
        return 3.3
    if marks >= 70:
        return 3.0
    if marks >= 65:
        return 2.7
    if marks >= 60:
        return 2.3
    if marks >= 55:
        return 2.0
    if marks >= 50:
        return 1.7
    return 0.0


def generate(n: int = 50, seed: int = 42) -> pd.DataFrame:
    random.seed(seed)
    rows = []
    for i in range(1, n + 1):
        name = f"{random.choice(FIRST)} {random.choice(LAST)}"
        roll = f"LGU-CS-2025-{i:03d}"
        # Most students decent, a few fail
        base = random.gauss(68, 14)
        marks = [max(0, min(100, round(random.gauss(base, 8), 1)))
                 for _ in SUBJECTS]
        gps = [_marks_to_gp(m) for m in marks]
        cgpa = round(sum(gps) / len(gps), 2)
        failed = sum(1 for m in marks if m < 50)
        status = "Fail" if failed > 0 else ("Pass" if cgpa >= 2.0 else "Probation")
        row = {"Roll Number": roll, "Student Name": name}
        for subj, m in zip(SUBJECTS, marks):
            row[subj] = m
        row["CGPA"] = cgpa
        row["Status"] = status
        rows.append(row)
    return pd.DataFrame(rows)


def save_sample(path: str = "sample_result.xlsx", n: int = 50) -> str:
    df = generate(n)
    df.to_excel(path, index=False, sheet_name="Final Result")
    return path


if __name__ == "__main__":
    import sys
    out = sys.argv[1] if len(sys.argv) > 1 else "sample_result.xlsx"
    save_sample(out)
    print(f"Sample file written: {out}")
