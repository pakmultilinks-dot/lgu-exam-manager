"""Regression checks for the previously crashing report module."""
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import pytest
from docx import Document

import importer
import reports
from sample_data import generate


def test_pass_fail_word_export_with_chart(tmp_path):
    """Chart export exercised the Python 3.10/3.11 f-string syntax error."""
    sample = generate(12, seed=3)  # fictional data only
    mapped = importer.apply_mapping(sample, importer.suggest_mapping(sample))
    output = tmp_path / "pass_fail.docx"
    reports.pass_fail_summary_word(mapped, str(output))
    assert output.is_file() and output.stat().st_size > 100
    paragraphs = "\n".join(p.text for p in Document(output).paragraphs)
    assert "Result proportion:" in paragraphs
    assert "Pass 8 / Fail 4" in paragraphs


def test_all_python_files_parse_as_python310_and_311():
    """Catch PEP 701-only f-string syntax before breaking older installations."""
    parso = pytest.importorskip("parso")
    base = Path(__file__).resolve().parents[1]
    sources = list(base.glob("*.py")) + list((base / "tests").glob("*.py"))
    assert sources
    for version in ("3.10", "3.11"):
        grammar = parso.load_grammar(version=version)
        for source in sources:
            tree = grammar.parse(source.read_text(encoding="utf-8"), error_recovery=True)
            issues = list(grammar.iter_errors(tree))
            assert not issues, f"Python {version}: {source.name}: {issues}"
