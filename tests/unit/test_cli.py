"""The installed environment check describes the implemented product accurately."""

import sys

from excel_csv_automation.__main__ import main


def test_environment_check_reports_ready_processing_support(monkeypatch, capsys):
    monkeypatch.setattr(sys, "argv", ["excel-csv-automation", "--check"])
    assert main() == 0
    output = capsys.readouterr().out
    assert "streamlit:" in output
    assert "Environment ready" in output
    assert "process CSV and XLSX files" in output
    assert "not implemented" not in output
