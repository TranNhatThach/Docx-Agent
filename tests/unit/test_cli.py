"""
Tests for CLI commands and JSON output formatting.
"""

from typer.testing import CliRunner
import json
from docx_agent.interfaces.cli.main import app

runner = CliRunner()


def test_cli_inspect(sample_docx):
    result = runner.invoke(app, ["inspect", str(sample_docx), "--json"])
    assert result.exit_code == 0
    data = json.loads(result.stdout)
    assert "paragraphs_count" in data
    assert data["paragraphs_count"] >= 4


def test_cli_read(sample_docx):
    result = runner.invoke(app, ["read", str(sample_docx), "--start", "0", "--end", "2", "--json"])
    assert result.exit_code == 0
    data = json.loads(result.stdout)
    assert "paragraphs" in data
    assert len(data["paragraphs"]) == 2


def test_cli_replace(sample_docx, temp_dir):
    out_file = temp_dir / "cli_replace_out.docx"
    result = runner.invoke(app, [
        "replace", str(sample_docx),
        "--target", "Giới thiệu",
        "--replace", "Tổng Quan",
        "--output", str(out_file),
        "--json",
    ])
    assert result.exit_code == 0
    data = json.loads(result.stdout)
    assert data["success"] is True
    assert data["replaced_count"] >= 1


def test_cli_format_text(sample_docx, temp_dir):
    out_file = temp_dir / "cli_fmt_text_out.docx"
    result = runner.invoke(app, [
        "format-text", str(sample_docx),
        "--target", "p_0001",
        "--font-name", "Times New Roman",
        "--font-size-pt", "18.0",
        "--bold",
        "--output", str(out_file),
        "--json",
    ])
    assert result.exit_code == 0
    data = json.loads(result.stdout)
    assert data["success"] is True


def test_cli_capabilities(sample_docx):
    result = runner.invoke(app, ["capabilities", str(sample_docx), "--json"])
    assert result.exit_code == 0
    data = json.loads(result.stdout)
    assert data["can_safely_edit"] is True
