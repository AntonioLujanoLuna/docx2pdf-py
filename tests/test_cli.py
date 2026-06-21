"""Command-line behavior and user-facing reporting."""

import pytest

from docx2pdf_py import ConversionResult, cli


def test_cli_reports_engine_actually_used(tmp_path, monkeypatch, capsys):
    source = tmp_path / "input.docx"
    source.write_bytes(b"placeholder")
    output = tmp_path / "output.pdf"
    monkeypatch.setattr(
        cli,
        "convert_detailed",
        lambda *args, **kwargs: ConversionResult(str(output), "weasyprint"),
    )

    assert cli.main([str(source), str(output)]) == 0
    assert "engine: weasyprint" in capsys.readouterr().out


def test_cli_refuses_to_overwrite_without_force(tmp_path):
    source = tmp_path / "input.docx"
    output = tmp_path / "output.pdf"
    source.write_bytes(b"placeholder")
    output.write_bytes(b"existing")

    with pytest.raises(SystemExit) as exc:
        cli.main([str(source), str(output)])
    assert exc.value.code == 2


def test_cli_quiet_suppresses_success_output(tmp_path, monkeypatch, capsys):
    source = tmp_path / "input.docx"
    source.write_bytes(b"placeholder")
    output = tmp_path / "output.pdf"
    monkeypatch.setattr(
        cli,
        "convert_detailed",
        lambda *args, **kwargs: ConversionResult(
            str(output), "libreoffice", ("word failed",)
        ),
    )

    assert cli.main([str(source), str(output), "--quiet"]) == 0
    captured = capsys.readouterr()
    assert captured.out == ""
    assert captured.err == ""
