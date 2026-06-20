"""CLI: pyvbaanalysis over loose files and folders."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyvbaanalysis import __version__
from pyvbaanalysis.cli import main

_DIRTY = 'Attribute VB_Name = "Mod1"\r\nPublic Sub S()\r\n    Dim n As Long\r\n    n = "x"\r\nEnd Sub\r\n'
_CLEAN = 'Attribute VB_Name = "Mod2"\r\nOption Explicit\r\n\r\nPublic Sub T()\r\nEnd Sub\r\n'


def _write(path: Path, content: str) -> Path:
    path.write_bytes(content.encode("cp1252"))
    return path


def test_cli_reports_diagnostics_and_exits_nonzero(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    _write(tmp_path / "Mod1.bas", _DIRTY)
    code = main([str(tmp_path)])
    out = capsys.readouterr().out
    assert code == 1
    assert "assignment-type-mismatch" in out
    assert "Mod1" in out


def test_cli_clean_project_exits_zero(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    _write(tmp_path / "Mod2.bas", _CLEAN)
    code = main([str(tmp_path)])
    assert code == 0
    assert "no diagnostics" in capsys.readouterr().out


def test_cli_json_format(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    _write(tmp_path / "Mod1.bas", _DIRTY)
    main([str(tmp_path), "--format", "json"])
    payload = json.loads(capsys.readouterr().out)
    assert payload[0]["project"] == "(loose files)"
    module = next(m for m in payload[0]["modules"] if m["module"] == "Mod1")
    codes = {d["code"] for d in module["diagnostics"]}
    assert "assignment-type-mismatch" in codes
    # Each diagnostic carries a 1-based line and an offset span.
    sample = module["diagnostics"][0]
    assert sample["line"] >= 1 and sample["start"] >= 0


def test_cli_json_schema_is_stable(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    # The JSON shape is a 1.0.0 compatibility promise; guard the exact field set.
    _write(tmp_path / "Mod1.bas", _DIRTY)
    main([str(tmp_path), "--format", "json"])
    payload = json.loads(capsys.readouterr().out)
    assert set(payload[0]) == {"project", "modules"}
    module = payload[0]["modules"][0]
    assert set(module) == {"module", "diagnostics"}
    diagnostic = module["diagnostics"][0]
    assert set(diagnostic) == {
        "code",
        "severity",
        "message",
        "start",
        "end",
        "line",
        "column",
        "spec_reference",
    }


def test_cli_severity_override_silences(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    _write(tmp_path / "Mod2.bas", _CLEAN.replace("Option Explicit\r\n", ""))  # drop Option Explicit
    code = main([str(tmp_path), "--severity", "option-explicit-missing=off"])
    out = capsys.readouterr().out
    assert code == 0
    assert "option-explicit-missing" not in out


def test_cli_bad_severity_code_exits_two(capsys: pytest.CaptureFixture[str], tmp_path: Path) -> None:
    _write(tmp_path / "Mod1.bas", _DIRTY)
    code = main([str(tmp_path), "--severity", "not-a-code=off"])
    assert code == 2
    assert "unknown diagnostic code" in capsys.readouterr().err


def test_cli_fail_level_error_ignores_warnings(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    # Mod2 without Option Explicit: only a warning, so --fail-level error exits 0.
    _write(tmp_path / "Mod2.bas", _CLEAN.replace("Option Explicit\r\n", ""))
    code = main([str(tmp_path), "--fail-level", "error"])
    assert code == 0
    assert "option-explicit-missing" in capsys.readouterr().out  # still reported


def test_cli_ignore_code_drops_it_from_report_and_exit(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    _write(tmp_path / "Mod1.bas", _DIRTY)
    code = main(
        [str(tmp_path), "--ignore", "assignment-type-mismatch", "--fail-level", "error"]
    )
    out = capsys.readouterr().out
    assert "assignment-type-mismatch" not in out
    assert code == 0  # the only error was ignored


def test_cli_select_keeps_only_named_codes(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    _write(tmp_path / "Mod1.bas", _DIRTY)
    main([str(tmp_path), "--select", "assignment-type-mismatch", "--format", "json"])
    payload = json.loads(capsys.readouterr().out)
    codes = {d["code"] for m in payload[0]["modules"] for d in m["diagnostics"]}
    assert codes == {"assignment-type-mismatch"}


def test_cli_severity_override_is_case_insensitive(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    # A mis-cased code must still apply, not validate-then-silently-do-nothing.
    _write(tmp_path / "Mod2.bas", _CLEAN.replace("Option Explicit\r\n", ""))
    code = main([str(tmp_path), "--severity", "Option-Explicit-Missing=off"])
    out = capsys.readouterr().out
    assert code == 0
    assert "option-explicit-missing" not in out


def test_cli_select_is_case_insensitive(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    _write(tmp_path / "Mod1.bas", _DIRTY)
    code = main([str(tmp_path), "--select", "Assignment-Type-Mismatch", "--fail-level", "error"])
    out = capsys.readouterr().out
    assert "assignment-type-mismatch" in out
    assert code == 1  # the selected error still counts toward the exit code


def test_cli_unknown_select_code_exits_two(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    # An unknown --select code must fail loudly, not silently drop everything and pass.
    _write(tmp_path / "Mod1.bas", _DIRTY)
    code = main([str(tmp_path), "--select", "not-a-real-code"])
    assert code == 2
    assert "unknown diagnostic code" in capsys.readouterr().err


def test_cli_folder_with_matching_named_subdir_does_not_crash(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    # A sub-directory whose name ends in .bas must be skipped, not loaded as a file.
    (tmp_path / "sub.bas").mkdir()
    _write(tmp_path / "Mod1.bas", _DIRTY)
    code = main([str(tmp_path)])
    assert code == 1
    assert "assignment-type-mismatch" in capsys.readouterr().out


def test_cli_only_filters_modules(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    _write(tmp_path / "Mod1.bas", _DIRTY)
    _write(tmp_path / "Mod2.bas", _CLEAN)
    main([str(tmp_path), "--only", "Mod2", "--format", "json"])
    payload = json.loads(capsys.readouterr().out)
    modules = {m["module"] for m in payload[0]["modules"]}
    assert modules == {"Mod2"}


def test_cli_version(capsys: pytest.CaptureFixture[str]) -> None:
    with pytest.raises(SystemExit) as exc:
        main(["--version"])
    assert exc.value.code == 0
    assert __version__ in capsys.readouterr().out


def test_cli_missing_path_exits_two(capsys: pytest.CaptureFixture[str]) -> None:
    code = main(["does_not_exist_12345.bas"])
    assert code == 2
    assert "not found" in capsys.readouterr().err


def test_cli_single_loose_file(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    path = _write(tmp_path / "Mod1.bas", _DIRTY)
    code = main([str(path)])
    assert code == 1
    assert "assignment-type-mismatch" in capsys.readouterr().out


def test_cli_duplicate_module_name_warns(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    # Two folders each with a Util.bas (same VB_Name) pooled into one project:
    # warn and analyze the first, rather than silently dropping one.
    clean_util = 'Attribute VB_Name = "Util"\r\nOption Explicit\r\nPublic Sub A()\r\nEnd Sub\r\n'
    a = tmp_path / "a"
    a.mkdir()
    b = tmp_path / "b"
    b.mkdir()
    _write(a / "Util.bas", clean_util)
    _write(b / "Util.bas", clean_util)
    code = main([str(a), str(b)])
    err = capsys.readouterr().err
    assert "duplicate module name" in err
    assert code == 0  # the surviving Util analyzed cleanly


def test_cli_unreadable_workbook_solo_exits_one(
    tmp_path: Path, capsys: pytest.CaptureFixture[str], monkeypatch: pytest.MonkeyPatch
) -> None:
    # A solo unreadable workbook is a read error (exit 1), not a no-input usage error.
    from pyvbaanalysis import cli as cli_mod
    from pyvbaanalysis.reader import WorkbookReadError

    bad = tmp_path / "broken.xlsm"
    bad.write_bytes(b"x")

    def _boom(_path: object) -> object:
        raise WorkbookReadError("cannot read broken.xlsm")

    monkeypatch.setattr(cli_mod, "read_workbook_modules", _boom)
    code = main([str(bad)])
    assert code == 1
    assert "cannot read broken.xlsm" in capsys.readouterr().err
