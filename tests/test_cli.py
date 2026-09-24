"""CLI: pyvbaanalysis over loose files and folders."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from pyvbaanalysis import __version__
from pyvbaanalysis.cli import main

_DIRTY = 'Attribute VB_Name = "Mod1"\r\nPublic Sub S()\r\n    Dim n As Long\r\n    n = "x"\r\nEnd Sub\r\n'
_CLEAN = 'Attribute VB_Name = "Mod2"\r\nOption Explicit\r\n\r\nPublic Sub T()\r\nEnd Sub\r\n'
# One module body to put behind each export's designer block. Its only "x" is an
# assignment-type-mismatch on the body's sixth line, at column 9.
_BODY = (
    'Attribute VB_Name = "{name}"\r\nOption Explicit\r\n\r\n'
    'Public Sub S()\r\n    Dim n As Long\r\n    n = "x"\r\nEnd Sub\r\n'
)
_CLASS_BLOCK = "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND\r\n"
_FORM_BLOCK = (
    "VERSION 5.00\r\nBegin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form1 \r\n"
    '   Caption = "Form1"\r\n   Begin Forms.CommandButton b\r\n      Caption = "OK"\r\n'
    "   End\r\nEnd\r\n"
)


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


@pytest.mark.parametrize(
    ("file_name", "designer_block", "line"),
    [("Plain.bas", "", 6), ("Widget.cls", _CLASS_BLOCK, 10), ("Form1.frm", _FORM_BLOCK, 13)],
    ids=["bas", "cls", "frm"],
)
def test_cli_positions_are_in_the_file_as_saved(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
    file_name: str,
    designer_block: str,
    line: int,
) -> None:
    # The reader strips a .cls or .frm export's designer block before analysis, and
    # the report still counts it: lines, columns and offsets are the file's own.
    text = designer_block + _BODY.format(name=Path(file_name).stem)
    args = [str(_write(tmp_path / file_name, text)), "--select", "assignment-type-mismatch"]
    main(args)
    assert f"    {line}:9 error assignment-type-mismatch" in capsys.readouterr().out
    main([*args, "--format", "json"])
    [diagnostic] = json.loads(capsys.readouterr().out)[0]["modules"][0]["diagnostics"]
    assert (diagnostic["line"], diagnostic["column"]) == (line, 9)
    assert diagnostic["start"] == text.index('"x"')
    assert text[diagnostic["start"] : diagnostic["end"]] == '"x"'


def test_cli_positions_in_a_container_are_in_its_module_text(
    tmp_path: Path, capsys: pytest.CaptureFixture[str], monkeypatch: pytest.MonkeyPatch
) -> None:
    # pyOpenVBA hands over a module's stored text, which starts at its Attribute
    # lines; the report numbers that text as it numbers a file. A designer block,
    # were a reader to hand one over, is counted the same way.
    from pyvbaanalysis import cli as cli_mod
    from pyvbaanalysis.reader import OfficeProject, loaded_module_from_text

    stored = loaded_module_from_text(_BODY.format(name="Stored"), pyopenvba_standard=True)
    exported = loaded_module_from_text(
        _CLASS_BLOCK + _BODY.format(name="Exported"), pyopenvba_standard=False
    )
    project = OfficeProject(modules=[stored, exported], host="excel", referenced_hosts=[])
    monkeypatch.setattr(cli_mod, "read_office_project", lambda _path: project)
    book = tmp_path / "book.xlsm"
    book.write_bytes(b"x")
    main([str(book), "--select", "assignment-type-mismatch", "--format", "json"])
    modules = json.loads(capsys.readouterr().out)[0]["modules"]
    lines = {m["module"]: [d["line"] for d in m["diagnostics"]] for m in modules}
    assert lines == {"Stored": [6], "Exported": [10]}


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


def test_cli_single_file_skips_cross_module_rules(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    # A single targeted file is a partial view: cross-module rules are suppressed.
    src = (
        'Attribute VB_Name = "Mod1"\r\nOption Explicit\r\n'
        "Public Sub S()\r\n    Call HelperElsewhere\r\nEnd Sub\r\n"
    )
    path = _write(tmp_path / "Mod1.bas", src)
    code = main([str(path)])
    out = capsys.readouterr().out
    assert "unknown-call" not in out
    assert code == 0


def test_cli_partial_project_flag(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    # Two files (a folder) are a whole project by default, so a call to a sub absent
    # from the set fires unknown-call; --partial-project suppresses it.
    caller = (
        'Attribute VB_Name = "ModA"\r\nOption Explicit\r\n'
        "Public Sub S()\r\n    Call OnlyInModC\r\nEnd Sub\r\n"
    )
    other = 'Attribute VB_Name = "ModB"\r\nOption Explicit\r\nPublic Sub T()\r\nEnd Sub\r\n'
    _write(tmp_path / "ModA.bas", caller)
    _write(tmp_path / "ModB.bas", other)
    code_full = main([str(tmp_path)])
    assert "unknown-call" in capsys.readouterr().out and code_full == 1
    code_partial = main([str(tmp_path), "--partial-project"])
    assert "unknown-call" not in capsys.readouterr().out and code_partial == 0


def test_cli_inline_suppression(tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
    src = (
        'Attribute VB_Name = "Mod1"\r\nOption Explicit\r\nSub S()\r\n'
        "    Dim a(10 To 1) As Long  '@pyvba-ignore: array-declaration-impossible-bounds\r\n"
        # Read, so the unused-variable rule has nothing to say and the run's exit
        # code turns on the suppressed finding alone.
        "    Debug.Print LBound(a)\r\n"
        "End Sub\r\n"
    )
    path = _write(tmp_path / "Mod1.bas", src)
    # Honored by default: the diagnostic is suppressed, so the run is clean.
    assert main([str(path)]) == 0
    capsys.readouterr()
    # --no-inline-suppression reports it anyway (audit mode).
    code = main([str(path), "--no-inline-suppression"])
    assert code == 1
    assert "array-declaration-impossible-bounds" in capsys.readouterr().out


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

    monkeypatch.setattr(cli_mod, "read_office_project", _boom)
    code = main([str(bad)])
    assert code == 1
    assert "cannot read broken.xlsm" in capsys.readouterr().err
