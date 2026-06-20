"""The public package-root surface: re-exports and the small public helpers."""

from __future__ import annotations

import pytest

import pyvbaanalysis
from pyvbaanalysis import line_col, validate_severity_overrides

_ROOT_EXPORTS = [
    "analyze_module",
    "analyze_project",
    "analyze_workbook",
    "analyze_loose_file",
    "analyze_loose_files",
    "ConditionalCompilationEnvironment",
    "DEFAULT_COMPILER_CONSTANTS",
    "ProjectIndexOptions",
    "DiagnosticSeverity",
    "line_col",
    "rule_metadata_by_code",
    "validate_severity_overrides",
]


def test_root_exports_present() -> None:
    for name in _ROOT_EXPORTS:
        assert hasattr(pyvbaanalysis, name), name
        assert name in pyvbaanalysis.__all__, name


def test_line_col_basics() -> None:
    source = "ab\ncde\nf"
    assert line_col(source, 0) == (1, 1)
    assert line_col(source, 1) == (1, 2)
    assert line_col(source, 3) == (2, 1)  # first char of line 2
    assert line_col(source, 5) == (2, 3)
    assert line_col(source, 7) == (3, 1)
    # An offset past the end clamps to the end.
    assert line_col(source, 999) == (3, 2)


def test_validate_severity_overrides_accepts_valid() -> None:
    # option-explicit-missing allows "off"; this must not raise.
    validate_severity_overrides({"option-explicit-missing": "off"})
    validate_severity_overrides(None)
    validate_severity_overrides({})


def test_validate_severity_overrides_rejects_unknown_code() -> None:
    with pytest.raises(ValueError, match="unknown diagnostic code"):
        validate_severity_overrides({"not-a-real-code": "off"})


def test_validate_severity_overrides_rejects_disallowed_value() -> None:
    # assignment-type-mismatch may only be downgraded to warning, not turned off.
    with pytest.raises(ValueError, match="not allowed"):
        validate_severity_overrides({"assignment-type-mismatch": "off"})


def test_validate_severity_overrides_rejects_bad_value() -> None:
    with pytest.raises(ValueError, match="invalid severity value"):
        validate_severity_overrides({"option-explicit-missing": "banana"})
