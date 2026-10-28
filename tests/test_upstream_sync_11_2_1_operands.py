"""Rule contracts from the XLIDE 11.1.0 -> 11.2.1 sync (operands, bounds, Is).

Each section names the upstream commit whose own tests it follows. The oracle
cases are read from the vendored vbe_oracle_cases.json and analyzed as the
extension analyzes the case's host.
"""

from __future__ import annotations

import json
from typing import Any

import pytest

from pyvbaanalysis import analyze_module, analyze_module_options_for
from pyvbaanalysis.conditional import ConditionalCompilationEnvironment
from pyvbaanalysis.diagnostics.model import DiagnosticSeverity
from pyvbaanalysis.evidence import DATA_DIR
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind, ProjectIndex

_ORACLE: dict[str, dict[str, Any]] = {
    case["id"]: case
    for case in json.loads((DATA_DIR / "vbe_oracle_cases.json").read_text(encoding="utf-8"))["cases"]
}

_KIND = {
    "standard": ModuleSymbolKind.STANDARD,
    "class": ModuleSymbolKind.CLASS,
    "document": ModuleSymbolKind.DOCUMENT,
    "userform": ModuleSymbolKind.USERFORM,
}


def _case(case_id: str) -> dict[str, Any]:
    return _ORACLE[case_id]


def _modules(case: dict[str, Any]) -> list[dict[str, Any]]:
    if "modules" in case:
        return list(case["modules"])
    return [{"name": "Module1", "type": "standard", "source": case["source"]}]


def _errors_for(case: dict[str, Any]) -> list[str]:
    """Each error as `code: spanned text`, analyzed in the case's host."""
    host = case.get("host", "excel")
    modules = _modules(case)
    index = ProjectIndex()
    for module in modules:
        index.set_module(ModuleInput(module["name"], _KIND.get(module.get("type", "standard"), ModuleSymbolKind.STANDARD), module["source"]))
    out: list[str] = []
    for module in modules:
        kind = _KIND.get(module.get("type", "standard"), ModuleSymbolKind.STANDARD)
        opts = analyze_module_options_for(
            index, module["name"], kind, host=host,
            conditional_compilation=ConditionalCompilationEnvironment(project_constants={}),
        )
        source = module["source"]
        out += [
            f"{d.code}: {source[d.span.start:d.span.end]}"
            for d in analyze_module(source, opts)
            if d.severity is DiagnosticSeverity.ERROR
        ]
    return out


def _codes_of(errors: list[str]) -> list[str]:
    return [finding.split(":")[0] for finding in errors]


_HOSTS = ("", "word_", "powerpoint_", "access_")


def _each_host(name: str, phase: str) -> list[str]:
    return [f"issue1412_{host}{name}_{phase}" for host in _HOSTS]


# -- 43c8c17b: ReDim element types by identity, bounds a Long cannot hold ----

_CHANGES = [
    *(
        case_id
        for name in ("r_project_collection_to_vba", "r_vba_collection_to_project", "r_long_to_integer")
        for case_id in _each_host(name, "compile")
    ),
    "issue1412_r_project_range_to_excel_compile",
    "issue1412_word_r_project_range_to_word_compile",
]


@pytest.mark.parametrize("case_id", _CHANGES)
def test_redim_changing_the_named_type_is_reported(case_id: str) -> None:
    case = _case(case_id)
    assert case["expected"] == "rejected"
    assert len([f for f in _errors_for(case) if f.startswith("redim-type-change:")]) == 1


_SAME = [
    *(
        case_id
        for name in ("r_project_collection_same", "r_collection_to_vba_no_class")
        for case_id in _each_host(name, "compile")
    ),
    "issue1412_r_range_to_excel_no_class_compile",
    "issue1412_word_r_range_to_word_no_class_compile",
]


@pytest.mark.parametrize("case_id", _SAME)
def test_redim_naming_the_same_type_is_quiet(case_id: str) -> None:
    case = _case(case_id)
    assert case["expected"] == "accepted"
    assert _errors_for(case) == []


_MALFORMED_LITERALS = [
    "l_dim_explicit_lower_long_overflow", "l_dim_upper_long_overflow", "l_redim_long_overflow",
    "l_redim_explicit_lower_long_overflow", "l_dim_integer_suffix_overflow",
]


@pytest.mark.parametrize("case_id", [c for name in _MALFORMED_LITERALS for c in _each_host(name, "compile")])
def test_suffixed_bound_out_of_range_is_only_the_literal(case_id: str) -> None:
    case = _case(case_id)
    assert case["expected"] == "rejected"
    assert _codes_of(_errors_for(case)) == ["suffixed-literal-overflow"]


_OVERFLOWS = [
    "l_redim_double_overflow", "l_redim_unsuffixed_overflow",
    "l_redim_explicit_upper_double_overflow", "l_redim_negative_lower_double_overflow",
]


@pytest.mark.parametrize("case_id", [c for name in _OVERFLOWS for c in _each_host(name, "runtime")])
def test_redim_bound_a_long_cannot_hold_overflows(case_id: str) -> None:
    case = _case(case_id)
    assert case["expected"] == "rejected"
    assert _codes_of(_errors_for(case)) == ["arithmetic-overflow"]


@pytest.mark.parametrize(
    "case_id",
    [
        *_each_host("l_dim_long_max", "compile"),
        *(c for name in ("l_redim_double_in_range", "b_element_bound", "b_long_bound") for c in _each_host(name, "runtime")),
    ],
)
def test_bounds_a_long_holds_are_quiet(case_id: str) -> None:
    case = _case(case_id)
    assert case["expected"] == "accepted"
    assert _errors_for(case) == []
