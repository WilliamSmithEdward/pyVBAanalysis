"""Rule contracts from the XLIDE 11.2.1 -> 11.2.11 sync (declarations and
statement shapes).

Each section names the upstream commit whose own tests it follows. The measured
cases are the oracle fixtures of the vendored corpus, analyzed as upstream's tests
analyze them: every module of the case in one project, in the case's host.
"""

from __future__ import annotations

import pytest

from pyvbaanalysis import analyze_module, analyze_module_options_for
from pyvbaanalysis.conditional import ConditionalCompilationEnvironment
from pyvbaanalysis.diagnostics.model import DiagnosticSeverity
from pyvbaanalysis.evidence import OracleCase
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind, ProjectIndex
from oracle_support import _RAW_MODULES, _HOSTS, CASES, oracle_module_input  # type: ignore[attr-defined]


def _measured(prefix: str) -> list[OracleCase]:
    return [case for case in CASES.values() if case.id.startswith(prefix)]


def _project_errors(inputs: list[ModuleInput], host: str | None = None) -> list[str]:
    """Upstream's errors(): every error of every module, as `Module: code: message`."""
    index = ProjectIndex()
    for module_input in inputs:
        index.set_module(module_input)
    out: list[str] = []
    for module_input in inputs:
        opts = analyze_module_options_for(
            index,
            module_input.module_name,
            module_input.module_kind,
            host=host,
            conditional_compilation=ConditionalCompilationEnvironment(project_constants={}),
        )
        out.extend(
            f"{module_input.module_name}: {diag.code}: {diag.message}"
            for diag in analyze_module(module_input.source, opts)
            if diag.severity == DiagnosticSeverity.ERROR
        )
    return out


def _case_errors(case: OracleCase) -> list[str]:
    inputs = [oracle_module_input(raw) for raw in _RAW_MODULES[case.id]]
    return _project_errors(inputs, _HOSTS.get(case.id))


def _module_errors(source: str) -> list[str]:
    return _project_errors([ModuleInput("Module1", ModuleSymbolKind.STANDARD, source)])


# -- 5cad8c3e: event sources and standard modules as types -------------------

_ISSUE_1530 = _measured("issue1530_")


def test_issue_1530_has_every_measured_case() -> None:
    assert len(_ISSUE_1530) == 108


@pytest.mark.parametrize("case", _ISSUE_1530, ids=lambda case: case.id)
def test_issue_1530_agrees_with_the_vbe(case: OracleCase) -> None:
    found = _case_errors(case)
    if case.expected == "accepted":
        assert found == []
    else:
        # One report, with the VBE's own reason.
        assert len(found) == 1
        reason = "A module is not a valid type" if "_module_type_" in case.id else "does not source automation events"
        assert reason in found[0]


def test_names_the_module_a_with_events_variable_or_a_dim_is_declared_as() -> None:
    found = _project_errors([
        ModuleInput("Helpers", ModuleSymbolKind.STANDARD, "Option Explicit\r\nSub P()\r\nEnd Sub\r\n"),
        ModuleInput("Holder", ModuleSymbolKind.CLASS, "Option Explicit\r\nPrivate WithEvents h As Helpers\r\nPrivate m As Helpers\r\n"),
    ])
    assert len(found) == 2
    assert "WithEvents variable 'h' is declared As Helpers, and Helpers is a standard module of this project." in found[0]
    assert "'Helpers' is a standard module of this project, and a module is not a type." in found[1]


def test_a_userform_whose_code_declares_an_event_is_a_valid_source() -> None:
    found = _project_errors([
        ModuleInput("Picker", ModuleSymbolKind.USERFORM, "Option Explicit\r\nPublic Event Picked(ByVal index As Long)\r\n"),
        ModuleInput(
            "Holder",
            ModuleSymbolKind.CLASS,
            "Option Explicit\r\nPrivate WithEvents p As Picker\r\nPrivate Sub p_Picked(ByVal index As Long)\r\n"
            "    Debug.Print index\r\nEnd Sub\r\n",
        ),
    ])
    assert found == []
