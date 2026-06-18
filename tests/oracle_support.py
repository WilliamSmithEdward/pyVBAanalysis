"""Shared helpers for validating rules against the vendored oracle corpus.

Not a test module (no test_ prefix); imported by the rule test files.
"""

from __future__ import annotations

from pyvbaanalysis.diagnostics import AnalyzeModuleOptions, analyze_module
from pyvbaanalysis.evidence import OracleCase, load_audit, load_oracle_cases
from pyvbaanalysis.symbols import ModuleSymbolKind

AUDIT = {a.code: a for a in load_audit()}
CASES = {c.id: c for c in load_oracle_cases()}

_KIND = {
    "standard": ModuleSymbolKind.STANDARD,
    "class": ModuleSymbolKind.CLASS,
    "document": ModuleSymbolKind.DOCUMENT,
    "userform": ModuleSymbolKind.USERFORM,
}


def case_codes(case: OracleCase) -> set[str]:
    """Union of diagnostic codes analyze_module emits across a case's modules."""
    out: set[str] = set()
    for module in case.modules:
        opts = AnalyzeModuleOptions(
            module_name=module.name, module_kind=_KIND.get(module.module_type, ModuleSymbolKind.STANDARD)
        )
        for diag in analyze_module(module.source, opts):
            out.add(diag.code)
    return out


def asserted_cases(code: str) -> list[OracleCase]:
    """Corpus cases the audit asserts the given code fires on."""
    return [CASES[i] for i in AUDIT[code].asserted_oracle_cases if i in CASES]


def accepted_cases() -> list[OracleCase]:
    """Every compile-accepted (valid) corpus case - used for no-false-positive sweeps."""
    return [c for c in CASES.values() if c.expected == "accepted"]
