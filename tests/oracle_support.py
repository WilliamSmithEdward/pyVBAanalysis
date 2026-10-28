"""Shared helpers for validating rules against the vendored oracle corpus.

Not a test module (no test_ prefix); imported by the rule test files.
"""

from __future__ import annotations

import json
from functools import lru_cache

from pyvbaanalysis import analyze_module_options_for
from pyvbaanalysis.conditional import ConditionalCompilationEnvironment
from pyvbaanalysis.diagnostics import analyze_module
from pyvbaanalysis.evidence import DATA_DIR, OracleCase, load_audit, load_oracle_cases
from pyvbaanalysis.host.host_libraries import host_token_for_libid
from pyvbaanalysis.symbols import ModuleInput, ModuleSymbolKind, ProjectIndex

AUDIT = {a.code: a for a in load_audit()}
CASES = {c.id: c for c in load_oracle_cases()}

# The libraries a case's `references` add before its modules (Outlook, by GUID), as
# the host tokens the analysis takes: a case the oracle compiled with Outlook
# referenced is analyzed against Outlook's model too (upstream oracleReferencedHosts).
_REFERENCED_HOSTS: dict[str, list[str]] = {
    raw["id"]: [token for ref in raw["references"] if (token := host_token_for_libid(ref["guid"])) is not None]
    for raw in json.loads((DATA_DIR / "vbe_oracle_cases.json").read_text(encoding="utf-8"))["cases"]
    if raw.get("references")
}

_KIND = {
    "standard": ModuleSymbolKind.STANDARD,
    "class": ModuleSymbolKind.CLASS,
    "document": ModuleSymbolKind.DOCUMENT,
    "userform": ModuleSymbolKind.USERFORM,
}


def _kind(module_type: str) -> ModuleSymbolKind:
    return _KIND.get(module_type, ModuleSymbolKind.STANDARD)


@lru_cache(maxsize=len(CASES))
def case_codes(case: OracleCase) -> frozenset[str]:
    """Union of diagnostic codes analyze_module emits across a case's modules.

    Each module is analyzed with the cross-module project context the real
    analyzer sees (mirrors how XLIDE runs a project): a ProjectIndex is built
    from every module in the case, and each module's options come from the same
    builder analyze_project uses. For a single-module case this context is
    empty, so behavior is unchanged.

    Every case was run in a default Excel workbook, so the host is named as Excel,
    and the reference list is known to name no other application's library rather
    than unknown, except the libraries a case's `references` add. The fresh projects also have known-empty Conditional Compilation
    Arguments, so an undefined #If name is provably zero rather than unknown.

    Cases are immutable. Reuse their full findings across rule-specific assertions
    so enlarging the corpus does not repeat the same complete analysis per rule.
    """
    index = ProjectIndex()
    for module in case.modules:
        index.set_module(ModuleInput(module.name, _kind(module.module_type), module.source))
    out: set[str] = set()
    for module in case.modules:
        opts = analyze_module_options_for(
            index, module.name, _kind(module.module_type), host="excel",
            referenced_hosts=_REFERENCED_HOSTS.get(case.id, []),
            conditional_compilation=ConditionalCompilationEnvironment(project_constants={}),
        )
        for diag in analyze_module(module.source, opts):
            out.add(diag.code)
    return frozenset(out)


def asserted_cases(code: str) -> list[OracleCase]:
    """Corpus cases the audit asserts the given code fires on."""
    return [CASES[i] for i in AUDIT[code].asserted_oracle_cases if i in CASES]


def accepted_cases() -> list[OracleCase]:
    """Every compile-accepted (valid) corpus case - used for no-false-positive sweeps."""
    return [c for c in CASES.values() if c.expected == "accepted"]


# Diagnostic kinds whose meaning is a deterministic *runtime* error rather than a
# compile error. Such a diagnostic firing on code that was only verified to
# *compile* (evidence_phase == "compile") is not a false positive: the case never
# asserted the code runs cleanly. It is only a false positive on a case that was
# verified at runtime (evidence_phase == "runtime") and accepted.
_RUNTIME_DIAGNOSTIC_KINDS = frozenset({"deterministic-runtime-error", "runtime-risk"})


def _is_runtime_diagnostic(code: str) -> bool:
    entry = AUDIT.get(code)
    return entry is not None and entry.diagnostic_kind in _RUNTIME_DIAGNOSTIC_KINDS


def accepted_case_constrains(code: str, case: OracleCase) -> bool:
    """Whether an accepted case asserts the given code must NOT fire on it.

    Compile-error diagnostics are constrained by every accepted (compile-valid)
    case. Runtime-error diagnostics are only constrained by runtime-verified
    accepted cases - a runtime diagnostic on compile-only-verified code flags a
    real runtime fault and is not a false positive.
    """
    if _is_runtime_diagnostic(code):
        return case.evidence_phase == "runtime"
    return True


def oracle_false_positives(case: OracleCase, codes: tuple[str, ...]) -> set[str]:
    """The subset of `codes` emitted on an accepted case that are true false positives."""
    return {c for c in case_codes(case) & set(codes) if accepted_case_constrains(c, case)}


def assert_oracle_behavior(code: str, skip_ids: frozenset[str] = frozenset()) -> int:
    """Validate a code against its asserted oracle cases.

    assertedOracleCases mixes positive cases (rejected -> the code must fire) and
    control cases (accepted -> the code must NOT fire, when the case constrains it).
    `skip_ids` names cases that require infrastructure a milestone defers (e.g. the
    host member-completion surface); they are not asserted. Returns cases checked.
    """
    checked = 0
    for case in asserted_cases(code):
        if case.id in skip_ids:
            continue
        emitted = case_codes(case)
        if case.expected == "rejected":
            assert code in emitted, f"{case.id}: expected {code} to fire, got {sorted(emitted)}"
        elif case.expected == "accepted" and accepted_case_constrains(code, case):
            assert code not in emitted, f"{case.id}: {code} must not fire on an accepted control"
        checked += 1
    return checked
