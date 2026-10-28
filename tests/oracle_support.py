"""Shared helpers for validating rules against the vendored oracle corpus.

Not a test module (no test_ prefix); imported by the rule test files.
"""

from __future__ import annotations

import json
import re
from functools import lru_cache
from typing import Any

from pyvbaanalysis import analyze_module_options_for
from pyvbaanalysis.conditional import ConditionalCompilationEnvironment
from pyvbaanalysis.diagnostics import analyze_module
from pyvbaanalysis.evidence import DATA_DIR, OracleCase, load_audit, load_oracle_cases
from pyvbaanalysis.host.host_libraries import host_token_for_libid
from pyvbaanalysis.symbols import ImplicitMember, ModuleInput, ModuleSymbolKind, ProjectIndex

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

# The names the cases' references go by, as the dir stream records them (measured;
# upstream unreferencedQualifiers1413.test.ts).
LIBRARY_NAMES = {
    "{420B2830-E718-11CF-893D-00A0C9054228}": "Scripting",
    "{3F4DACA7-160D-11D2-A8E9-00104B365C9F}": "VBScript_RegExp_55",
    "{B691E011-1797-432E-907A-4D8C69339129}": "ADODB",
    "{F5078F18-C551-11D3-89B9-0000F81FE221}": "MSXML2",
    "{0002E157-0000-0000-C000-000000000046}": "VBIDE",
    "{662901FC-6951-4854-9EB2-D9A2570F2B2E}": "WinHttp",
    "{50A7E9B0-70EF-11D1-B75A-00A0C90564FE}": "Shell32",
    "{00062FFF-0000-0000-C000-000000000046}": "Outlook",
    "{00020813-0000-0000-C000-000000000046}": "Excel",
}
# What a new file of each host records before any reference is added, and its
# project's name (measured).
DEFAULT_LIBRARIES = {
    "excel": ["stdole", "Office"],
    "word": ["stdole", "Normal", "Office"],
    "powerpoint": ["stdole", "Office"],
    "access": ["stdole", "DAO"],
}
PROJECT_NAMES = {"excel": "VBAProject", "word": "Project", "powerpoint": "VBAProject", "access": "Database1"}

# Each case's references by library name: the oracle's fresh file's defaults plus
# what the case adds, as the extension reads them from the project.
_REFERENCED_LIBRARIES: dict[str, list[str]] = {
    raw["id"]: [
        *DEFAULT_LIBRARIES[raw.get("host") or "excel"],
        *(LIBRARY_NAMES[ref["guid"]] for ref in raw.get("references", [])),
    ]
    for raw in json.loads((DATA_DIR / "vbe_oracle_cases.json").read_text(encoding="utf-8"))["cases"]
}

_PROG_ID_RE = re.compile(r"^Forms\.(\w+)\.\d+$")


def oracle_module_input(module: dict[str, Any]) -> ModuleInput:
    """An oracle fixture's module as the analyzer's project input, with what its
    designer and attributes declare: a UserForm's controls, typed by ProgID, a
    document module's designer class, a class's PredeclaredId (upstream's
    tests/helpers/oracleModuleInput.ts, XLIDE issue #1419). The VBE oracle builds
    the same module from the same fields, so the analyzer sees the project the
    VBE compiled."""
    controls = module.get("designerControls")
    implicit: list[ImplicitMember] | None
    if controls is not None:
        implicit = []
        for control in controls:
            match = _PROG_ID_RE.match(control["progId"])
            if match is None:
                raise ValueError(f"No MSForms type for the ProgID {control['progId']}")
            implicit.append(ImplicitMember(control["name"], f"MSForms.{match.group(1)}"))
    elif module.get("implicitMembers") is not None:
        implicit = [ImplicitMember(member["name"], member["type"]) for member in module["implicitMembers"]]
    else:
        implicit = None
    code_name = module.get("codeName")
    return ModuleInput(
        code_name if code_name is not None else module["name"],
        _kind(module.get("type", "standard")),
        module["source"],
        implicit_members=implicit,
        predeclared_id=module.get("predeclaredId"),
        designer_class=module.get("designerClass"),
        # A renamed module is still the object it was made as, which the
        # extension reads from the module's VB_Base.
        document_type=_document_type_of(module["name"]) if code_name is not None else None,
    )


def _document_type_of(name: str) -> str:
    """What a new file's document module is, by the name it was made with."""
    lower = name.lower()
    return "workbook" if lower == "thisworkbook" else "document" if lower == "thisdocument" else "worksheet"


_RAW_MODULES: dict[str, list[dict[str, Any]]] = {
    raw["id"]: raw["modules"]
    for raw in json.loads((DATA_DIR / "vbe_oracle_cases.json").read_text(encoding="utf-8"))["cases"]
    if "modules" in raw
}

# The host each case was measured in; a case that names none ran in Excel.
_HOSTS: dict[str, str] = {
    raw["id"]: raw["host"]
    for raw in json.loads((DATA_DIR / "vbe_oracle_cases.json").read_text(encoding="utf-8"))["cases"]
    if raw.get("host")
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
    raw_modules = {raw["name"]: raw for raw in _RAW_MODULES.get(case.id, [])}
    inputs = [
        oracle_module_input(raw_modules[module.name])
        if module.name in raw_modules
        else ModuleInput(module.name, _kind(module.module_type), module.source)
        for module in case.modules
    ]
    for module_input in inputs:
        index.set_module(module_input)
    out: set[str] = set()
    for module_input in inputs:
        opts = analyze_module_options_for(
            index, module_input.module_name, module_input.module_kind,
            host=_HOSTS.get(case.id, "excel"),
            referenced_hosts=_REFERENCED_HOSTS.get(case.id, []),
            # A rejected case is judged as the extension runs it, with the
            # project's references and name known (upstream's dedicated tests,
            # e.g. unreferencedQualifiers1413.test.ts). An accepted case is swept
            # as upstream's oracleAcceptedNoFalsePositives.test.ts sweeps it,
            # with the reference list unknown: a single-module case omits the
            # fresh file's document modules, which a known list would then miss.
            **(
                {
                    "referenced_libraries": _REFERENCED_LIBRARIES[case.id],
                    "project_name": PROJECT_NAMES[_HOSTS.get(case.id, "excel")],
                }
                if case.expected == "rejected"
                else {}
            ),
            conditional_compilation=ConditionalCompilationEnvironment(project_constants={}),
        )
        for diag in analyze_module(module_input.source, opts):
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


# Audit entries the upstream analyzer is silent on as well (harness.py corpus at
# XLIDE 11.2.1 reports 0 differing modules), keyed by (case, code), with the
# upstream test that keeps each one quiet.
UPSTREAM_SILENT: dict[tuple[str, str], str] = {
    ("me_assign_document_module_runtime", "invalid-assignment-target"):
        "upstream excelMeasuredFalsePositives.test.ts keeps it quiet",
    ("project_time_variant_function_index_runtime", "argument-count"):
        "upstream excelMeasuredFalsePositives.test.ts keeps it quiet",
    ("if_single_line_else_after_colon_duplicate_in_block_if_compile", "else-branch-order"):
        "upstream's 53e6868d test expects only else-without-if here",
    ("document_type_thisworkbook_range_runtime", "assignment-object-type-mismatch"):
        "no upstream test asserts it, and upstream does not report it",
    ("document_type_sheet_range_runtime", "assignment-object-type-mismatch"):
        "no upstream test asserts it, and upstream does not report it",
    ("reftype_access_office_commandbar_compile", "invalid-as-type-name"):
        "upstream hostReferenceTypes.test.ts KNOWN_SILENT: Office types are in the Access model",
    **{
        (case_id, "invalid-as-type-name"):
            "upstream hostReferenceTypes.test.ts: Word's Normal template could hold the name"
        for case_id in (
            "host_word_excel_worksheet_type_compile",
            "reftype_word_unknown_compile",
            "outlook_word_noref_dim_bare_mailitem_compile",
        )
    },
    **{
        (f"issue1411_{host}b_other_long_local_compile", "argument-count"):
            "upstream receiverShadowing1411.test.ts: a member of a Long, not a call"
        for host in ("", "word_", "powerpoint_", "access_")
    },
    **{
        (f"issue1413_{host}noref_shell_compile", "invalid-as-type-name"):
            "upstream unreferencedQualifiers1413.test.ts reports it as missing-library-reference only"
        for host in ("", "word_", "powerpoint_", "access_")
    },
}


def assert_oracle_behavior(code: str, skip_ids: frozenset[str] = frozenset()) -> int:
    """Validate a code against its asserted oracle cases.

    assertedOracleCases mixes positive cases (rejected -> the code must fire) and
    control cases (accepted -> the code must NOT fire, when the case constrains it).
    `skip_ids` names cases that require infrastructure a milestone defers (e.g. the
    host member-completion surface); they are not asserted. Returns cases checked.
    """
    checked = 0
    for case in asserted_cases(code):
        if case.id in skip_ids or (case.id, code) in UPSTREAM_SILENT:
            continue
        emitted = case_codes(case)
        if case.expected == "rejected":
            assert code in emitted, f"{case.id}: expected {code} to fire, got {sorted(emitted)}"
        elif case.expected == "accepted" and accepted_case_constrains(code, case):
            assert code not in emitted, f"{case.id}: {code} must not fire on an accepted control"
        checked += 1
    return checked
