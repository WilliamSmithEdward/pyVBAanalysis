"""Replay the calls recorder.py recorded through the port.

Each recorded call carries the diagnostics upstream returned. A standalone call is
replayed with the same options. A call made with project context is replayed by
rebuilding its ProjectIndex from the recorded module inputs and asking the port's
own index method for each project option, so the index and the option wiring
are compared along with the rules. Diagnostics are compared by code, span and
message, as multisets.
"""

from __future__ import annotations

import json
from collections import Counter
from collections.abc import Mapping
from pathlib import Path
from typing import Any

from recorder import STAMPED_METHODS

from pyvbaanalysis import analyze_module
from pyvbaanalysis.conditional import ConditionalCompilationEnvironment
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions, VbaDiagnostic
from pyvbaanalysis.host import get_excel_object_model, host_object_model_for_token
from pyvbaanalysis.symbols import (
    ImplicitMember,
    ModuleInput,
    ModuleSymbolKind,
    ProjectIndex,
    ProjectIndexOptions,
)

_KIND = {kind.value: kind for kind in ModuleSymbolKind}

# Upstream option keys the port reads, beyond the project context. Any other key
# is reported, so a new upstream option cannot be dropped silently.
_PLAIN_KEYS = frozenset(
    {
        "moduleName",
        "moduleKind",
        "conditionalCompilation",
        "host",
        "implicitMembers",
        "documentType",
        "referencedHosts",
        "severityOverrides",
        "designerClass",
    }
)

# Project options, by upstream name, with the AnalyzeModuleOptions field each fills.
_PROJECT_FIELDS = {
    "knownIdentifiers": "known_identifiers",
    "knownProcedures": "known_procedures",
    "knownNonTypeNames": "known_non_type_names",
    "projectProcedures": "project_procedures",
    "projectClassMembers": "project_class_members",
    "projectTypes": "project_types",
    "projectVisibleSymbols": "project_visible_symbols",
    "projectIntegerConstants": "project_integer_constants",
    "projectStringLiteralWords": "project_string_literal_words",
    "implementedInterfaces": "implemented_interfaces",
}
_SET_FIELDS = frozenset(
    {"knownIdentifiers", "knownProcedures", "knownNonTypeNames", "projectStringLiteralWords", "implementedInterfaces"}
)


def _environment(recorded: Mapping[str, Any] | None) -> ConditionalCompilationEnvironment | None:
    if recorded is None:
        return None
    return ConditionalCompilationEnvironment(
        compiler_constants=recorded.get("compilerConstants"),
        project_constants=recorded.get("projectConstants"),
    )


def _implicit_members(recorded: list[Mapping[str, str]] | None) -> list[ImplicitMember] | None:
    if recorded is None:
        return None
    return [ImplicitMember(member["name"], member["type"]) for member in recorded]


def _plain_options(opts: Mapping[str, Any], **project: Any) -> AnalyzeModuleOptions:
    return AnalyzeModuleOptions(
        module_name=opts.get("moduleName"),
        module_kind=_KIND.get(opts.get("moduleKind", "")),
        conditional_compilation=_environment(opts.get("conditionalCompilation")),
        host=opts.get("host"),
        implicit_members=_implicit_members(opts.get("implicitMembers")),
        document_type=opts.get("documentType"),
        # Upstream reads an absent list as "nothing referenced"; the port reads None
        # as unknown and stays silent, so the replay passes what upstream assumed.
        referenced_hosts=opts.get("referencedHosts", []),
        severity_overrides=opts.get("severityOverrides"),
        designer_class=opts.get("designerClass"),
        **project,
    )


def _rebuilt_index(project: Mapping[str, Any]) -> ProjectIndex:
    index_options = project.get("indexOptions") or {}
    index = ProjectIndex(
        ProjectIndexOptions(conditional_compilation=_environment(index_options.get("conditionalCompilation")))
    )
    for item in project["inputs"]:
        index.set_module(
            ModuleInput(
                item["moduleName"],
                _KIND.get(item.get("moduleKind", ""), ModuleSymbolKind.STANDARD),
                item["source"],
                conditional_compilation=_environment(item.get("conditionalCompilation")),
                implicit_members=_implicit_members(item.get("implicitMembers")),
                predeclared_id=item.get("predeclaredId"),
                designer_class=item.get("designerClass"),
            )
        )
    return index


def _project_options(row: Mapping[str, Any]) -> AnalyzeModuleOptions:
    project = row["project"]
    index = _rebuilt_index(project)
    fields: dict[str, Any] = {}
    for key, origin in project["derived"].items():
        if origin["method"] == "__literal":
            value = origin["args"][0]
            fields[_PROJECT_FIELDS[key]] = set(value) if key in _SET_FIELDS else value
        else:
            fields[_PROJECT_FIELDS[key]] = getattr(index, STAMPED_METHODS[origin["method"]])(*origin["args"])
    host_model = project.get("hostModel")
    if host_model is not None:
        if "token" in host_model:
            token = host_model["token"]
            fields["host_model"] = get_excel_object_model() if token == "excel" else host_object_model_for_token(token)
        else:
            fields["host_model"] = host_model["model"]
    return _plain_options(row["opts"], **fields)


def _shown(code: str, start: int, end: int, message: str) -> str:
    return f"{code} @{start}-{end}: {message}"


def _port_result(source: str, opts: AnalyzeModuleOptions) -> Counter[str]:
    found: list[VbaDiagnostic] = analyze_module(source, opts)
    return Counter(_shown(d.code, d.span.start, d.span.end, d.message) for d in found)


def replay(calls: Path, show: int = 10, only_code: str | None = None) -> int:
    """Replay every unique recorded call; print what differs. Returns 1 when any
    call differs or cannot be rebuilt, else 0."""
    rows = [json.loads(line) for line in calls.read_text(encoding="utf-8").splitlines() if line]
    if not rows:
        raise SystemExit(f"{calls} holds no calls; run `harness.py record` first.")
    counts = {"standalone": 0, "project": 0}
    differing: Counter[str] = Counter()
    by_code: Counter[str] = Counter()
    ignored: Counter[str] = Counter()
    failures: list[str] = []
    examples: list[str] = []
    seen: set[str] = set()
    for row in rows:
        kind = "standalone" if row["simple"] else "project"
        if kind == "project" and row.get("project") is None:
            continue  # recorded as not replayable; summarize() reports those
        key = json.dumps([row["source"], row["opts"], row.get("project")], sort_keys=True)
        if key in seen:
            continue
        seen.add(key)
        ignored.update(name for name in row["opts"] if name not in _PLAIN_KEYS)
        # A call the port cannot rebuild is reported with the rest, not raised.
        try:
            opts = _plain_options(row["opts"]) if kind == "standalone" else _project_options(row)
        except Exception as error:
            failures.append(f"{type(error).__name__}: {error}")
            continue
        counts[kind] += 1
        expected = Counter(_shown(d["code"], d["start"], d["end"], d["message"]) for d in row["result"])
        actual = _port_result(row["source"], opts)
        if expected == actual:
            continue
        only_upstream = expected - actual
        only_port = actual - expected
        if only_code and only_code not in {item.split(" ", 1)[0] for item in [*only_upstream, *only_port]}:
            continue
        differing[kind] += 1
        by_code.update(f"upstream only {item.split(' ', 1)[0]}" for item in only_upstream.elements())
        by_code.update(f"port only     {item.split(' ', 1)[0]}" for item in only_port.elements())
        if len(examples) < show:
            modules = ""
            if kind == "project":
                inputs = row["project"]["inputs"]
                modules = " project=[" + ", ".join(f"{i['moduleName']}({i.get('moduleKind')})" for i in inputs) + "]"
            lines = [f"--- {kind} call, module={row['opts'].get('moduleName')}{modules}", row["source"].rstrip()[:900]]
            lines += [f"   upstream only: {item}" for item in sorted(only_upstream.elements())]
            lines += [f"   port only:     {item}" for item in sorted(only_port.elements())]
            examples.append("\n".join(lines))

    print(
        f"replayed {counts['standalone']} unique standalone calls ({differing['standalone']} differ) "
        f"and {counts['project']} with project context ({differing['project']} differ)"
    )
    if failures:
        print(f"{len(failures)} could not be rebuilt, first: {failures[0]}")
    if ignored:
        # Tests sometimes pass keys analyzeModule itself ignores, such as the
        # completion context's meType; a new upstream option would show up here too.
        print(f"option keys the port has no field for (check upstream reads them): {dict(ignored)}")
    for name, count in by_code.most_common():
        print(f"  {count:4d}  {name}")
    if examples:
        print()
        print("\n\n".join(examples))
    return 1 if differing or failures else 0
