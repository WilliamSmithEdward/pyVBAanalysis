"""Run the same inputs through both analyzers and compare what they report.

Two probes. Each has an upstream half (upstream/*.mjs, run by tsx against a pin
through tools/xlide_source.mjs) and a port half here, and both write one JSON
line per finding, so the comparison reads data rather than parsing text:

* corpus: every case of the oracle corpus, each module standalone ("S") and with
  the case's modules as a project ("P"), one line per diagnostic.
* projects: labelled multi-module cases (cases.py) with the case's host and
  reference list, one line per module holding all its findings.

Findings are compared as multisets per module, by code, span and message.
"""

from __future__ import annotations

import json
import os
import subprocess
from collections import Counter
from collections.abc import Iterable, Mapping
from pathlib import Path
from typing import Any

from pyvbaanalysis import analyze_module, analyze_module_options_for, analyze_project
from pyvbaanalysis.diagnostics import AnalyzeModuleOptions, VbaDiagnostic
from pyvbaanalysis.symbols import ImplicitMember, ModuleInput, ModuleSymbolKind, ProjectIndex

_UPSTREAM = Path(__file__).resolve().parent / "upstream"

# Upstream's moduleKindFromType (src/vbaProjectAnalysis.ts): a VB6 UserControl,
# PropertyPage or Designer and an Access form or report are object modules with a
# designer, like a UserForm, and a type it does not name is a standard module.
_KIND_FROM_TYPE = {
    "class": ModuleSymbolKind.CLASS,
    "document": ModuleSymbolKind.DOCUMENT,
    "userform": ModuleSymbolKind.USERFORM,
    "usercontrol": ModuleSymbolKind.USERFORM,
    "propertypage": ModuleSymbolKind.USERFORM,
    "designer": ModuleSymbolKind.USERFORM,
    "accessform": ModuleSymbolKind.USERFORM,
    "accessreport": ModuleSymbolKind.USERFORM,
}


def kind_from_type(module_type: str | None) -> ModuleSymbolKind:
    return _KIND_FROM_TYPE.get(module_type or "", ModuleSymbolKind.STANDARD)


def _finding(d: VbaDiagnostic) -> str:
    return f"{d.code} @{d.span.start}-{d.span.end}: {d.message}"


# -- the port's halves ---------------------------------------------------------


def _corpus_modules(case: Mapping[str, Any]) -> list[tuple[str, ModuleSymbolKind, str]]:
    # As corpus_messages.mjs reads them: a single-module case carries its source.
    if case.get("modules") is not None:
        raw = case["modules"]
    else:
        name = case.get("moduleName") if case.get("moduleName") is not None else "Module1"
        raw = [{"name": name, "type": "standard", "source": case["source"]}]
    return [(m["name"], kind_from_type(m.get("type", "standard")), m["source"]) for m in raw]


def port_corpus(corpus: Path) -> list[dict[str, Any]]:
    """The port over every oracle case, standalone and as a project. The reference
    list is given as known and empty, which is what upstream assumes of an absent one."""
    records: list[dict[str, Any]] = []

    def emit(pass_: str, case_id: str, module: str, found: Iterable[VbaDiagnostic]) -> None:
        records.extend(
            {
                "pass": pass_,
                "id": case_id,
                "module": module,
                "code": d.code,
                "start": d.span.start,
                "end": d.span.end,
                "message": d.message,
            }
            for d in found
        )

    for case in json.loads(corpus.read_text(encoding="utf-8"))["cases"]:
        modules = _corpus_modules(case)
        for name, kind, source in modules:
            opts = AnalyzeModuleOptions(module_name=name, module_kind=kind, referenced_hosts=[])
            emit("S", case["id"], name, analyze_module(source, opts))
        index = ProjectIndex()
        for name, kind, source in modules:
            index.set_module(ModuleInput(name, kind, source))
        for name, kind, source in modules:
            opts = analyze_module_options_for(index, name, kind, referenced_hosts=[])
            emit("P", case["id"], name, analyze_module(source, opts))
    return records


def port_projects(cases: Path) -> list[dict[str, Any]]:
    """The port over labelled project cases, with each case's host and references."""
    records: list[dict[str, Any]] = []
    for case in json.loads(cases.read_text(encoding="utf-8")):
        modules = [
            ModuleInput(
                m["name"],
                kind_from_type(m.get("type", "standard")),
                m["source"],
                implicit_members=[ImplicitMember(i["name"], i["type"]) for i in m["implicitMembers"]]
                if m.get("implicitMembers")
                else None,
                predeclared_id=m.get("predeclaredId"),
                designer_class=m.get("designerClass") or None,
            )
            for m in case["modules"]
        ]
        results = analyze_project(modules, host=case.get("host") or None, referenced_hosts=case.get("referenced") or [])
        records.extend(
            {"label": case["label"], "module": module.module_name, "findings": [_finding(d) for d in results[module.module_name]]}
            for module in modules
        )
    return records


# -- the upstream halves -------------------------------------------------------


def upstream(script: str, pin: Path, argument: Path, timeout: int) -> list[dict[str, Any]]:
    """Run one upstream probe against ``pin`` and read its JSON lines."""
    result = subprocess.run(
        ["npx", "-y", "tsx", str(_UPSTREAM / script), str(argument)],
        env=dict(os.environ, XLIDE_ROOT=str(pin)),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        # npx is a .cmd shim on Windows, which only a shell runs.
        shell=os.name == "nt",
        timeout=timeout,
    )
    if result.returncode != 0:
        raise SystemExit(f"{script} failed ({result.returncode}): {result.stderr.strip()[-2000:]}")
    return [json.loads(line) for line in result.stdout.splitlines() if line.strip()]


# -- comparing -----------------------------------------------------------------


def write_jsonl(path: Path, records: Iterable[Mapping[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("".join(json.dumps(r, ensure_ascii=True) + "\n" for r in records), encoding="utf-8")


def _grouped_corpus(records: Iterable[Mapping[str, Any]]) -> dict[str, Counter[str]]:
    groups: dict[str, Counter[str]] = {}
    for r in records:
        key = f"{r['pass']} {r['id']} :: {r['module']}"
        groups.setdefault(key, Counter())[f"{r['code']} @{r['start']}-{r['end']}: {r['message']}"] += 1
    return groups


def _grouped_projects(records: Iterable[Mapping[str, Any]]) -> dict[str, Counter[str]]:
    return {f"{r['label']} :: {r['module']}": Counter(r["findings"]) for r in records}


def compare(
    upstream_records: list[dict[str, Any]], port_records: list[dict[str, Any]], kind: str, show: int
) -> int:
    """Print the modules whose findings differ, as findings only one side has.
    Returns 1 when any differ, else 0."""
    group = _grouped_corpus if kind == "corpus" else _grouped_projects
    upstream_groups = group(upstream_records)
    port_groups = group(port_records)
    total = sum(sum(c.values()) for c in upstream_groups.values())
    differing = [
        key
        for key in sorted(set(upstream_groups) | set(port_groups))
        if upstream_groups.get(key, Counter()) != port_groups.get(key, Counter())
    ]
    unit = "diagnostics" if kind == "corpus" else "findings"
    print(f"{kind}: {total} upstream {unit} over {len(upstream_groups)} modules; {len(differing)} modules differ")
    for key in differing[:show]:
        up = upstream_groups.get(key, Counter())
        py = port_groups.get(key, Counter())
        print(f"== {key}")
        for item in sorted((up - py).elements()):
            print(f"   upstream only: {item}")
        for item in sorted((py - up).elements()):
            print(f"   port only:     {item}")
    if len(differing) > show:
        print(f"... and {len(differing) - show} more")
    return 1 if differing else 0
