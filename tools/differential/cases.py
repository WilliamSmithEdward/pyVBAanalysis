"""Build labelled project cases for `harness.py projects` from files on disk.

A case is {label, host?, referenced, modules: [{name, type, source}]}, the shape
both upstream/project_probe.mjs and probes.port_projects read. Both analyzers get
exactly what the port's readers produced, so a difference between them is a
difference in analysis, not in reading.

* An Office file (workbook, document, presentation or database) becomes one case,
  with the host its extension implies and the libraries its reference list names.
  A list that could not be read is given as empty to both sides.
* A folder holding exported .bas/.cls/.frm files becomes one case.
* With a pin, every fenced VBA block in upstream's syntax_corpus markdown
  becomes a one-module case, a class when it opens with a class header.

Exported modules and markdown blocks name no host, as the CLI names none for a
loose file, unless one is asked for. An Office file always names its own.
"""

from __future__ import annotations

import re
import sys
from collections import defaultdict
from pathlib import Path
from typing import Any

from pyvbaanalysis.host import host_token_for_file_name
from pyvbaanalysis.reader import load_loose_module, read_office_project

_LOOSE_SUFFIXES = frozenset({".bas", ".cls", ".frm"})
_SKIPPED_FOLDERS = frozenset({".git", ".venv", "node_modules", "artifacts", "dist", "build", "__pycache__"})
_FENCE_OPEN = re.compile(r"^```(vba|vb)\b", re.IGNORECASE)
_CLASS_HEADER = re.compile(r"^(VERSION 1\.0 CLASS|Attribute VB_Base)", re.MULTILINE)


def _office_case(path: Path, label: str) -> dict[str, Any] | None:
    try:
        project = read_office_project(path)
    except Exception as error:
        # A container with no readable VBA project has nothing to compare.
        print(f"skip {label}: {error}", file=sys.stderr)
        return None
    return {
        "label": label,
        "host": project.host,
        "referenced": project.referenced_hosts or [],
        "modules": [{"name": m.name, "type": m.kind.value, "source": m.source} for m in project.modules],
    }


def _loose_case(paths: list[Path], label: str, host: str | None) -> dict[str, Any] | None:
    modules: list[dict[str, str]] = []
    seen: set[str] = set()
    for path in sorted(paths):
        try:
            loaded = load_loose_module(path)
        except Exception as error:
            print(f"skip {path}: {error}", file=sys.stderr)
            continue
        # A project holds one module per name; a second export of it is a copy.
        if loaded.name.lower() in seen:
            continue
        seen.add(loaded.name.lower())
        modules.append({"name": loaded.name, "type": loaded.kind.value, "source": loaded.source})
    if not modules:
        return None
    case: dict[str, Any] = {"label": label, "referenced": [], "modules": modules}
    if host:
        case["host"] = host
    return case


def _folder_cases(root: Path, host: str | None) -> list[dict[str, Any]]:
    cases: list[dict[str, Any]] = []
    loose: dict[Path, list[Path]] = defaultdict(list)
    for path in sorted(root.rglob("*")):
        if not path.is_file() or _SKIPPED_FOLDERS & set(path.relative_to(root).parts):
            continue
        label = f"{root.name}/{path.relative_to(root).as_posix()}"
        if path.suffix.lower() in _LOOSE_SUFFIXES:
            loose[path.parent].append(path)
        elif host_token_for_file_name(path.name) is not None:
            case = _office_case(path, label)
            if case is not None:
                cases.append(case)
    for folder, paths in sorted(loose.items()):
        case = _loose_case(paths, f"{root.name}/{folder.relative_to(root).as_posix()}", host)
        if case is not None:
            cases.append(case)
    return cases


def _markdown_cases(pin: Path, host: str | None) -> list[dict[str, Any]]:
    cases: list[dict[str, Any]] = []
    corpus = pin / "syntax_corpus"
    for markdown in sorted(corpus.glob("**/*.md")):
        inside = False
        block: list[str] = []
        count = 0
        for line in markdown.read_text(encoding="utf-8").splitlines():
            if not inside and _FENCE_OPEN.match(line.strip()):
                inside, block = True, []
            elif inside and line.strip().startswith("```"):
                inside = False
                count += 1
                source = "\n".join(block) + "\n"
                case: dict[str, Any] = {
                    "label": f"md:{markdown.relative_to(corpus).as_posix()}#{count}",
                    "referenced": [],
                    "modules": [
                        {
                            "name": "Module1",
                            "type": "class" if _CLASS_HEADER.search(source) else "standard",
                            "source": source,
                        }
                    ],
                }
                if host:
                    case["host"] = host
                cases.append(case)
            elif inside:
                block.append(line)
    return cases


def build_cases(paths: list[Path], markdown_pin: Path | None, host: str | None) -> list[dict[str, Any]]:
    """Cases from each path (an Office file, or a folder searched for Office files
    and folders of exported modules), plus the pin's markdown blocks if asked.
    ``host`` is named by the cases no file implies a host for."""
    cases = _markdown_cases(markdown_pin, host) if markdown_pin is not None else []
    for path in paths:
        if path.is_dir():
            cases.extend(_folder_cases(path, host))
        elif path.is_file():
            case = _office_case(path, path.name)
            if case is not None:
                cases.append(case)
        else:
            raise SystemExit(f"{path} does not exist.")
    return cases
