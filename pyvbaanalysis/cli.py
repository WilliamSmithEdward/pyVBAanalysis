"""Command-line interface: analyze VBA from files, folders, or Excel workbooks.

Run with ``python -m pyvbaanalysis``. Each PATH may be:

* a loose ``.bas`` / ``.cls`` / ``.frm`` export file,
* a folder (its loose export files are analyzed together as one project), or
* a macro-enabled Excel workbook (read via pyOpenVBA).

Loose files and folders are pooled into a single project so cross-module references
resolve; each workbook is analyzed as its own project. Exit codes (so it slots into
a CI gate): 0 when everything analyzed cleanly, 1 when diagnostics were reported or
a file could not be read, and 2 for a usage error (missing path or no analyzable
input).
"""

from __future__ import annotations

import argparse
import json
import sys
from collections.abc import Iterable, Sequence
from pathlib import Path

from . import __version__
from .diagnostics import (
    STRUCTURAL_DIAGNOSTIC_RULES,
    VbaDiagnostic,
    line_col,
    rule_metadata_by_code,
    validate_severity_overrides,
)
from .project import analyze_project
from .reader import (
    EXCEL_EXTENSIONS,
    LOOSE_EXTENSIONS,
    LooseFileReadError,
    WorkbookReadError,
    load_loose_module,
    read_workbook_modules,
)
from .symbols import ModuleInput


_SEVERITY_RANK = {"information": 0, "warning": 1, "error": 2}


def _parse_severity_overrides(items: Sequence[str] | None) -> dict[str, str]:
    """Parse repeated ``--severity CODE=LEVEL`` flags into an overrides mapping."""
    overrides: dict[str, str] = {}
    for item in items or []:
        code, sep, level = item.partition("=")
        if not sep or not code.strip():
            raise ValueError(f"--severity expects CODE=LEVEL, got {item!r}")
        overrides[code.strip()] = level.strip()
    return overrides


def _filter_codes(
    results: Sequence[_ProjectResult],
    select: Sequence[str] | None,
    ignore: Sequence[str] | None,
) -> None:
    """Keep only --select codes and drop --ignore codes from each result, in place.

    Codes are matched case-insensitively (diagnostic codes are canonically lowercase).
    """
    selected = {code.strip().lower() for code in select} if select else None
    ignored = {code.strip().lower() for code in ignore} if ignore else None
    for result in results:
        for module, diagnostics in result.diagnostics.items():
            if selected is not None:
                diagnostics = [d for d in diagnostics if d.code in selected]
            if ignored is not None:
                diagnostics = [d for d in diagnostics if d.code not in ignored]
            result.diagnostics[module] = diagnostics


def _gather_loose_paths(paths: Iterable[Path]) -> list[Path]:
    """Expand folders into their loose export files, keep loose files as given."""
    out: list[Path] = []
    for path in paths:
        if path.is_dir():
            out.extend(
                sorted(
                    p
                    for p in path.rglob("*")
                    if p.is_file() and p.suffix.lower() in LOOSE_EXTENSIONS
                )
            )
        else:
            out.append(path)
    return out


def _analyze_loose_group(
    paths: Sequence[Path],
    only: Sequence[str] | None,
    severity_overrides: dict[str, str] | None,
) -> tuple[list[_ProjectResult], list[str], list[str]]:
    if not paths:
        return [], [], []
    modules = []
    loaded_paths = []
    errors: list[str] = []
    for path in paths:
        try:
            modules.append(load_loose_module(path))
            loaded_paths.append(path)
        except LooseFileReadError as exc:
            errors.append(str(exc))
    if not modules:
        return [], [], errors
    # Pooled files can collide on module name (same VB_Name or stem); the project
    # requires unique names, so keep the first and warn about the rest.
    seen: dict[str, Path] = {}
    unique = []
    warnings: list[str] = []
    for module, path in zip(modules, loaded_paths):
        key = module.name.lower()
        if key in seen:
            warnings.append(
                f"duplicate module name {module.name!r} ({path} and {seen[key]}); "
                "analyzing the first only"
            )
            continue
        seen[key] = path
        unique.append(module)
    inputs = [
        ModuleInput(module_name=m.name, module_kind=m.kind, source=m.source) for m in unique
    ]
    diagnostics = analyze_project(
        inputs, only=only or None, severity_overrides=severity_overrides or None
    )
    sources = {m.name: m.source for m in unique}
    return [_ProjectResult("(loose files)", diagnostics, sources)], warnings, errors


def _analyze_workbook_group(
    paths: Sequence[Path],
    only: Sequence[str] | None,
    severity_overrides: dict[str, str] | None,
) -> tuple[list[_ProjectResult], list[str]]:
    results: list[_ProjectResult] = []
    errors: list[str] = []
    for path in paths:
        try:
            modules = read_workbook_modules(path)
        except WorkbookReadError as exc:
            errors.append(f"{path}: {exc}")
            continue
        inputs = [
            ModuleInput(module_name=m.name, module_kind=m.kind, source=m.source) for m in modules
        ]
        diagnostics = analyze_project(
            inputs, only=only or None, severity_overrides=severity_overrides or None
        )
        sources = {m.name: m.source for m in modules}
        results.append(_ProjectResult(str(path), diagnostics, sources))
    return results, errors


class _ProjectResult:
    """One analyzed project (a workbook or the pooled loose files) for reporting."""

    def __init__(
        self,
        label: str,
        diagnostics: dict[str, list[VbaDiagnostic]],
        sources: dict[str, str],
    ) -> None:
        self.label = label
        self.diagnostics = diagnostics
        self.sources = sources


def _render_text(results: Sequence[_ProjectResult]) -> str:
    lines: list[str] = []
    for result in results:
        lines.append(f"# {result.label}")
        any_in_project = False
        for module_name, diagnostics in result.diagnostics.items():
            if not diagnostics:
                continue
            any_in_project = True
            source = result.sources.get(module_name, "")
            lines.append(f"  {module_name}")
            for diag in diagnostics:
                line, column = line_col(source, diag.span.start)
                lines.append(
                    f"    {line}:{column} {diag.severity.value} {diag.code}  {diag.message}"
                )
        if not any_in_project:
            lines.append("  (no diagnostics)")
    return "\n".join(lines)


def _render_json(results: Sequence[_ProjectResult]) -> str:
    payload = []
    for result in results:
        modules = []
        for module_name, diagnostics in result.diagnostics.items():
            source = result.sources.get(module_name, "")
            modules.append(
                {
                    "module": module_name,
                    "diagnostics": [
                        {
                            "code": diag.code,
                            "severity": diag.severity.value,
                            "message": diag.message,
                            "start": diag.span.start,
                            "end": diag.span.end,
                            "line": line_col(source, diag.span.start)[0],
                            "column": line_col(source, diag.span.start)[1],
                            "spec_reference": diag.spec_reference,
                        }
                        for diag in diagnostics
                    ],
                }
            )
        payload.append({"project": result.label, "modules": modules})
    return json.dumps(payload, indent=2)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pyvbaanalysis",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        description="Static analysis for Excel VBA: analyze loose .bas/.cls/.frm files, "
        "folders of them, or Excel workbooks.",
        epilog=(
            "exit codes:\n"
            "  0  no diagnostics at or above the fail level\n"
            "  1  diagnostics reported (at the fail level), or a file could not be read\n"
            "  2  usage error (a path was not found, or nothing analyzable was given)\n"
        ),
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    parser.add_argument("paths", nargs="+", type=Path, help="files, folders, or Excel workbooks")
    parser.add_argument(
        "--only",
        action="append",
        metavar="NAME",
        help="analyze only the named module(s); repeatable (project context still uses all)",
    )
    parser.add_argument(
        "--severity",
        action="append",
        metavar="CODE=LEVEL",
        help="override a code's severity (LEVEL is off/information/warning/error); repeatable",
    )
    parser.add_argument(
        "--select",
        action="append",
        metavar="CODE",
        help="report only these diagnostic codes; repeatable",
    )
    parser.add_argument(
        "--ignore",
        action="append",
        metavar="CODE",
        help="hide these diagnostic codes from the report; repeatable",
    )
    parser.add_argument(
        "--fail-level",
        choices=("error", "warning", "information"),
        default="information",
        dest="fail_level",
        help="exit non-zero only when a diagnostic at or above this severity is reported "
        "(default: information, i.e. any diagnostic)",
    )
    parser.add_argument(
        "--format",
        choices=("text", "json"),
        default="text",
        help="output format (default: text)",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    """Entry point. Returns a process exit code:

    * 0 - analyzed cleanly, no diagnostics.
    * 1 - diagnostics were reported, or a file could not be read or analyzed.
    * 2 - usage error: a path was not found, or nothing analyzable was given.
    """
    args = _build_parser().parse_args(argv)
    paths = [Path(p) for p in args.paths]

    missing = [str(p) for p in paths if not p.exists()]
    if missing:
        print("error: path not found: " + ", ".join(missing), file=sys.stderr)
        return 2

    try:
        overrides = _parse_severity_overrides(args.severity)
        validate_severity_overrides(overrides)
    except ValueError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    valid_codes = set(rule_metadata_by_code()) | {
        meta.code for meta in STRUCTURAL_DIAGNOSTIC_RULES.values()
    }
    filter_codes = [*(args.select or []), *(args.ignore or [])]
    unknown_codes = sorted({c for c in filter_codes if c.strip().lower() not in valid_codes})
    if unknown_codes:
        print("error: unknown diagnostic code(s): " + ", ".join(unknown_codes), file=sys.stderr)
        return 2

    workbook_paths = [p for p in paths if p.is_file() and p.suffix.lower() in EXCEL_EXTENSIONS]
    other_paths = [p for p in paths if p not in workbook_paths]
    loose_paths = _gather_loose_paths(other_paths)
    unknown = [p for p in loose_paths if p.suffix.lower() not in LOOSE_EXTENSIONS]
    loose_paths = [p for p in loose_paths if p.suffix.lower() in LOOSE_EXTENSIONS]

    loose_results, warnings, loose_errors = _analyze_loose_group(loose_paths, args.only, overrides)
    workbook_results, workbook_errors = _analyze_workbook_group(workbook_paths, args.only, overrides)
    results = [*loose_results, *workbook_results]
    errors = [*loose_errors, *workbook_errors]
    _filter_codes(results, args.select, args.ignore)

    for path in unknown:
        print(f"warning: skipping unsupported file {path}", file=sys.stderr)
    for warning in warnings:
        print(f"warning: {warning}", file=sys.stderr)
    for error in errors:
        print(f"error: {error}", file=sys.stderr)

    if not results:
        if errors:
            return 1  # a workbook failed to read; the error was printed above
        print("error: no analyzable VBA modules found", file=sys.stderr)
        return 2

    rendered = _render_json(results) if args.format == "json" else _render_text(results)
    print(rendered)

    fail_rank = _SEVERITY_RANK[args.fail_level]
    counted = sum(
        1
        for result in results
        for diagnostics in result.diagnostics.values()
        for diag in diagnostics
        if _SEVERITY_RANK[diag.severity.value] >= fail_rank
    )
    return 1 if (counted > 0 or errors) else 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
