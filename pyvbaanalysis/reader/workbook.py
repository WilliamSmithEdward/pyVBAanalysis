"""Read and analyze VBA modules from Excel files.

This is the only part of pyVBAanalysis that touches a binary Office container, and
the only importer of pyOpenVBA (the one external runtime dependency, used for direct
VBA reads). pyVBAanalysis analyzes Excel VBA: the rules resolve against the Excel
host object model, so reading is scoped to Excel workbooks. pyOpenVBA is imported
lazily inside the functions so ``import pyvbaanalysis`` stays light.

pyOpenVBA yields each component's full export text (header included) plus a coarse
standard/other kind. The shared vbe_module helper refines the kind and strips the
designer header, so a workbook and a folder of loose files go through the same code.
"""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from pathlib import Path
from typing import Any

from ..conditional import ConditionalCompilationEnvironment
from ..diagnostics import VbaDiagnostic
from ..project import analyze_project
from ..symbols import ModuleInput
from .vbe_module import LoadedModule, loaded_module_from_text

# The Excel container extensions read_workbook_modules / analyze_workbook accept.
EXCEL_EXTENSIONS = frozenset({".xlsm", ".xlsb", ".xlam", ".xls"})


class WorkbookReadError(RuntimeError):
    """Raised when an Excel file cannot be opened or its VBA cannot be read."""


def _require_pyopenvba() -> Any:
    try:
        import pyopenvba
    except ImportError as exc:
        raise WorkbookReadError(
            "pyOpenVBA is required to read VBA from Excel files. Reinstall pyvbaanalysis "
            "(pyOpenVBA is a dependency)."
        ) from exc
    return pyopenvba


def read_workbook_modules(path: str | Path) -> list[LoadedModule]:
    """Every VBA module in an Excel file as a LoadedModule (name, kind, code body).

    Supports the macro-enabled Excel formats pyOpenVBA reads (.xlsm, .xlsb, .xlam,
    and legacy .xls). Raises WorkbookReadError if the extension is not an Excel
    workbook or the container has no readable VBA project.
    """
    pyopenvba = _require_pyopenvba()
    file_path = Path(path)
    if file_path.suffix.lower() not in EXCEL_EXTENSIONS:
        raise WorkbookReadError(
            f"Unsupported file extension {file_path.suffix!r}; expected an Excel workbook "
            f"({', '.join(sorted(EXCEL_EXTENSIONS))})."
        )
    standard_kind = pyopenvba.VBAModuleKind.standard
    modules: list[LoadedModule] = []
    try:
        with pyopenvba.ExcelFile(file_path) as workbook:
            for component in workbook.vba_project().modules:
                modules.append(
                    loaded_module_from_text(
                        component.source,
                        name=component.name,
                        pyopenvba_standard=(component.kind == standard_kind),
                    )
                )
    except WorkbookReadError:
        raise
    except Exception as exc:
        # A corrupt or unsupported container raises pyOpenVBA errors, and also raw
        # zipfile / struct errors from the underlying format parsing. Wrap them all
        # at this untrusted-file boundary so callers get a clean WorkbookReadError.
        raise WorkbookReadError(f"Could not read VBA from {file_path}: {exc}") from exc
    return modules


def analyze_workbook(
    path: str | Path,
    *,
    only: Iterable[str] | None = None,
    severity_overrides: Mapping[str, str] | None = None,
    conditional_compilation: ConditionalCompilationEnvironment | None = None,
) -> dict[str, list[VbaDiagnostic]]:
    """Analyze every VBA module in an Excel file with full cross-module context.

    Returns a dict mapping module name to that module's diagnostics. Pass ``only`` to
    report just the named modules while still indexing the whole project for context.
    ``conditional_compilation`` sets a project-wide #If/#Const baseline.
    """
    modules = read_workbook_modules(path)
    inputs = [
        ModuleInput(module_name=module.name, module_kind=module.kind, source=module.source)
        for module in modules
    ]
    return analyze_project(
        inputs,
        only=only,
        severity_overrides=severity_overrides,
        conditional_compilation=conditional_compilation,
    )
