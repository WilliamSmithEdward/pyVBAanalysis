"""Read and analyze VBA modules from Office macro containers.

This is the only part of pyVBAanalysis that touches a binary Office container, and
the only importer of pyOpenVBA (the one external runtime dependency, used for direct
VBA reads). Every host pyOpenVBA reads is covered here (Excel, Word, PowerPoint, and
read-only Access), and the file's extension selects both the container reader and the
host object model the rules resolve against, so Word VBA is never measured against
Excel's surface. pyOpenVBA is imported lazily inside the functions so
``import pyvbaanalysis`` stays light.

pyOpenVBA yields each component's full export text (header included) plus a coarse
standard/other kind. The shared vbe_module helper refines the kind and strips the
designer header, so a workbook and a folder of loose files go through the same code.
"""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from pathlib import Path
from typing import Any

from ..conditional import ConditionalCompilationEnvironment
from ..host.host_registry import host_token_for_file_name
from ..diagnostics import VbaDiagnostic
from ..project import analyze_project
from ..symbols import ModuleInput
from .vbe_module import LoadedModule, loaded_module_from_text

# The Excel container extensions read_workbook_modules / analyze_workbook accept.
EXCEL_EXTENSIONS = frozenset({".xlsm", ".xlsb", ".xlam", ".xls"})
# Every container extension read_office_modules / analyze_office_file accept, by
# host token. Scoped to what pyOpenVBA 3.4.0 actually opens: the host registry
# recognizes more extensions (templates and add-ins) than the engine can read, and
# claiming one we cannot open would trade a clear error for a confusing one.
WORD_EXTENSIONS = frozenset({".docm", ".dotm", ".doc"})
# .ppt is deliberately absent: pyOpenVBA 3.4.0 lists it but reads it as a plain
# CFB, and a legacy .ppt keeps its VBA project in a zlib-compressed CFB inside an
# ExOleObjStg record, so every open fails on the missing dir stream. Listing it
# would trade a clear "unsupported extension" for a confusing parse error.
POWERPOINT_EXTENSIONS = frozenset({".pptm", ".potm"})
ACCESS_EXTENSIONS = frozenset({".accdb", ".mdb"})
OFFICE_EXTENSIONS = EXCEL_EXTENSIONS | WORD_EXTENSIONS | POWERPOINT_EXTENSIONS | ACCESS_EXTENSIONS


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
    suffix = file_path.suffix.lower()
    if suffix not in EXCEL_EXTENSIONS:
        hint = (
            " Use read_office_modules / analyze_office_file for non-Excel hosts."
            if suffix in OFFICE_EXTENSIONS
            else ""
        )
        raise WorkbookReadError(
            f"Unsupported file extension {file_path.suffix!r}; expected an Excel workbook "
            f"({', '.join(sorted(EXCEL_EXTENSIONS))}).{hint}"
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
    inline_suppression: bool = True,
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
        inline_suppression=inline_suppression,
    )


def _read_access_modules(pyopenvba: Any, file_path: Path) -> list[LoadedModule]:
    """Every VBA module in an Access database. Read-only by construction: Access
    executes compiled p-code, so pyOpenVBA offers no write path and neither do we."""
    modules: list[LoadedModule] = []
    with pyopenvba.AccessReader(file_path) as database:
        # An Access module carries no designer header, and its class modules name
        # the generic VBA class base rather than a document coclass, so the text
        # alone cannot tell class from standard. The dir catalog can.
        try:
            class_names = {
                entry.name for entry in database.read_project_info().modules
                if entry.is_class_module
            }
        except Exception:
            class_names = set()
        for name in database.vba_module_names():
            modules.append(
                loaded_module_from_text(
                    database.read_vba_module_with_attributes(name),
                    name=name,
                    pyopenvba_standard=(name not in class_names),
                )
            )
    return modules


def read_office_modules(path: str | Path) -> list[LoadedModule]:
    """Every VBA module in an Office macro container as a LoadedModule.

    Covers every host pyOpenVBA reads: Excel (.xlsm, .xlsb, .xlam, .xls), Word
    (.docm, .dotm, .doc), PowerPoint (.pptm, .potm, .ppt) and Access (.accdb, .mdb,
    read-only). Raises WorkbookReadError if the extension is not a readable container
    or the container has no readable VBA project.
    """
    pyopenvba = _require_pyopenvba()
    file_path = Path(path)
    suffix = file_path.suffix.lower()
    if suffix not in OFFICE_EXTENSIONS:
        raise WorkbookReadError(
            f"Unsupported file extension {file_path.suffix!r}; expected an Office macro "
            f"container ({', '.join(sorted(OFFICE_EXTENSIONS))})."
        )
    try:
        if suffix in ACCESS_EXTENSIONS:
            return _read_access_modules(pyopenvba, file_path)
        if suffix in WORD_EXTENSIONS:
            opener = pyopenvba.WordFile
        elif suffix in POWERPOINT_EXTENSIONS:
            opener = pyopenvba.PowerPointFile
        else:
            opener = pyopenvba.ExcelFile
        standard_kind = pyopenvba.VBAModuleKind.standard
        modules: list[LoadedModule] = []
        with opener(file_path) as container:
            for component in container.vba_project().modules:
                modules.append(
                    loaded_module_from_text(
                        component.source,
                        name=component.name,
                        pyopenvba_standard=(component.kind == standard_kind),
                    )
                )
        return modules
    except WorkbookReadError:
        raise
    except Exception as exc:
        # A corrupt or unsupported container raises pyOpenVBA errors, and also raw
        # zipfile / struct errors from the underlying format parsing. Wrap them all
        # at this untrusted-file boundary so callers get a clean WorkbookReadError.
        raise WorkbookReadError(f"Could not read VBA from {file_path}: {exc}") from exc


def analyze_office_file(
    path: str | Path,
    *,
    only: Iterable[str] | None = None,
    severity_overrides: Mapping[str, str] | None = None,
    conditional_compilation: ConditionalCompilationEnvironment | None = None,
    inline_suppression: bool = True,
) -> dict[str, list[VbaDiagnostic]]:
    """Analyze every VBA module in an Office macro container with full cross-module
    context, resolving against the host object model the container implies.

    The file's extension selects the host, so Word VBA answers to Word's model and
    never to Excel's. Returns a dict mapping module name to that module's
    diagnostics; ``only`` reports just the named modules while still indexing the
    whole project for context.
    """
    modules = read_office_modules(path)
    inputs = [
        ModuleInput(module_name=module.name, module_kind=module.kind, source=module.source)
        for module in modules
    ]
    return analyze_project(
        inputs,
        only=only,
        severity_overrides=severity_overrides,
        conditional_compilation=conditional_compilation,
        inline_suppression=inline_suppression,
        host=host_token_for_file_name(Path(path).name),
    )
