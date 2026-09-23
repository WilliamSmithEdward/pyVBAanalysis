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
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from ..conditional import ConditionalCompilationEnvironment
from ..host.host_libraries import referenced_host_tokens
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


def _require_excel_extension(file_path: Path) -> None:
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


def read_workbook_modules(path: str | Path) -> list[LoadedModule]:
    """Every VBA module in an Excel file as a LoadedModule (name, kind, code body).

    Supports the macro-enabled Excel formats pyOpenVBA reads (.xlsm, .xlsb, .xlam,
    and legacy .xls). Raises WorkbookReadError if the extension is not an Excel
    workbook or the container has no readable VBA project.
    """
    pyopenvba = _require_pyopenvba()
    file_path = Path(path)
    _require_excel_extension(file_path)
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

    The Excel-only form of analyze_office_file: any other extension is refused, and
    an Excel file gets exactly the same analysis, the workbook's reference list
    included, so the two entry points never disagree about one file.

    Returns a dict mapping module name to that module's diagnostics. Pass ``only`` to
    report just the named modules while still indexing the whole project for context.
    ``conditional_compilation`` sets a project-wide #If/#Const baseline.
    """
    _require_excel_extension(Path(path))
    return analyze_office_file(
        path,
        only=only,
        severity_overrides=severity_overrides,
        conditional_compilation=conditional_compilation,
        inline_suppression=inline_suppression,
    )


def _read_access_project(
    pyopenvba: Any, file_path: Path
) -> tuple[list[LoadedModule], list[str] | None]:
    """Every VBA module in an Access database, and its references' libids. Read-only
    by construction: Access executes compiled p-code, so pyOpenVBA offers no write
    path and neither do we.

    The libids come back None, UNKNOWN rather than empty, when the dir catalog cannot
    be read: a project whose reference list was not seen has not been shown to lack
    any library.
    """
    modules: list[LoadedModule] = []
    with pyopenvba.AccessReader(file_path) as database:
        # An Access module carries no designer header, and its class modules name
        # the generic VBA class base rather than a document coclass, so the text
        # alone cannot tell class from standard. The dir catalog can.
        try:
            info = database.read_project_info()
        except Exception:
            info = None
        class_names = (
            {entry.name for entry in info.modules if entry.is_class_module}
            if info is not None
            else set()
        )
        for name in database.vba_module_names():
            modules.append(
                loaded_module_from_text(
                    database.read_vba_module_with_attributes(name),
                    name=name,
                    pyopenvba_standard=(name not in class_names),
                )
            )
    libids = _reference_libids(info.references) if info is not None else None
    return modules, libids


def _reference_libids(references: Iterable[Any]) -> list[str]:
    """The libid of each reference that has one, in declaration order.

    A project reference or a control reference may carry no registered libid; it
    names no Office library, so it is skipped rather than guessed at.
    """
    return [ref.libid for ref in references if isinstance(getattr(ref, "libid", None), str)]


def _read_office_project(path: str | Path) -> tuple[list[LoadedModule], list[str] | None]:
    """Every VBA module in an Office macro container, and the libid of every library
    its project references, from one open of the container. The libids are None when
    the reference list could not be read, which is different from an empty list."""
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
            return _read_access_project(pyopenvba, file_path)
        if suffix in WORD_EXTENSIONS:
            opener = pyopenvba.WordFile
        elif suffix in POWERPOINT_EXTENSIONS:
            opener = pyopenvba.PowerPointFile
        else:
            opener = pyopenvba.ExcelFile
        standard_kind = pyopenvba.VBAModuleKind.standard
        modules: list[LoadedModule] = []
        with opener(file_path) as container:
            project = container.vba_project()
            for component in project.modules:
                modules.append(
                    loaded_module_from_text(
                        component.source,
                        name=component.name,
                        pyopenvba_standard=(component.kind == standard_kind),
                    )
                )
            libids = _reference_libids(project.references)
        return modules, libids
    except WorkbookReadError:
        raise
    except Exception as exc:
        # A corrupt or unsupported container raises pyOpenVBA errors, and also raw
        # zipfile / struct errors from the underlying format parsing. Wrap them all
        # at this untrusted-file boundary so callers get a clean WorkbookReadError.
        raise WorkbookReadError(f"Could not read VBA from {file_path}: {exc}") from exc


@dataclass(frozen=True, slots=True)
class OfficeProject:
    """One Office macro container, read: its modules, the host its extension
    implies, and the other Office libraries its project references.

    ``referenced_hosts`` is None when the reference list could not be read, which
    is different from an empty list: a project whose references were not seen has
    not been shown to lack any library.
    """

    modules: list[LoadedModule]
    host: str | None
    referenced_hosts: list[str] | None


def read_office_project(path: str | Path) -> OfficeProject:
    """Every VBA module in an Office macro container, with its host and the other
    Office libraries its project references, from one open of the container."""
    modules, libids = _read_office_project(path)
    host = host_token_for_file_name(Path(path).name)
    return OfficeProject(
        modules=modules,
        host=host,
        referenced_hosts=referenced_host_tokens(host, libids) if libids is not None else None,
    )


def read_office_modules(path: str | Path) -> list[LoadedModule]:
    """Every VBA module in an Office macro container as a LoadedModule.

    Covers every host pyOpenVBA reads: Excel (.xlsm, .xlsb, .xlam, .xls), Word
    (.docm, .dotm, .doc), PowerPoint (.pptm, .potm) and Access (.accdb, .mdb,
    read-only). Raises WorkbookReadError if the extension is not a readable container
    or the container has no readable VBA project.
    """
    return _read_office_project(path)[0]


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
    project = read_office_project(path)
    inputs = [
        ModuleInput(module_name=module.name, module_kind=module.kind, source=module.source)
        for module in project.modules
    ]
    return analyze_project(
        inputs,
        only=only,
        severity_overrides=severity_overrides,
        conditional_compilation=conditional_compilation,
        inline_suppression=inline_suppression,
        host=project.host,
        # The container carries the project's reference list, so a library the
        # project does not reference is provably absent here, and a Word document
        # that references Excel is analyzed against both.
        referenced_hosts=project.referenced_hosts,
    )
