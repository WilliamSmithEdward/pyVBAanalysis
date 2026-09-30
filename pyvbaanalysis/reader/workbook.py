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

import re
from collections.abc import Iterable, Mapping
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Any

from ..conditional import ConditionalCompilationEnvironment, ConditionalValue, parse_project_conditional_constants
from ..host.host_libraries import referenced_host_tokens
from ..host.host_registry import host_token_for_file_name
from ..diagnostics import VbaDiagnostic
from ..project import analyze_project
from ..symbols import ImplicitMember, ModuleSymbolKind
from .vbe_module import LoadedModule, loaded_module_from_text

# The Excel container extensions read_workbook_modules / analyze_workbook accept.
EXCEL_EXTENSIONS = frozenset({".xlsm", ".xlsb", ".xlam", ".xls"})
# Every container extension read_office_modules / analyze_office_file accept, by
# host token. Scoped to what pyOpenVBA 6.3.3 actually opens: the host registry
# recognizes more extensions than the engine can read (.xltm, .xlt, .xla, .dot,
# .ppam, .ppsm, .ppa are refused by extension there), and claiming one we cannot
# open would trade a clear error for a confusing one.
WORD_EXTENSIONS = frozenset({".docm", ".dotm", ".doc"})
# A legacy .ppt keeps its VBA project in a zlib-compressed CFB inside an
# ExOleObjStg record; pyOpenVBA reads it from 6.x on (3.4.0 did not).
POWERPOINT_EXTENSIONS = frozenset({".pptm", ".potm", ".ppt"})
# .mda and .accda are Access add-ins, stored as a database is.
ACCESS_EXTENSIONS = frozenset({".accdb", ".mdb", ".accda", ".mda"})
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


# dir stream record ids ([MS-OVBA] 2.3.4.2). PROJECTVERSION has no size field:
# its body is a fixed 10 bytes. PROJECTMODULES opens the module records, past
# which no project information record occurs.
_REC_PROJECTCODEPAGE = 0x0003
_REC_PROJECTVERSION = 0x0009
_REC_PROJECTCONSTANTS = 0x000C
_REC_PROJECTMODULES = 0x000F
_REC_PROJECTCONSTANTS_UNICODE = 0x003C


def _project_conditional_constants_text(dir_raw: bytes) -> str | None:
    """The project's Conditional Compilation Arguments as the VBE shows them
    (``Name = Value : Name = Value``), read from its decompressed dir stream, or
    None when it declares none.

    The UTF-16LE record is preferred, as upstream prefers it: it holds the same
    text and needs no code page. A writer that emits only the MBCS record gets
    it decoded in the project's code page.
    """
    records: dict[int, bytes] = {}
    pos = 0
    while pos + 6 <= len(dir_raw):
        record_id = int.from_bytes(dir_raw[pos : pos + 2], "little")
        if record_id == _REC_PROJECTVERSION:
            pos += 12
            continue
        size = int.from_bytes(dir_raw[pos + 2 : pos + 6], "little")
        records.setdefault(record_id, dir_raw[pos + 6 : pos + 6 + size])
        pos += 6 + size
        if record_id == _REC_PROJECTMODULES:
            break
    unicode_text = records.get(_REC_PROJECTCONSTANTS_UNICODE)
    if unicode_text:
        return unicode_text.decode("utf-16-le", errors="replace")
    mbcs_text = records.get(_REC_PROJECTCONSTANTS)
    if not mbcs_text:
        return None
    code_page_record = records.get(_REC_PROJECTCODEPAGE, b"")
    code_page = int.from_bytes(code_page_record[:2], "little") if len(code_page_record) >= 2 else 1252
    from pyopenvba.vba import encoding_for_codepage

    return mbcs_text.decode(encoding_for_codepage(code_page), errors="replace")


def _project_constants(dir_raw: bytes | None) -> dict[str, ConditionalValue]:
    """The parsed Conditional Compilation Arguments of a dir stream; empty when
    it declares none or could not be read."""
    if not dir_raw:
        return {}
    return parse_project_conditional_constants(_project_conditional_constants_text(dir_raw))


def _read_access_project(
    pyopenvba: Any, file_path: Path
) -> tuple[list[LoadedModule], list[str] | None, dict[str, ConditionalValue]]:
    """Every VBA module in an Access database, and its references' libids. Read-only
    by construction: modules come through pyOpenVBA's AccessReader, and form and
    report designs through an AccessDatabase opened from the file's bytes, which
    has no path to save back to.

    A form's or report's code module is a UserForm-like module whose designer
    class is Access.Form or Access.Report, with the design's sections and
    controls as its members, as upstream reads it.

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
        # By lowercased name: the catalog and the module list can spell one module
        # in two cases (`Basket` and `basket`), and VBA names are case-insensitive.
        class_names = (
            {entry.name.lower() for entry in info.modules if entry.is_class_module}
            if info is not None
            else set()
        )
        for name in database.vba_module_names():
            modules.append(
                loaded_module_from_text(
                    database.read_vba_module_with_attributes(name),
                    name=name,
                    pyopenvba_standard=(name.lower() not in class_names),
                )
            )
    database = _access_database(file_path)
    designs = _access_designs(database) if database is not None else {}
    try:
        dir_raw = database.dir_stream()[0] if database is not None else None
    except Exception:
        dir_raw = None
    modules = [
        replace(
            module,
            kind=ModuleSymbolKind.USERFORM,
            designer_class=designs[module.name.lower()][0],
            implicit_members=designs[module.name.lower()][1],
        )
        if module.name.lower() in designs
        else module
        for module in modules
    ]
    libids = _reference_libids(info.references) if info is not None else None
    return modules, libids, _project_constants(dir_raw)


# The Access library's class for a control type whose class is not simply its
# type name (upstream ACCESS_CONTROL_CLASSES in accessDesign.ts).
_ACCESS_CONTROL_CLASSES = {
    "CheckBox": "Checkbox",
    "TextBox": "Textbox",
    "ComboBox": "Combobox",
    "Subform": "SubForm",
    "Tab": "TabControl",
    "WebBrowser": "WebBrowserControl",
    "EdgeBrowser": "Edge",
}
_ACCESS_DESIGN_CLASSES = {"form": "Access.Form", "report": "Access.Report"}
_ACCESS_MODULE_PREFIXES = {"form": "Form_", "report": "Report_"}
_ACCESS_IDENTIFIER_UNSAFE = re.compile(r"[^A-Za-z0-9_\u0080-￿]")


def _access_vba_identifier(name: str) -> str:
    """The name VBA knows an Access section or control by: each ASCII character
    an identifier cannot hold becomes an underscore, and a name that then opens
    with a digit or an underscore takes `Ctl` in front. `Order Date` is
    `Order_Date`, `2ndBox` is `Ctl2ndBox` (upstream accessVbaIdentifier)."""
    converted = _ACCESS_IDENTIFIER_UNSAFE.sub("_", name)
    return f"Ctl{converted}" if re.match(r"[0-9_]", converted) else converted


def _access_database(file_path: Path) -> Any:
    """The database opened from its bytes, so nothing can be written back to the
    file, or None when it cannot be opened that way."""
    try:
        from pyopenvba.access import AccessDatabase

        return AccessDatabase(file_path.read_bytes())
    except Exception:
        return None


def _access_designs(database: Any) -> dict[str, tuple[str, tuple[ImplicitMember, ...]]]:
    """Each form's and report's code module, by lowercased name, with the class
    its design makes it and the sections and controls that are its members.

    A design that cannot be read leaves its module as the dir catalog classed
    it. Members are never the whole surface: a bound form also has one for
    every field of its record source, which only the running database knows.
    """
    out: dict[str, tuple[str, tuple[ImplicitMember, ...]]] = {}
    try:
        designs = [("form", design) for design in database.forms()]
        designs += [("report", design) for design in database.reports()]
    except Exception:
        return out
    for kind, design in designs:
        taken: set[str] = set()
        members: list[ImplicitMember] = []
        for obj in design.objects[1:]:
            if not obj.name:
                continue
            name = _access_vba_identifier(obj.name)
            if name.lower() in taken:
                continue
            taken.add(name.lower())
            if obj.is_section:
                members.append(ImplicitMember(name, "Access.Section"))
            elif obj.type_name:
                members.append(
                    ImplicitMember(name, f"Access.{_ACCESS_CONTROL_CLASSES.get(obj.type_name, obj.type_name)}")
                )
            else:
                members.append(ImplicitMember(name, "Access.Control"))
        module_name = f"{_ACCESS_MODULE_PREFIXES[kind]}{design.name}"
        out[module_name.lower()] = (_ACCESS_DESIGN_CLASSES[kind], tuple(members))
    return out


def _form_controls(container: Any) -> dict[str, tuple[ImplicitMember, ...]]:
    """Each UserForm's controls, by lowercased form name, read from the form's
    designer storage: the members its code-behind uses and never declares.

    Every control counts, a Frame's or MultiPage's children included, each
    typed as pyOpenVBA names its class (``MSForms.TextBox``, or
    ``ActiveX.Control`` for a control from another library). MSForms names are
    unique across a form, so a name seen twice is one control. A form whose
    designer streams do not reconcile gets no entry, the way upstream leaves a
    form it cannot read without members.
    """
    out: dict[str, tuple[ImplicitMember, ...]] = {}
    try:
        forms = container.forms()
    except Exception:
        return out
    for form in forms:
        seen: set[str] = set()
        members: list[ImplicitMember] = []
        for control in form.walk():
            if not control.name or control.name.lower() in seen:
                continue
            seen.add(control.name.lower())
            members.append(ImplicitMember(control.name, _FORM_CONTROL_CLASSES.get(control.kind, control.kind)))
        out[form.name.lower()] = tuple(members)
    return out


# A MultiPage's pages are sites whose class pyOpenVBA reads as the form's own,
# MSForms.Form; to VBA each is an MSForms.Page, as upstream types it.
_FORM_CONTROL_CLASSES = {"MSForms.Form": "MSForms.Page"}


def _reference_libids(references: Iterable[Any]) -> list[str]:
    """The libid of each reference that has one, in declaration order.

    A project reference or a control reference may carry no registered libid; it
    names no Office library, so it is skipped rather than guessed at.
    """
    return [ref.libid for ref in references if isinstance(getattr(ref, "libid", None), str)]


def _read_office_project(
    path: str | Path,
) -> tuple[list[LoadedModule], list[str] | None, dict[str, ConditionalValue]]:
    """Every VBA module in an Office macro container, the libid of every library
    its project references, and its Conditional Compilation Arguments, from one
    open of the container. The libids are None when the reference list could not
    be read, which is different from an empty list."""
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
            constants = _project_constants(getattr(project, "dir_raw", None))
            controls = _form_controls(container)
        modules = [
            replace(module, implicit_members=controls[module.name.lower()])
            if module.kind is ModuleSymbolKind.USERFORM and module.name.lower() in controls
            else module
            for module in modules
        ]
        return modules, libids, constants
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

    ``project_constants`` are the project's Conditional Compilation Arguments
    (Tools > Project Properties in the VBE), which every module's ``#If`` sees.
    """

    modules: list[LoadedModule]
    host: str | None
    referenced_hosts: list[str] | None
    project_constants: Mapping[str, ConditionalValue] = field(default_factory=dict)

    def conditional_compilation(
        self, caller: ConditionalCompilationEnvironment | None = None
    ) -> ConditionalCompilationEnvironment | None:
        """The #If baseline for this project: the caller's environment with the
        project's own constants added. A constant the caller sets wins, whatever
        case it is spelled in, since VBA names are case-insensitive."""
        if not self.project_constants:
            return caller
        caller_constants = dict(caller.project_constants or {}) if caller is not None else {}
        overridden = {name.lower() for name in caller_constants}
        merged = {
            name: value for name, value in self.project_constants.items() if name.lower() not in overridden
        }
        merged.update(caller_constants)
        return ConditionalCompilationEnvironment(
            compiler_constants=caller.compiler_constants if caller is not None else None,
            project_constants=merged,
        )


def read_office_project(path: str | Path) -> OfficeProject:
    """Every VBA module in an Office macro container, with its host, the other
    Office libraries its project references and its Conditional Compilation
    Arguments, from one open of the container."""
    modules, libids, constants = _read_office_project(path)
    host = host_token_for_file_name(Path(path).name)
    return OfficeProject(
        modules=modules,
        host=host,
        referenced_hosts=referenced_host_tokens(host, libids) if libids is not None else None,
        project_constants=constants,
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
    whole project for context. ``conditional_compilation`` adds to, and overrides,
    the Conditional Compilation Arguments the project itself declares.
    """
    project = read_office_project(path)
    inputs = [module.as_module_input() for module in project.modules]
    return analyze_project(
        inputs,
        only=only,
        severity_overrides=severity_overrides,
        # The project's own Conditional Compilation Arguments decide its #If
        # branches, as they do when the VBE compiles it; a constant the caller
        # passes overrides the file's.
        conditional_compilation=project.conditional_compilation(conditional_compilation),
        inline_suppression=inline_suppression,
        host=project.host,
        # The container carries the project's reference list, so a library the
        # project does not reference is provably absent here, and a Word document
        # that references Excel is analyzed against both.
        referenced_hosts=project.referenced_hosts,
    )
