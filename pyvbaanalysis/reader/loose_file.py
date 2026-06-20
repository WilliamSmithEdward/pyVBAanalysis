"""Analyze loose VBA export files (.bas / .cls / .frm) from disk.

These are the files the VBE produces on "Export File" and the files pyOpenVBA's
``pull`` writes. Each carries its export header, so loading one means reading the
text, classifying the module kind (from the extension and header), deriving the
module name, and stripping the designer header before analysis. Loading several at
once runs them as one project so cross-module references resolve.
"""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from pathlib import Path

from ..conditional import ConditionalCompilationEnvironment
from ..diagnostics import VbaDiagnostic
from ..project import analyze_project
from ..symbols import ModuleInput
from .vbe_module import LoadedModule, loaded_module_from_text

# The loose VBE export extensions the loaders recognize.
LOOSE_EXTENSIONS = frozenset({".bas", ".cls", ".frm"})


class LooseFileReadError(Exception):
    """A loose .bas/.cls/.frm file could not be read from disk."""

# VBE exports are CP1252 by default; tolerate UTF-8 (with or without BOM) too. These
# can fail on bytes they do not map; latin-1 (which decodes every byte and never
# raises) is the final fallback below, so a stray byte never aborts a load.
_VBE_ENCODINGS = ("utf-8-sig", "cp1252")


def _read_text(path: Path) -> str:
    raw = path.read_bytes()
    for encoding in _VBE_ENCODINGS:
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1")


def load_loose_module(path: str | Path) -> LoadedModule:
    """Load one .bas/.cls/.frm file into a LoadedModule (name, kind, code body).

    Raises LooseFileReadError if the path cannot be read (it is a directory, is
    missing, or is not readable), so callers never see a raw OS exception.
    """
    file_path = Path(path)
    try:
        text = _read_text(file_path)
    except OSError as exc:
        raise LooseFileReadError(f"could not read {file_path}: {exc}") from exc
    return loaded_module_from_text(
        text, extension=file_path.suffix, name_fallback=file_path.stem
    )


def analyze_loose_file(
    path: str | Path,
    *,
    severity_overrides: Mapping[str, str] | None = None,
    conditional_compilation: ConditionalCompilationEnvironment | None = None,
    whole_project: bool = False,
    inline_suppression: bool = True,
) -> list[VbaDiagnostic]:
    """Analyze a single loose VBA file and return its diagnostics.

    The file is analyzed as a one-module project. Because that is a partial view of
    any real project, ``whole_project`` defaults to False, which suppresses the rules
    that need every module (undeclared-variable, unknown-call, member-not-found) so a
    symbol defined in another file is not reported as undefined. Pass
    ``whole_project=True`` if this file genuinely is the entire project, or use
    analyze_loose_files to analyze several files together with shared context.
    ``conditional_compilation`` sets the #If/#Const baseline.
    """
    module = load_loose_module(path)
    results = analyze_project(
        [ModuleInput(module_name=module.name, module_kind=module.kind, source=module.source)],
        severity_overrides=severity_overrides,
        conditional_compilation=conditional_compilation,
        whole_project=whole_project,
        inline_suppression=inline_suppression,
    )
    return results[module.name]


def analyze_loose_files(
    paths: Iterable[str | Path],
    *,
    only: Iterable[str] | None = None,
    severity_overrides: Mapping[str, str] | None = None,
    conditional_compilation: ConditionalCompilationEnvironment | None = None,
    whole_project: bool = True,
    inline_suppression: bool = True,
) -> dict[str, list[VbaDiagnostic]]:
    """Analyze several loose VBA files as one project with cross-module context.

    Returns a dict mapping module name to that module's diagnostics. Pass ``only`` to
    report just the named modules while still indexing every file for context.
    ``conditional_compilation`` sets a project-wide #If/#Const baseline. ``whole_project``
    defaults to True (these files are treated as the whole project); pass False if they
    are only a fragment, to suppress the rules that need every module.
    """
    modules = [load_loose_module(path) for path in paths]
    inputs = [
        ModuleInput(module_name=module.name, module_kind=module.kind, source=module.source)
        for module in modules
    ]
    return analyze_project(
        inputs,
        only=only,
        severity_overrides=severity_overrides,
        conditional_compilation=conditional_compilation,
        whole_project=whole_project,
        inline_suppression=inline_suppression,
    )
