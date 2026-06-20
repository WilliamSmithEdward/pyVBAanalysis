"""pyVBAanalysis: pure-Python static analysis for Excel VBA.

It reports a diagnostic only when it is provably correct; anything unknown or
ambiguous stays quiet (a no-false-positive discipline).

The headline entry points are re-exported here for convenience:

* analyze_module(source, opts): analyze one module's source text.
* analyze_project(modules): analyze a set of modules with cross-module context.
* analyze_workbook(path): analyze the VBA in an Excel workbook (read via pyOpenVBA,
  the one external runtime dependency).
* analyze_loose_file(path) / analyze_loose_files(paths): analyze loose
  .bas / .cls / .frm export files.

The full analysis model (tokens, AST, symbols, types) lives in the submodules
(pyvbaanalysis.lexer, .parser, .symbols, .diagnostics, and the rest); see
docs/api-reference.md for the import map.
"""

from __future__ import annotations

from .conditional import DEFAULT_COMPILER_CONSTANTS, ConditionalCompilationEnvironment
from .diagnostics import (
    AnalyzeModuleOptions,
    DiagnosticSeverity,
    VbaDiagnostic,
    analyze_module,
    line_col,
    rule_metadata_by_code,
    validate_severity_overrides,
)
from .project import analyze_module_options_for, analyze_project, build_project_index
from .reader import analyze_loose_file, analyze_loose_files, analyze_workbook
from .symbols import ModuleInput, ModuleSymbolKind, ProjectIndex, ProjectIndexOptions

__version__ = "1.0.0"

__all__ = [
    "AnalyzeModuleOptions",
    "ConditionalCompilationEnvironment",
    "DEFAULT_COMPILER_CONSTANTS",
    "DiagnosticSeverity",
    "ModuleInput",
    "ModuleSymbolKind",
    "ProjectIndex",
    "ProjectIndexOptions",
    "VbaDiagnostic",
    "__version__",
    "analyze_loose_file",
    "analyze_loose_files",
    "analyze_module",
    "analyze_module_options_for",
    "analyze_project",
    "analyze_workbook",
    "build_project_index",
    "line_col",
    "rule_metadata_by_code",
    "validate_severity_overrides",
]
