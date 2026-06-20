"""The diagnostics model: the diagnostic shape and its enums.

Ported from xlide_vscode/src/analyzer/diagnostics/ruleMetadata.ts (the enums) and
analysisContext.ts (the VbaDiagnostic shape). Offset-based, host-agnostic; no
editor dependency.
"""

from __future__ import annotations

import enum
from dataclasses import dataclass

from ..parser.nodes import Span


class DiagnosticSeverity(str, enum.Enum):
    """Severity of a diagnostic, independent of any editor enum."""

    ERROR = "error"
    WARNING = "warning"
    INFORMATION = "information"


class DiagnosticCategory(str, enum.Enum):
    """Broad purpose bucket used by tests, docs, and filtering."""

    SYNTAX = "syntax"
    LEXER = "lexer"
    PARSER = "parser"
    REALTIME_RECOVERY = "realtime-recovery"
    DECLARATION = "declaration"
    SEMANTIC = "semantic"
    PROJECT_SYMBOL = "project-symbol"
    MODULE_KIND = "module-kind"
    EXCEL_HOST = "excel-host"
    STYLE = "style"


class DiagnosticEvidenceKind(str, enum.Enum):
    """Why a rule is surfaced at its default severity."""

    COMPILE_ERROR = "compile-error"
    DETERMINISTIC_RUNTIME_ERROR = "deterministic-runtime-error"
    RUNTIME_RISK = "runtime-risk"
    STYLE_POLICY = "style-policy"


class DiagnosticSuppressionScope(str, enum.Enum):
    """Analysis suppression scopes exposed by the workbook analysis UI."""

    BLOCK = "block"
    MEMBER = "member"
    MODULE = "module"


@dataclass(frozen=True, slots=True)
class VbaEdit:
    """A deterministic text edit attached to a diagnostic's code-action data."""

    span: Span
    new_text: str


@dataclass(frozen=True, slots=True)
class VbaMissingRequiredArgumentPlaceholderData:
    parameter_name: str
    edit: VbaEdit


@dataclass(frozen=True, slots=True)
class VbaCreateProcedureStubData:
    procedure_name: str
    edit: VbaEdit


@dataclass(frozen=True, slots=True)
class VbaDiagnosticData:
    """Optional structured data for deterministic editor actions."""

    missing_required_argument_placeholder: VbaMissingRequiredArgumentPlaceholderData | None = None
    create_procedure_stub: VbaCreateProcedureStubData | None = None


@dataclass(frozen=True, slots=True)
class VbaDiagnostic:
    """A single diagnostic produced by the analyzer (offset-based)."""

    code: str
    message: str
    # Effective severity (after any user override).
    severity: DiagnosticSeverity
    span: Span
    # MS-VBAL (or other) reference for the rule, when known.
    spec_reference: str | None = None
    data: VbaDiagnosticData | None = None


def line_col(source: str, offset: int) -> tuple[int, int]:
    """The 1-based (line, column) of a character offset in ``source``.

    Offsets are the unit of ``VbaDiagnostic.span`` (``span.start`` / ``span.end``).
    The column counts characters from the start of the line. An offset past the end
    of the source clamps to the end.
    """
    clamped = max(0, min(offset, len(source)))
    line = source.count("\n", 0, clamped) + 1
    last_newline = source.rfind("\n", 0, clamped)
    return line, clamped - (last_newline + 1) + 1
