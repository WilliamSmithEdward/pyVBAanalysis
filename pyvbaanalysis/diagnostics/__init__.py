"""VBA diagnostics: the diagnostic model, rule metadata, and the analysis engine."""

from .model import (
    DiagnosticCategory,
    DiagnosticEvidenceKind,
    DiagnosticSeverity,
    DiagnosticSuppressionScope,
    VbaCreateProcedureStubData,
    VbaDiagnostic,
    VbaDiagnosticData,
    VbaEdit,
    VbaMissingRequiredArgumentPlaceholderData,
)
from .rule_metadata import (
    DEFAULT_DIAGNOSTIC_SUPPRESSION_SCOPES,
    DIAGNOSTIC_RULES,
    DiagnosticRuleMetadata,
    load_rule_metadata,
    rule_metadata_by_code,
)

__all__ = [
    "DiagnosticSeverity",
    "DiagnosticCategory",
    "DiagnosticEvidenceKind",
    "DiagnosticSuppressionScope",
    "VbaDiagnostic",
    "VbaDiagnosticData",
    "VbaEdit",
    "VbaMissingRequiredArgumentPlaceholderData",
    "VbaCreateProcedureStubData",
    "DiagnosticRuleMetadata",
    "DIAGNOSTIC_RULES",
    "DEFAULT_DIAGNOSTIC_SUPPRESSION_SCOPES",
    "load_rule_metadata",
    "rule_metadata_by_code",
]
