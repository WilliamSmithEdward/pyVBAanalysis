"""VBA diagnostics: the diagnostic model, rule metadata, and the analysis engine."""

from .analyze_module import analyze_module
from .context import AnalyzeModuleOptions, PushFn, RulePassContext, is_object_module_kind
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
from .registry import DIAGNOSTIC_RULE_REGISTRY, DiagnosticRuleEntry
from .rule_metadata import (
    DEFAULT_DIAGNOSTIC_SUPPRESSION_SCOPES,
    DIAGNOSTIC_RULES,
    STRUCTURAL_DIAGNOSTIC_RULES,
    DiagnosticRuleMetadata,
    diagnostic_metadata_for_code,
    load_rule_metadata,
    normalize_diagnostic_severity_override,
    rule_metadata_by_code,
)

__all__ = [
    "analyze_module",
    "AnalyzeModuleOptions",
    "RulePassContext",
    "PushFn",
    "is_object_module_kind",
    "DiagnosticRuleEntry",
    "DIAGNOSTIC_RULE_REGISTRY",
    "STRUCTURAL_DIAGNOSTIC_RULES",
    "diagnostic_metadata_for_code",
    "normalize_diagnostic_severity_override",
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
