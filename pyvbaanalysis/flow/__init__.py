"""Control-flow analysis helpers (procedure labels)."""

from .procedure_labels import (
    VbaProcedureLabel,
    VbaProcedureLabelReference,
    collect_procedure_label_declarations,
    collect_procedure_label_references,
    collect_procedure_labels,
    statement_label_references,
)

__all__ = [
    "VbaProcedureLabel",
    "VbaProcedureLabelReference",
    "collect_procedure_label_declarations",
    "collect_procedure_label_references",
    "collect_procedure_labels",
    "statement_label_references",
]
