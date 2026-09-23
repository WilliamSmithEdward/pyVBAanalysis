"""VBA type model: the pure type-name helpers (expression typing is in diagnostics)."""

from .type_names import (
    NumericBounds,
    is_known_scalar_type,
    is_numeric_type,
    normalize_type,
    numeric_literal_bounds,
)

__all__ = [
    "NumericBounds",
    "is_known_scalar_type",
    "is_numeric_type",
    "normalize_type",
    "numeric_literal_bounds",
]
