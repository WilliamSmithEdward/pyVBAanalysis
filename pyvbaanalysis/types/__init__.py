"""VBA type model: pure type-name helpers now; the inference engine (M8) later."""

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
