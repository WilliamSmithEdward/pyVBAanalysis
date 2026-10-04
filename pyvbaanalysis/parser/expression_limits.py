"""Ported from xlide_vscode/src/analyzer/parser/expressionLimits.ts."""

from __future__ import annotations

# Shared recovery limit for recursive expression parsing, value folding and array shapes.
MAX_EXPRESSION_DEPTH = 256
