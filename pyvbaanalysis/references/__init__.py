"""Reference analysis: how each occurrence of a name uses it (read, write, or both).

Ported from xlide_vscode/src/analyzer/references.
"""

from __future__ import annotations

from .reference_kinds import ReferenceKind, classify_reference_kinds

__all__ = ["ReferenceKind", "classify_reference_kinds"]
