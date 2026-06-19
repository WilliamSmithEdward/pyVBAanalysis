"""Member-access completion engine (reduced port of src/analyzer/completion).

Exposes the single seam the diagnostics consume: ``resolve_member_surface_at``
(wrapped by ``rules.shared.resolve_exhaustive_member_surface``) plus the
``MemberCompletionContext`` the engine assembles per pass. Completion-UX paths
(completion rows, signatures, docs, definitions) are intentionally not ported.
"""

from __future__ import annotations

from .member_access import (
    MemberCompletionContext,
    MemberCompletionEntry,
    ResolvedMemberSurface,
    is_known_object_assignment_type,
    resolve_member_surface_at,
    resolve_receiver_type_at,
)

__all__ = [
    "MemberCompletionContext",
    "MemberCompletionEntry",
    "ResolvedMemberSurface",
    "is_known_object_assignment_type",
    "resolve_member_surface_at",
    "resolve_receiver_type_at",
]
