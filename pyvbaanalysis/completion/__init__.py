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
from .type_completion import (
    OLE_AUTOMATION_TYPES,
    VBA_PRIMITIVE_TYPES,
    TypeCompletion,
    TypeCompletionKind,
    host_type_names,
    is_creatable_type_completion,
    project_type_candidates,
    resolve_type_name,
    type_completion_candidates,
)

__all__ = [
    "MemberCompletionContext",
    "MemberCompletionEntry",
    "ResolvedMemberSurface",
    "is_known_object_assignment_type",
    "resolve_member_surface_at",
    "resolve_receiver_type_at",
    "OLE_AUTOMATION_TYPES",
    "VBA_PRIMITIVE_TYPES",
    "TypeCompletion",
    "TypeCompletionKind",
    "host_type_names",
    "is_creatable_type_completion",
    "project_type_candidates",
    "resolve_type_name",
    "type_completion_candidates",
]
