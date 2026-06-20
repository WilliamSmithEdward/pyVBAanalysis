"""Per-pass shared state for the diagnostics engine.

Ported from analysisContext.ts. RulePassContext carries the member-completion
context (member_ctx) as a first-class field: it is assembled once per pass and the
member-not-found, object-state, call-shape, type-of-is, and assignment rules read
it through ctx.member_ctx. The host object model reaches that context through
AnalyzeModuleOptions.host_model.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass
from typing import Any, Protocol

from ..completion import MemberCompletionContext
from ..conditional import ConditionalActivityTracker, ConditionalCompilationEnvironment
from ..lexer.token_helpers import statement_tokens as _compute_statement_tokens
from ..lexer.token_kinds import VbaToken
from ..parser.nodes import ModuleNode, Span
from ..symbols.symbol_model import (
    ModuleSymbolKind,
    ModuleSymbols,
    VbaProcedureSignature,
    VbaProjectClassMembers,
    VbaProjectTypeName,
    VbaSymbol,
)
from .model import VbaDiagnosticData


@dataclass(slots=True)
class AnalyzeModuleOptions:
    """Inputs for analyze_module.

    document_type and host_model are typed Any to avoid importing the completion
    and host packages here; their concrete types are EventHandlerDocumentType and
    HostObjectModel. Both are read by rules: document_type drives the
    event-handler-module-scope rule, and host_model feeds the member-completion
    context and the type-name resolver.
    """

    module_name: str | None = None
    module_kind: ModuleSymbolKind | None = None
    # True when the cross-module fields below represent the COMPLETE project (every
    # module that could define a symbol visible here). False for a partial view, e.g.
    # a single file analyzed in isolation; that suppresses the rules that need the
    # whole project (undeclared-variable, unknown-call, member-not-found), since a
    # symbol declared in an unseen module is then indistinguishable from an undefined
    # one and reporting it would be a false positive.
    whole_project: bool = True
    # Honor ``'@pyvba-ignore`` suppression directives in the source. Set False to
    # report every diagnostic regardless of in-source suppression (an audit run).
    inline_suppression: bool = True
    document_type: Any = None  # EventHandlerDocumentType (from the completion package)
    # Per-rule severity overrides keyed by stable diagnostic code; "off" disables.
    severity_overrides: Mapping[str, str] | None = None
    # Lowercased procedure names callable as bare identifiers from this module.
    known_procedures: AbstractSet[str] | None = None
    # Lowercased bare identifiers visible from this module.
    known_identifiers: AbstractSet[str] | None = None
    # Exported callable signatures grouped by lowercased procedure name.
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None = None
    project_class_members: Sequence[VbaProjectClassMembers] | None = None
    # Project-defined type names (class/document/userform, user Type, Enum) visible
    # to this module, the registry the type-name resolver searches.
    project_types: Sequence[VbaProjectTypeName] | None = None
    project_visible_symbols: Sequence[VbaSymbol] | None = None
    known_non_type_names: AbstractSet[str] | None = None
    project_integer_constants: Mapping[str, str | None] | None = None
    host_model: Any = None  # HostObjectModel (from the host package)
    conditional_compilation: ConditionalCompilationEnvironment | None = None
    parsed_module: ModuleNode | None = None


class PushFn(Protocol):
    """The diagnostics sink every rule reports through."""

    def __call__(
        self, rule: str, message: str, span: Span, data: VbaDiagnosticData | None = None
    ) -> None: ...


@dataclass(slots=True)
class RulePassContext:
    """Everything one diagnostics pass computes once and every rule shares.

    `member_ctx` is the member-resolution context primed with the per-pass AST and
    full-source token stream (mirrors analysisContext.ts memberCtx assembly); the
    member-not-found and object-state rules read it.
    """

    source: str
    module_name: str
    module_kind: ModuleSymbolKind
    opts: AnalyzeModuleOptions
    mod: ModuleNode
    symbols: ModuleSymbols
    activity: ConditionalActivityTracker | None
    member_ctx: MemberCompletionContext


def is_object_module_kind(module_kind: ModuleSymbolKind | None) -> bool:
    """True for the object module kinds (class, document, userform) that own a Me."""
    return module_kind in (
        ModuleSymbolKind.CLASS,
        ModuleSymbolKind.DOCUMENT,
        ModuleSymbolKind.USERFORM,
    )


# Statement-token cache (audit #5): independent rules re-tokenize the same
# statement many times per pass. Tokens are cached per source string (LRU of 2)
# and per statement span, so one pass lexes each statement once. Callers must not
# mutate the returned lists.
_STATEMENT_TOKEN_CACHE_MAX = 2
_statement_token_cache: list[tuple[str, dict[tuple[int, int], list[VbaToken]]]] = []


def statement_tokens(source: str, span: Span) -> list[VbaToken]:
    """Significant tokens of a statement span (no comments/newlines), memoized per pass."""
    entry: dict[tuple[int, int], list[VbaToken]] | None = None
    for i, (cached_source, by_span) in enumerate(_statement_token_cache):
        if cached_source == source:
            entry = by_span
            if i > 0:
                _statement_token_cache.insert(0, _statement_token_cache.pop(i))
            break
    if entry is None:
        entry = {}
        _statement_token_cache.insert(0, (source, entry))
        if len(_statement_token_cache) > _STATEMENT_TOKEN_CACHE_MAX:
            _statement_token_cache.pop()
    key = (span.start, span.end)
    toks = entry.get(key)
    if toks is None:
        toks = _compute_statement_tokens(source, span.start, span.end)
        entry[key] = toks
    return toks
