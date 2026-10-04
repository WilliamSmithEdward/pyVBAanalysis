"""Rule family: object-variable state.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/objectState.ts: member
access on a variable declared as a scalar, and an object variable that is still
Nothing when it is used, which raises Run-time error '91'. The latter tracks an
unset->set lattice per local over the shared dataflow walk, falling back to the
GoTo-following straight-line walk for procedures with unstructured flow.
"""

from __future__ import annotations

import itertools
import re
from collections.abc import Callable, Iterator, Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass, field

from ...completion import MemberCompletionContext, is_known_object_assignment_type
from ...completion.member_access import precedes_leading_member_dot
from ...conditional import ConditionalActivityTracker
from ...flow.procedure_labels import statement_label_declarations, statement_label_references
from ...flow.procedure_unstructured import procedure_has_unstructured_flow
from ...identity_cache import IdentityLru
from ...js_compat import js_number, js_number_to_string
from ...lexer.token_helpers import match_paren_from, split_top_level_token_groups
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    DoBlockNode,
    ForBlockNode,
    IfBlockNode,
    IfBranchKind,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    ProcKind,
    SelectBlockNode,
    Span,
    StatementNode,
    VariableGroupNode,
    WhileBlockNode,
    WithBlockNode,
    is_leaf_statement,
    iter_body_nodes,
)
from ...symbols.name_resolution import BareIdentifierContext, BareIdentifierResolutionScope
from ...symbols.symbol_model import ModuleSymbols, SymbolVisibility, VbaSymbol, VbaSymbolKind
from ...types.type_inference import (
    SourceDeclaredType,
    declared_type_for_source_binding,
    def_type_of,
    function_result_for,
    object_holding_default,
    object_value_needs_index,
    procedure_symbol_for,
    read_only_host_default,
    return_assignment_type_for,
    source_identifier_binding,
    statement_may_change_module_variable,
    type_environment_for,
    unreachable_statements_in,
)
from ...types.type_names import is_known_scalar_type, normalize_type
from ..block_headers import block_header_statements
from ..condition_operands import condition_operands
from ..condition_value import ConditionFacts, condition_value
from ..context import PushFn, statement_tokens
from ..dataflow import (
    Lattice,
    StraightLineDataflowHooks,
    walk_branch_merged_body,
    walk_straight_line_body,
)
from ..module_state import untouched_module_variables_in
from ..straight_line_values import OBJECT_NOTHING
from ..walker import (
    ProcedureStatementVisitor,
    active_module_members,
    bare_assignment_target,
    block_header_line_span,
    for_each_statement,
    is_inactive_node,
    locals_named_whole,
    raw_expression_tokens,
    set_assignment_target,
    statement_and_branch_spans,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)
from .host_arguments import literal_intersect_is_nothing, range_method_owner
from .shared import ONE_VALUE_BUILTINS, builtin_name_before, resolve_exhaustive_member_surface
from .type_of_is import object_let_assignment_verdict

# The operators that read an object's default member as an operand.
_OPERAND_OPERATORS: frozenset[str] = frozenset(
    {"=", "<", ">", "<=", ">=", "<>", "+", "-", "*", "/", "\\", "&", "^"}
)

# `/\(\s*\)\s*$/`: a return type that is an array.
_ARRAY_RETURN_RE = re.compile(r"\(\s*\)\s*\Z")
# `/[%&^]$/`: an integer literal's type suffix.
_INTEGER_SUFFIX_RE = re.compile(r"[%&^]\Z")


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    """`toks[i]` as JavaScript reads it: undefined below 0 and past the end."""
    return toks[i] if 0 <= i < len(toks) else None


def _raw(tok: VbaToken | None) -> str | None:
    return tok.raw_text if tok is not None else None


def _host_name(member_ctx: MemberCompletionContext) -> str | None:
    model = member_ctx.model
    return model.get("hostName") if model is not None else None


def _word_re(name: str) -> re.Pattern[str]:
    """`new RegExp(`\\b${name}\\b`)`: JavaScript's ASCII word boundary."""
    return re.compile(rf"\b{re.escape(name)}\b", re.ASCII)


def _set_nothing_re(name: str) -> re.Pattern[str]:
    """`new RegExp(`\\bset\\s+${name}\\s*=\\s*nothing\\b`, 'i')`."""
    return re.compile(rf"\bset\s+{re.escape(name)}\s*=\s*nothing\b", re.IGNORECASE | re.ASCII)


def _literal_index_text(raw: str) -> str:
    """`${Number(raw.replace(/[%&^]$/, ''))}`: an integer literal's index as a key part."""
    return js_number_to_string(js_number(_INTEGER_SUFFIX_RE.sub("", raw, count=1)))


def check_scalar_member_access(
    source: str,
    symbols: ModuleSymbols,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    push: PushFn,
    member_ctx: MemberCompletionContext | None = None,
) -> ProcedureStatementVisitor:
    """Per-statement rule: rides the shared procedure-statement walk (audit #0)."""
    ctx = member_ctx if member_ctx is not None else MemberCompletionContext()
    # The project's other standard modules: `Foo.Foo()` names module Foo
    # before its Function Foo, so the Function's Long is no receiver (issue #403).
    other_modules = {
        surface.name.lower()
        for surface in ctx.project_class_members or []
        if surface.kind == "standardModule" and surface.name.lower() != symbols.module_name.lower()
    }

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        env = type_environment_for(symbols, member)
        proc_sym = procedure_symbol_for(symbols, member)

        def resolve_declared_type(name: str) -> SourceDeclaredType:
            if (
                name.lower() in other_modules
                and source_identifier_binding(
                    symbols,
                    proc_sym,
                    project_visible_symbols,
                    name,
                    BareIdentifierContext.MEMBER_RECEIVER,
                ).scope
                is BareIdentifierResolutionScope.PROJECT
            ):
                return SourceDeclaredType(resolved=True)
            return declared_type_for_source_binding(
                symbols,
                proc_sym,
                project_visible_symbols,
                name,
                BareIdentifierContext.MEMBER_RECEIVER,
            )

        def visitor(stmt: LeafStatementNode) -> None:
            for name, as_type, span, vbe_error in _scalar_member_accesses(
                source, stmt.span, env, resolve_declared_type
            ):
                push(
                    "scalarMemberAccess",
                    f"Member access on '{name}' is invalid because it is declared as {as_type}. "
                    f"This is a VBE compile error: {vbe_error}.",
                    span,
                )

        return visitor

    return factory


def _scalar_member_accesses(
    source: str,
    span: Span,
    env: Mapping[str, str],
    resolve_declared_type: Callable[[str], SourceDeclaredType] | None = None,
) -> list[tuple[str, str, Span, str]]:
    toks = statement_tokens(source, span)
    out: list[tuple[str, str, Span, str]] = []
    for i in range(len(toks) - 1):
        if toks[i + 1].raw_text != ".":
            continue
        if _raw(_at(toks, i - 1)) == ".":
            continue
        name = token_name(toks[i])
        if not name:
            continue
        declared_type = resolve_declared_type(name) if resolve_declared_type is not None else None
        as_type = (
            declared_type.as_type
            if declared_type is not None and declared_type.resolved
            else env.get(name.lower())
        )
        normalized = normalize_type(as_type)
        if not as_type or not normalized or not is_known_scalar_type(normalized):
            continue
        member_name = token_name(toks[i + 2]) if i + 2 < len(toks) else None
        vbe_error = "Invalid qualifier" if member_name else "Syntax error"
        out.append(
            (
                name,
                as_type,
                Span(span.start + toks[i].start, span.start + toks[i + 1].end),
                vbe_error,
            )
        )
    return out


@dataclass(slots=True)
class _LocalObjectVariable:
    name: str
    as_type: str
    # A Function's own result: Nothing until the function Sets it, so a Let
    # into it raises 91 (issue #193). Only a Let reads it; inside the function
    # its name with a dot or in a With is a recursive call.
    let_only: bool = False
    # A Variant: Nothing only once `Set v = Nothing` (issue #343). Only member reads are judged.
    variant: bool = False
    # A variable of the module, followed from `Set mc = Nothing` in the procedure (issue #618).
    module: bool = False


# 'unset' | 'set' | 'unknown'
ObjectVariableState = str

_Finding = tuple[str, str, Span]


def check_object_variable_not_set(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    member_ctx: MemberCompletionContext,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    nothing_functions = _functions_returning_nothing(source, mod, member_ctx, activity)
    object_functions: dict[str, ProcedureNode] = {}
    for candidate in active_module_members(mod, activity):
        if (
            isinstance(candidate, ProcedureNode)
            and candidate.proc_kind is ProcKind.FUNCTION
            and len(candidate.params) > 0
            and candidate.name.lower() not in nothing_functions
            and bool(candidate.return_type)
            and not _ARRAY_RETURN_RE.search(candidate.return_type or "")
            and is_known_object_assignment_type(candidate.return_type, member_ctx)
        ):
            object_functions[candidate.name.lower()] = candidate
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        if nothing_functions or object_functions:

            def visit_calls(stmt: LeafStatementNode) -> None:
                for span in statement_and_branch_spans(stmt):
                    toks = statement_tokens(source, span)
                    for start, end, message in _nothing_result_member_access(
                        source, toks, nothing_functions
                    ):
                        push(
                            "objectVariableNotSet",
                            message,
                            Span(span.start + start, span.start + end),
                        )
                    for start, end, message in _nothing_call_member_access(
                        source, toks, object_functions, symbols, activity
                    ):
                        push(
                            "objectVariableNotSet",
                            message,
                            Span(span.start + start, span.start + end),
                        )

            for_each_statement(member.body, visit_calls, activity)
        _check_go_to_into_with(source, member, activity, push)
        _check_for_each_over_empty_object_array(source, member, symbols, member_ctx, activity, push)
        # A module variable nothing ever sets is Nothing in every procedure (issue #241).
        untouched: Mapping[str, VbaSymbol] = untouched_module_variables_in(source, symbols, member)
        unset = [
            (lower, variable)
            for lower, variable in untouched.items()
            if variable.as_type is not None
            and is_known_object_assignment_type(variable.as_type, member_ctx)
        ]
        if unset:
            objects = dict(unset)

            def visit_unset(stmt: LeafStatementNode) -> None:
                for span in statement_and_branch_spans(stmt):
                    toks = statement_tokens(source, span)
                    i = 0
                    while i + 2 < len(toks):
                        variable = objects.get((token_name(toks[i]) or "").lower())
                        before = _raw(_at(toks, i - 1))
                        if (
                            variable is None
                            or before == "."
                            or before == "!"
                            or (toks[i + 1].raw_text != "." and toks[i + 1].raw_text != "!")
                            or not token_name(toks[i + 2])
                        ):
                            i += 1
                            continue
                        scope = (
                            "the project"
                            if variable.visibility is SymbolVisibility.PUBLIC
                            or variable.visibility is SymbolVisibility.GLOBAL
                            else "this module"
                        )
                        push(
                            "objectVariableNotSet",
                            f"Object variable '{toks[i].raw_text}' is never set anywhere in {scope}, so it is "
                            "Nothing here. This will raise Run-time error '91': Object variable or With block "
                            "variable not set.",
                            Span(span.start + toks[i].start, span.start + toks[i].end),
                        )
                        i += 1

            for_each_statement(member.body, visit_unset, activity)
        for rule, message, span in _object_state_walk(
            source, mod, member, symbols, member_ctx, activity
        ).findings:
            push(rule, message, span)


# Words that raise or end before a Function could return: Err.Raise, Error, End, Stop.
_PREEMPTING_WORDS: frozenset[str] = frozenset({"raise", "error", "stop"})
_END_FOLLOWERS: tuple[str, ...] = ("function", "if", "select", "with", "sub", "property")


def _functions_returning_nothing(
    source: str,
    mod: ModuleNode,
    member_ctx: MemberCompletionContext,
    activity: ConditionalActivityTracker | None,
) -> dict[str, ProcedureNode]:
    """The Functions of the module, by lowercased name, that return an object
    and never name their result: each returns Nothing (issue #240, measured
    in Excel 16.0), so `F().Count` raises 91. A body that may raise or end
    first is left out."""
    out: dict[str, ProcedureNode] = {}
    for member in active_module_members(mod, activity):
        if (
            not isinstance(member, ProcedureNode)
            or member.proc_kind is not ProcKind.FUNCTION
            or not member.return_type
            or _ARRAY_RETURN_RE.search(member.return_type)
        ):
            continue
        if not is_known_object_assignment_type(member.return_type, member_ctx):
            continue
        lower = member.name.lower()
        body = statement_tokens(source, Span(member.span.start, member.span.end))
        # The header names it once and `End Function` closes it. `Set F =
        # Nothing` names it and still returns Nothing (issue #343).
        named = sum(
            1
            for i, tok in enumerate(body)
            if (token_name(tok) or "").lower() == lower
            and token_name(tok) is not None
            and not (
                token_text(_at(body, i - 1)) == "set"
                and _raw(_at(body, i + 1)) == "="
                and token_text(_at(body, i + 2)) == "nothing"
            )
        )
        preempts = any(
            token_text(tok) in _PREEMPTING_WORDS
            or (
                token_text(tok) == "end"
                and i > 0
                and token_text(_at(body, i + 1)) not in _END_FOLLOWERS
            )
            for i, tok in enumerate(body)
        )
        if named == 1 and not preempts:
            out[lower] = member
    return out


def _nothing_result_member_access(
    source: str,
    toks: Sequence[VbaToken],
    functions: Mapping[str, ProcedureNode],
) -> list[tuple[int, int, str]]:
    """`F().Count` or `F.Count` on a Function that returns Nothing. Offsets are the statement's."""
    out: list[tuple[int, int, str]] = []
    for i in range(len(toks) - 1):
        fn = functions.get((token_name(toks[i]) or "").lower())
        before = _raw(_at(toks, i - 1))
        if fn is None or before == "." or before == "!":
            continue
        end = i
        if toks[i + 1].raw_text == "(":
            depth = 0
            for k in range(i + 1, len(toks)):
                raw = toks[k].raw_text
                depth += 1 if raw == "(" else -1 if raw == ")" else 0
                if depth == 0:
                    end = k
                    break
        elif len(fn.params) > 0:
            continue
        if _raw(_at(toks, end + 1)) != "." or not token_name(_at(toks, end + 2)):
            continue
        sets_nothing = (
            len(fn.body) > 0
            and _set_nothing_re(fn.name).search(source[fn.span.start : fn.span.end]) is not None
        )
        what = (
            "sets its result to Nothing"
            if sets_nothing
            else "never sets its result, so it returns Nothing"
        )
        out.append(
            (
                toks[i].start,
                toks[end].end,
                f"Function '{fn.name}' {what}, and '.{toks[end + 2].raw_text}' has no object to reach. "
                "This will raise Run-time error '91': Object variable or With block variable not set.",
            )
        )
    return out


def _is_literal_argument(arg: Sequence[VbaToken]) -> bool:
    parts = [tok for tok in arg if tok.kind is not TokenKind.COMMENT]
    atom = parts[1:] if parts and parts[0].raw_text == "-" else parts
    return len(atom) == 1 and (
        atom[0].kind is TokenKind.INTEGER_LITERAL
        or atom[0].kind is TokenKind.FLOAT_LITERAL
        or (
            len(parts) == 1
            and (
                atom[0].kind is TokenKind.STRING_LITERAL or token_text(atom[0]) in ("true", "false")
            )
        )
    )


def _nothing_call_member_access(
    source: str,
    toks: Sequence[VbaToken],
    functions: Mapping[str, ProcedureNode],
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
) -> list[tuple[int, int, str]]:
    """`MaybeColl(False).Count`: an object Function of the module called with
    literal arguments, which leave its result Nothing (issue #562). Offsets
    are the statement's."""
    out: list[tuple[int, int, str]] = []
    for i in range(len(toks) - 1):
        fn = functions.get((token_name(toks[i]) or "").lower())
        before = _raw(_at(toks, i - 1))
        if fn is None or before == "." or before == "!" or toks[i + 1].raw_text != "(":
            continue
        close = match_paren_from(toks, i + 1)
        if close < 0 or _raw(_at(toks, close + 1)) != "." or not token_name(_at(toks, close + 2)):
            continue
        args = split_top_level_token_groups(toks, i + 2, ",", close)
        if close == i + 2 or not all(_is_literal_argument(arg) for arg in args):
            continue
        passed: list[Sequence[VbaToken] | None] = []
        for arg in args:
            word = token_text(arg[0] if arg else None)
            passed.append(
                raw_expression_tokens("-1" if word == "true" else "0")
                if word in ("true", "false")
                else arg
            )
        result = function_result_for(source, fn, symbols, activity, passed, True)
        if result is not OBJECT_NOTHING:
            continue
        out.append(
            (
                toks[i].start,
                toks[close].end,
                f"Function '{fn.name}' returns Nothing for these arguments, and '.{toks[close + 2].raw_text}' "
                "has no object to reach. This will raise Run-time error '91': Object variable or With block "
                "variable not set.",
            )
        )
    return out


@dataclass(slots=True)
class _ModuleObjectFacts:
    """What the walk needs to know about the module's other procedures (issue #343)."""

    # The Functions that return Nothing, by lowercased name.
    nothing_functions: Mapping[str, ProcedureNode]
    # The procedures whose object parameter's first use is a member read, by
    # lowercased name: for each such parameter's position, the read as written,
    # `c.Count`. Passed Nothing, the procedure raises 91 there.
    member_first: Mapping[str, Mapping[int, str]]
    # Of `intersect` and `union`, the ones that are Excel's here: no procedure of the project takes the name.
    excel_range_methods: AbstractSet[str]


# Keyed by the module node, the source, the activity and the member context: a
# parse is reused under another host.
_MODULE_OBJECT_FACTS = IdentityLru(capacity=8)


def _module_object_facts(
    source: str,
    mod: ModuleNode,
    member_ctx: MemberCompletionContext,
    activity: ConditionalActivityTracker | None,
) -> _ModuleObjectFacts:
    cached: _ModuleObjectFacts | None = _MODULE_OBJECT_FACTS.get(mod, source, activity, member_ctx)
    if cached is not None:
        return cached
    member_first: dict[str, dict[int, str]] = {}
    for member in active_module_members(mod, activity):
        if (
            not isinstance(member, ProcedureNode)
            or member.proc_kind is ProcKind.PROPERTY_LET
            or member.proc_kind is ProcKind.PROPERTY_SET
        ):
            continue
        reads: dict[int, str] = {}
        for k, param in enumerate(member.params):
            if (
                not param.param_array
                and not param.is_array
                and param.as_type
                and is_known_object_assignment_type(param.as_type, member_ctx)
            ):
                read = _first_use_member_read(source, member, param.name.lower(), activity)
                if read:
                    reads[k] = read
        if reads:
            member_first[member.name.lower()] = reads
    excel_range_methods: set[str] = set()
    host = _host_name(member_ctx)
    if (host if host is not None else "Excel") == "Excel":
        own = {
            member.name.lower()
            for member in active_module_members(mod, activity)
            if isinstance(member, ProcedureNode)
        }
        for surface in member_ctx.project_class_members or []:
            if surface.kind == "standardModule":
                for surface_member in surface.members:
                    own.add(surface_member.name.lower())
        for name in ("intersect", "union"):
            if name not in own:
                excel_range_methods.add(name)
    facts = _ModuleObjectFacts(
        nothing_functions=_functions_returning_nothing(source, mod, member_ctx, activity),
        member_first=member_first,
        excel_range_methods=excel_range_methods,
    )
    _MODULE_OBJECT_FACTS.put(facts, mod, source, activity, member_ctx)
    return facts


# Statement heads that leave, jump, raise or change error handling.
_STRAIGHT_LINE_ENDS: frozenset[str] = frozenset(
    {"on", "resume", "gosub", "goto", "exit", "end", "stop", "return", "error", "err"}
)


def _first_use_member_read(
    source: str,
    member: ProcedureNode,
    lower: str,
    activity: ConditionalActivityTracker | None,
) -> str | None:
    """The member read, `c.Count`, when the first statement of the procedure to
    name `lower` reads a member of it, and every statement before that runs in a
    straight line: no block, label, On Error or single-line If."""
    for node in member.body:
        if is_inactive_node(activity, node) or isinstance(node, VariableGroupNode):
            continue
        if (
            not is_leaf_statement(node)
            or len(statement_and_branch_spans(node)) > 1
            or len(statement_label_declarations(source, node.span)) > 0
        ):
            return None
        toks = statement_tokens_after_leading_label(source, node.span)
        head = token_text(_at(toks, 0))
        if head in _STRAIGHT_LINE_ENDS:
            return None
        at = next(
            (
                i
                for i, tok in enumerate(toks)
                if (token_name(tok) or "").lower() == lower
                and token_name(tok) is not None
                and _raw(_at(toks, i - 1)) != "."
                and _raw(_at(toks, i - 1)) != "!"
            ),
            -1,
        )
        if at < 0:
            continue
        name = token_name(_at(toks, at + 2))
        return (
            f"{toks[at].raw_text}.{name}"
            if head != "set" and _raw(_at(toks, at + 1)) == "." and name
            else None
        )
    return None


@dataclass(slots=True)
class _ObjectStateWalk:
    """What one procedure's object-state walk found, and the state at each Let."""

    findings: list[_Finding] = field(default_factory=list)
    # The state of the target at each bare Let into a tracked object, by the target's offset.
    lets: dict[int, ObjectVariableState] = field(default_factory=dict)


# Keyed by the procedure node; the source, the activity and the member context
# must match too, since a parse is reused under another host.
_OBJECT_STATE_WALKS = IdentityLru(capacity=128)


def object_let_state_at(
    source: str,
    mod: ModuleNode,
    member: ProcedureNode,
    symbols: ModuleSymbols,
    member_ctx: MemberCompletionContext,
    activity: ConditionalActivityTracker | None,
    offset: int,
) -> ObjectVariableState:
    """Whether the object a Let assigns through at `offset` is provably set, or
    provably Nothing, there (issue #193): set-required names 438 only for one
    that holds an object, and leaves one still Nothing to object-variable-not-set.
    One of 'set', 'unset' or 'unknown'."""
    return _object_state_walk(source, mod, member, symbols, member_ctx, activity).lets.get(
        offset, "unknown"
    )


def _object_state_walk(
    source: str,
    mod: ModuleNode,
    member: ProcedureNode,
    symbols: ModuleSymbols,
    member_ctx: MemberCompletionContext,
    activity: ConditionalActivityTracker | None,
) -> _ObjectStateWalk:
    cached: _ObjectStateWalk | None = _OBJECT_STATE_WALKS.get(member, source, activity, member_ctx)
    if cached is not None:
        return cached
    walk = _ObjectStateWalk()

    def collect(rule: str, message: str, span: Span, data: object = None) -> None:
        walk.findings.append((rule, message, span))

    _walk_object_state(
        source,
        _module_object_facts(source, mod, member_ctx, activity),
        member,
        symbols,
        member_ctx,
        activity,
        collect,
        walk.lets,
    )
    _OBJECT_STATE_WALKS.put(walk, member, source, activity, member_ctx)
    return walk


def _walk_object_state(
    source: str,
    facts: _ModuleObjectFacts,
    member: ProcedureNode,
    symbols: ModuleSymbols,
    member_ctx: MemberCompletionContext,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    lets: dict[int, ObjectVariableState],
) -> None:
    locals_ = _local_object_variables_for(source, symbols, member, member_ctx)
    elements = _object_array_elements(source, symbols, member, member_ctx, activity)
    if not locals_ and not elements.keys:
        return
    state: dict[str, ObjectVariableState] = {}
    for key, local in locals_.items():
        # A Variant starts Empty, which is no object and not Nothing; what a
        # module variable holds as the procedure starts is not known.
        state[key] = "unknown" if local.variant or local.module else "unset"
    # Each element of a fixed array of objects is Nothing until Set (issue
    # #489); a dynamic one's are, once ReDim allocates them.
    for key, element in elements.keys.items():
        state[key] = "unset" if elements.arrays[element[0]][1] else "unknown"
    # The locals some statement anywhere in the procedure Sets: a `GoSub`
    # may run any of those statements before control comes back (issue
    # #108), so after it none of them is provably still Nothing.
    set_anywhere: set[str] = set()

    def collect_sets(stmt: LeafStatementNode) -> None:
        for span in statement_and_branch_spans(stmt):
            target = set_assignment_target(source, span)
            lower = target[0].lower() if target is not None else None
            if lower and lower in locals_:
                set_anywhere.add(lower)

    for_each_statement(member.body, collect_sets, activity)
    # The module variables a statement may change through what it calls.
    module_variables = [lower for lower, local in locals_.items() if local.module]

    def module_touches(stmt: LeafStatementNode) -> list[str]:
        return [
            lower
            for lower in module_variables
            if statement_may_change_module_variable(source, symbols, member, stmt.span, lower)
        ]

    # The GoTo-following walk runs the body until its labels settle, and
    # reports on its last run (issue #271).
    silent = False

    def report(rule: str, message: str, span: Span, data: object = None) -> None:
        if not silent:
            push(rule, message, span)

    reporter: PushFn = report
    walk = (
        walk_straight_line_body
        if procedure_has_unstructured_flow(source, member, activity)
        else walk_branch_merged_body
    )
    # The For Each loops with no way out but the end, by node identity.
    nothing_after: set[int] = set()
    # A statement a known guard keeps from running (issue #273), by node identity.
    unreachable = unreachable_statements_in(source, member, symbols, activity)

    def on_statement(stmt: LeafStatementNode) -> None:
        _check_object_variable_not_set_statement(
            source, stmt, locals_, state, set_anywhere, member_ctx, reporter, lets, facts, elements
        )
        # Code the statement runs may set a module variable (issue #618).
        for lower in module_touches(stmt):
            state[lower] = "unknown"

    def on_block(node: BodyNode) -> None:
        # The header runs as the block is entered, with the state as it
        # stands: `For i = 1 To c.Count`, `Select Case c.Count` (issue #233).
        if isinstance(node, (SelectBlockNode, DoBlockNode, WhileBlockNode)) or (
            isinstance(node, ForBlockNode) and not node.each
        ):
            before, after = block_header_statements(source, node)
            if before is not None:
                _check_object_variable_not_set_statement(
                    source,
                    before,
                    locals_,
                    state,
                    set_anywhere,
                    member_ctx,
                    reporter,
                    lets,
                    facts,
                    elements,
                )
            # `Loop Until x` reads x after the body, with what it entered with
            # when the body never names x (issue #424). Any local the line
            # reads counts, `c` of `Loop While c.Count < 1` too (issue #560).
            if after is not None and isinstance(node, DoBlockNode):
                in_body = source[
                    (before.span.end if before is not None else node.span.start) : after.span.start
                ].lower()
                after_toks = statement_tokens_after_leading_label(source, after.span)
                named = False
                for tok in after_toks:
                    name = token_name(tok)
                    lower = name.lower() if name is not None else None
                    if lower is not None and lower in locals_ and _word_re(lower).search(in_body):
                        named = True
                        break
                if not named:
                    _check_object_variable_not_set_statement(
                        source,
                        after,
                        locals_,
                        state,
                        set_anywhere,
                        member_ctx,
                        reporter,
                        lets,
                        facts,
                    )
        # A block If's own line and its ElseIf lines read their conditions
        # as the block is entered (issue #424).
        if isinstance(node, IfBlockNode):
            for branch in node.branches:
                if branch.branch_kind is not IfBranchKind.ELSE:
                    header = StatementNode(
                        span=branch.header_span,
                        raw=source[branch.header_span.start : branch.header_span.end],
                    )
                    _check_object_variable_not_set_statement(
                        source,
                        header,
                        locals_,
                        state,
                        set_anywhere,
                        member_ctx,
                        reporter,
                        lets,
                        facts,
                    )
        # A For Each that runs to its end leaves the control variable
        # Nothing, so an access after the loop is right to report. One
        # the body can leave early - Exit For, or a GoTo out of it -
        # leaves it on the current element, so nothing is proven
        # (issue #108: `Exit For` on the first sheet, then `ws.Name`).
        if isinstance(node, ForBlockNode):
            # `For Each x In c` with c still Nothing raises 424, not 91:
            # the loop asks the collection for its enumerator (issue #121).
            over = (
                node.source_expression.strip().lower()
                if node.each and node.source_expression
                else None
            )
            over_local = locals_.get(over) if over else None
            if (
                over
                and over_local is not None
                and not over_local.let_only
                and not over_local.variant
                and state.get(over) == "unset"
                and node.source_expression_span
            ):
                reporter(
                    "objectVariableNotSet",
                    f"Object variable '{over_local.name}' is Nothing when For Each asks it for its elements. "
                    "This will raise Run-time error '424': Object required.",
                    node.source_expression_span,
                )
            lower = node.control_variable.lower() if node.control_variable else None
            control = locals_.get(lower) if lower else None
            if node.each and lower and control is not None and not control.let_only:
                if not _body_can_leave_loop(source, node, activity) and not control.variant:
                    nothing_after.add(id(node))
                elif state.get(lower) == "unset":
                    state[lower] = "unknown"
            return
        if not isinstance(node, WithBlockNode):
            return
        receiver = _unset_with_object_receiver(source, node, locals_, state, elements)
        if receiver is not None:
            name, element, span = receiver
            what = f"Element {name} of '{element}'" if element else f"Object variable '{name}'"
            reporter(
                "objectVariableNotSet",
                f"{what} is Nothing before With member access. This will raise Run-time error '91': "
                "Object variable or With block variable not set.",
                span,
            )

    # A For Each that ends leaves its control variable Nothing, over an
    # empty collection too (issue #336, measured in Excel 16.0).
    def after_block(node: BodyNode) -> None:
        lower = (
            node.control_variable.lower()
            if isinstance(node, ForBlockNode) and node.control_variable
            else None
        )
        if lower and id(node) in nothing_after:
            state[lower] = "unset"

    def touches_in_statement(stmt: LeafStatementNode) -> set[str]:
        passed: Mapping[str, int] = locals_named_whole(
            source, stmt.span, locals_, _OBJECT_READ_ONLY_INTRINSICS
        )
        touched = set(passed.keys())
        for span in statement_and_branch_spans(stmt):
            for key in _element_touches(
                statement_tokens_after_leading_label(source, span), elements
            ):
                touched.add(key)
        for module_lower in module_touches(stmt):
            touched.add(module_lower)
        # A single-line If's branches Set too. A Let gives a Variant a value.
        for span in statement_and_branch_spans(stmt):
            target = set_assignment_target(source, span)
            set_lower = target[0].lower() if target is not None else None
            if set_lower and set_lower in locals_:
                touched.add(set_lower)
            let_target = bare_assignment_target(source, span)
            let_lower = let_target[0].lower() if let_target is not None else None
            let_local = locals_.get(let_lower) if let_lower else None
            if let_lower and let_local is not None and let_local.variant:
                touched.add(let_lower)
        return touched

    def demote_to_unknown(lower: str) -> None:
        if state.get(lower) == "unset":
            state[lower] = "unknown"

    def restore_state(snapshot: Mapping[str, str]) -> None:
        state.clear()
        state.update(snapshot)

    def set_state(key: str, value: str) -> None:
        state[key] = value

    # A local never Set is Nothing: `If c Is Nothing Then Exit Function`
    # always leaves (issue #273). 'set' proves nothing, since a Set from
    # a call may store Nothing.
    def is_nothing(lower: str) -> bool | None:
        local = locals_.get(lower)
        return (
            True
            if local is not None and not local.let_only and state.get(lower) == "unset"
            else None
        )

    def known_condition(condition: Sequence[VbaToken]) -> bool | None:
        value: bool | None = condition_value(
            condition, ConditionFacts(value=lambda lower: None, is_nothing=is_nothing)
        )
        return value

    def set_silent(quiet: bool) -> None:
        nonlocal silent
        silent = quiet

    hooks = StraightLineDataflowHooks(
        on_statement=on_statement,
        on_block=on_block,
        after_block=after_block,
        touches_in_statement=touches_in_statement,
        demote_to_unknown=demote_to_unknown,
        snapshot_state=lambda: dict(state),
        restore_state=restore_state,
        set_state=set_state,
        lattice=Lattice(init="unset", good="set", unknown="unknown"),
        known_condition=known_condition,
        set_silent=set_silent,
    )
    walk(
        source,
        member.body,
        lambda node: is_inactive_node(activity, node) or id(node) in unreachable,
        hooks,
    )


def _check_go_to_into_with(
    source: str,
    proc: ProcedureNode,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """`GoTo L` from outside a With block to a label inside it skips the With
    statement, so the With has no object: the first leading-dot member after the
    label raises 91 (issue #184, measured in Excel 16.0). A GoTo inside the same
    With runs, and so does one into a For loop."""
    if not any(isinstance(node, WithBlockNode) for node in iter_body_nodes(proc.body)):
        return
    labels: dict[str, tuple[tuple[int, ...], str | None]] = {}
    jumps: list[tuple[str, str, Span, tuple[int, ...]]] = []
    # Each frame: the list, the next index in it, and the identities of the
    # With blocks around it, outermost first.
    stack: list[tuple[Sequence[BodyNode], list[int], tuple[int, ...]]] = [(proc.body, [0], ())]
    while stack:
        body, cursor, withs = stack[-1]
        if cursor[0] >= len(body):
            stack.pop()
            continue
        i = cursor[0]
        cursor[0] += 1
        node = body[i]
        if is_inactive_node(activity, node):
            continue
        if is_leaf_statement(node):
            for label in statement_label_declarations(source, node.span) if withs else []:
                if label.key not in labels:
                    labels[label.key] = (
                        withs,
                        _first_leading_dot_member(source, body, i, activity),
                    )
            for ref in statement_label_references(source, node.span):
                if ref.statement_kind == "goto":
                    jumps.append((ref.key, ref.text, ref.span, withs))
            continue
        child = getattr(node, "body", None)
        if isinstance(child, list):
            stack.append(
                (child, [0], (*withs, id(node)) if isinstance(node, WithBlockNode) else withs)
            )
    for key, text, span, jump_withs in jumps:
        target = labels.get(key)
        if target is not None and target[1] and any(block not in jump_withs for block in target[0]):
            push(
                "objectVariableNotSet",
                f"GoTo {text} jumps into a With block past its With statement, so '{target[1]}' after the "
                "label has no object. This will raise Run-time error '91': Object variable or With block "
                "variable not set.",
                span,
            )


def _first_leading_dot_member(
    source: str,
    body: Sequence[BodyNode],
    from_: int,
    activity: ConditionalActivityTracker | None,
) -> str | None:
    """The first leading-dot member (`.Add`) that runs from `body[from_]` on, in
    the statements that follow in a straight line. A block may not run, and an
    Exit, GoTo or Return leaves, so either ends the search."""
    for j in range(from_, len(body)):
        node = body[j]
        if is_inactive_node(activity, node) or isinstance(node, VariableGroupNode):
            continue
        if not is_leaf_statement(node):
            return None
        toks = statement_tokens_after_leading_label(source, node.span)
        head = token_text(_at(toks, 0))
        if j > from_ and head in ("elseif", "else", "case"):
            return None
        # A one-line If always runs its condition, and its branches maybe.
        then = (
            next((k for k, tok in enumerate(toks) if token_text(tok) == "then"), -1)
            if isinstance(node, StatementNode) and node.single_line_if_branches is not None
            else -1
        )
        limit = then if then >= 0 else len(toks)
        for k in range(limit):
            if (
                toks[k].raw_text == "."
                and token_name(_at(toks, k + 1))
                and (k == 0 or precedes_leading_member_dot(toks[k - 1]))
            ):
                return f".{toks[k + 1].raw_text}"
        if then >= 0 or head in ("exit", "goto", "return", "resume", "end"):
            return None
    return None


def _body_can_leave_loop(
    source: str, loop: ForBlockNode, activity: ConditionalActivityTracker | None
) -> bool:
    """Whether the loop body can leave the loop before it ends: an `Exit For` at its
    own depth (one inside a nested For leaves that one), or any `GoTo`."""

    def skip(node: BodyNode) -> bool:
        # A nested For's Exit For is its own.
        return is_inactive_node(activity, node) or isinstance(node, ForBlockNode)

    for node in iter_body_nodes(loop.body, skip):
        if isinstance(getattr(node, "body", None), list):
            continue
        for span in statement_and_branch_spans(node):  # type: ignore[arg-type]
            toks = statement_tokens_after_leading_label(source, span)
            head = token_text(_at(toks, 0))
            if (head == "exit" and token_text(_at(toks, 1)) == "for") or head == "goto":
                return True
    return False


def _single_line_if_arms(
    toks: Sequence[VbaToken], then_index: int, span: Span
) -> tuple[Span, Span | None]:
    """The Then and Else arms of a single-line If, as offsets. A nested one-line If
    sits inside the outer Then arm, and an Else belongs to the innermost If still
    open: in `If A Then If B Then X Else Y`, Y is B's (issue #575)."""
    start = span.start + toks[then_index].end
    open_ = 0
    for i in range(then_index + 1, len(toks)):
        word = token_text(toks[i])
        if word == "if":
            open_ += 1
        elif word == "else":
            if open_ == 0:
                return (
                    Span(start, span.start + toks[i].start),
                    Span(span.start + toks[i].end, span.end),
                )
            open_ -= 1
    return Span(start, span.end), None


def _nothing_guard_names(condition: Sequence[VbaToken]) -> tuple[set[str], set[str]]:
    """The tracked names a single-line If's condition guards, as (Then arm, Else
    arm): `Not d Is Nothing` guards the Then arm, `d Is Nothing` the Else arm
    (issue #108: the block form already read the guard, the one-line form did not)."""
    then_arm: set[str] = set()
    else_arm: set[str] = set()
    for i in range(len(condition) - 2):
        if token_text(condition[i + 1]) != "is" or token_text(condition[i + 2]) != "nothing":
            continue
        name = token_name(condition[i])
        if not name:
            continue
        if token_text(_at(condition, i - 1)) == "not":
            then_arm.add(name.lower())
        else:
            else_arm.add(name.lower())
    return then_arm, else_arm


def _check_object_variable_not_set_statement(
    source: str,
    stmt: LeafStatementNode,
    locals_: Mapping[str, _LocalObjectVariable],
    state: dict[str, ObjectVariableState],
    set_anywhere: AbstractSet[str],
    member_ctx: MemberCompletionContext,
    push: PushFn,
    lets: dict[int, ObjectVariableState],
    facts: _ModuleObjectFacts,
    elements: _ObjectArrayElements | None = None,
) -> None:
    if elements is None:
        elements = _ObjectArrayElements()
    toks = statement_tokens_after_leading_label(source, stmt.span)
    head = token_text(_at(toks, 0))
    # `GoSub Label` runs the subroutine, which may Set any of the locals,
    # before the statement after it (issue #108).
    if head == "gosub" or (head == "on" and any(token_text(tok) == "gosub" for tok in toks)):
        for key in [*set_anywhere, *elements.keys.keys()]:
            if state.get(key) == "unset":
                state[key] = "unknown"
        return
    lower: str | None
    # `a(0).Count` on an element never set (issue #489, measured in Excel 16.0).
    if elements.keys:
        for text, start, end in _unset_element_accesses(toks, elements, state):
            array = elements.arrays.get(text[: text.find("(")].lower())
            push(
                "objectVariableNotSet",
                f"Element {text} of '{array[0] if array is not None else text}' is Nothing before member "
                "access. This will raise Run-time error '91': Object variable or With block variable not set.",
                Span(stmt.span.start + start, stmt.span.start + end),
            )
        _update_elements(toks, elements, state)
    # The arms of a single-line If and what its condition proves about them.
    branches = statement_and_branch_spans(stmt)
    then_guards: set[str] = set()
    else_guards: set[str] = set()
    then_arm: Span | None = None
    else_arm: Span | None = None
    if head == "if" and len(branches) > 1:
        then_index = next(
            (index for index, tok in enumerate(toks) if index > 0 and token_text(tok) == "then"), -1
        )
        if then_index > 0:
            then_guards, else_guards = _nothing_guard_names(toks[1:then_index])
            then_arm, else_arm = _single_line_if_arms(toks, then_index, stmt.span)

    def guarded_at(name: str, offset: int) -> bool:
        def within(span: Span | None) -> bool:
            return span is not None and span.start <= offset < span.end

        return (name in then_guards and within(then_arm)) or (
            name in else_guards and within(else_arm)
        )

    # A bare `obj = value` is a Let through the object's default member
    # (issue #107), which needs an object to reach: on a variable still
    # Nothing it raises 91, the same as a member access would.
    for span in branches:
        let_target = bare_assignment_target(source, span)
        lower = let_target[0].lower() if let_target is not None else None
        let_local = locals_.get(lower) if lower else None
        if let_target is None or not lower or let_local is None or let_local.variant:
            continue
        let_state = (
            "unknown" if guarded_at(lower, let_target[1].start) else state.get(lower, "unknown")
        )
        lets[let_target[1].start] = let_state
        # A type with no default member for the Let, or one that needs an
        # argument, is set-required's to report, with the 91 when it is still
        # Nothing (issue #193): the fix there is the Set.
        verdict = object_let_assignment_verdict(let_local.as_type, member_ctx)
        if (
            let_state == "unset"
            and verdict != "noDefault"
            and verdict != "argument"
            and not object_holding_default(let_local.as_type, member_ctx)
            and not read_only_host_default(let_local.as_type, member_ctx)
        ):
            what = (
                f"The result '{let_target[0]}'"
                if let_local.let_only
                else f"Object variable '{let_target[0]}'"
            )
            push(
                "objectVariableNotSet",
                f"{what} is Nothing before the default-member assignment. This will raise Run-time error "
                "'91': Object variable or With block variable not set.",
                let_target[1],
            )
    # `If c Then` reads c's value for the condition: on c still Nothing that
    # raises 91 whatever its type's default member (issue #268, measured in
    # Excel 16.0 on a Collection). A type with no default member is
    # object-default-value's, 438 or 91 (issue #415).
    # So do a loop's condition, Select Case, IIf, Not and And (issue #424).
    for operand in condition_operands(toks):
        index: int = operand.index
        form: str = operand.form
        lower = toks[index].raw_text.lower()
        local = locals_.get(lower)
        at = Span(stmt.span.start + toks[index].start, stmt.span.start + toks[index].end)
        if (
            local is not None
            and not local.let_only
            and not local.variant
            and state.get(lower) == "unset"
            and not guarded_at(lower, stmt.span.start + toks[index].start)
            and object_let_assignment_verdict(local.as_type, member_ctx) != "noDefault"
        ):
            if form == "condition":
                reads = "the condition reads"
            elif form == "select":
                reads = "Select Case reads"
            elif form == "iif":
                reads = "IIf reads"
            else:
                reads = f"'{'Not' if form == 'not' else 'the Boolean operator'}' reads"
            push(
                "objectVariableNotSet",
                f"Object variable '{toks[index].raw_text}' is Nothing when {reads} its value. This will raise "
                "Run-time error '91': Object variable or With block variable not set.",
                at,
            )
        elif (
            local is not None
            and not local.variant
            and form in ("condition", "iif")
            and state.get(lower) == "set"
            and normalize_type(local.as_type) != "collection"
            and object_value_needs_index(local.as_type, member_ctx)
        ):
            # Set, a Word Paragraphs or Tables has an Item that needs an index,
            # as a Collection's does (issue #438, measured in Word 16.0). A
            # Collection is condition-values'.
            article = "an" if re.match(r"[aeiou]", local.as_type or "", re.IGNORECASE) else "a"
            push(
                "objectDefaultValue",
                f"'{toks[index].raw_text}' is {article} {local.as_type}: its default member Item needs an "
                f"index, so {'IIf' if form == 'iif' else 'the condition'} has no value to read. This will "
                "raise Run-time error '450': Wrong number of arguments or invalid property assignment.",
                at,
            )
    # `x = c` reads c's default member, which needs an object: on c still
    # Nothing it raises 91 (issue #256, measured in Excel 16.0). A type with
    # no default member, or one that needs an argument, is object-default-value's.
    for span in branches:
        target = bare_assignment_target(source, span)
        value = (
            [tok for tok in target[2] if tok.kind is not TokenKind.COMMENT]
            if target is not None
            else []
        )
        value_name = token_name(value[0]) if len(value) == 1 else None
        lower = value_name.lower() if value_name is not None else None
        local = locals_.get(lower) if lower else None
        if (
            target is None
            or lower is None
            or local is None
            or local.let_only
            or local.variant
            or target[0].lower() in locals_
            or state.get(lower) != "unset"
            or guarded_at(lower, span.start + value[0].start)
            or object_let_assignment_verdict(local.as_type, member_ctx) != "lets"
        ):
            continue
        push(
            "objectVariableNotSet",
            f"Object variable '{value[0].raw_text}' is Nothing when its default member is read. This will "
            "raise Run-time error '91': Object variable or With block variable not set.",
            Span(span.start + value[0].start, span.start + value[0].end),
        )
    # `x + 1`, `x & "a"` read x's default member as an operand: on x still
    # Nothing that raises 91 (issue #462, measured in Word and PowerPoint
    # 16.0 on a Range, the Selection and a TextRange).
    host = _host_name(member_ctx)
    for span in branches:
        operand_toks = statement_tokens(source, span)
        # A Set's `=` is no operator: `Set x = y` reads neither value. Its
        # value indexed, `Set p = o(1)`, calls o's default member, which
        # needs o (issue #296, measured in Excel 16.0: 91).
        if (
            token_text(_at(operand_toks, 0)) == "set"
            or set_assignment_target(source, span) is not None
        ):
            eq = next((k for k, tok in enumerate(operand_toks) if tok.raw_text == "="), -1)
            set_value = _at(operand_toks, eq + 1)
            set_name = token_name(set_value)
            lower = set_name.lower() if set_name is not None else None
            local = locals_.get(lower) if lower else None
            if (
                eq > 0
                and set_value is not None
                and lower is not None
                and local is not None
                and not local.let_only
                and not local.variant
                and _raw(_at(operand_toks, eq + 2)) == "("
                and state.get(lower) == "unset"
                and not guarded_at(lower, span.start + set_value.start)
                and match_paren_from(operand_toks, eq + 2) == len(operand_toks) - 1
                and object_let_assignment_verdict(local.as_type, member_ctx) != "noDefault"
            ):
                push(
                    "objectVariableNotSet",
                    f"Object variable '{set_value.raw_text}' is Nothing when its default member is indexed. "
                    "This will raise Run-time error '91': Object variable or With block variable not set.",
                    Span(span.start + set_value.start, span.start + set_value.end),
                )
            continue
        bare_target = bare_assignment_target(source, span)
        eq = (
            next((k for k, tok in enumerate(operand_toks) if tok.raw_text == "="), -1)
            if bare_target is not None
            else -1
        )
        then = (
            next((k for k, tok in enumerate(operand_toks) if token_text(tok) == "then"), -1)
            if span is stmt.span and len(branches) > 1
            else -1
        )
        limit = then if then > 0 else len(operand_toks)

        def is_operator(
            index: int, eq: int = eq, operand_toks: Sequence[VbaToken] = operand_toks
        ) -> bool:
            tok = _at(operand_toks, index)
            return (
                index != eq
                and tok is not None
                and (
                    (tok.kind is TokenKind.OPERATOR and tok.raw_text in _OPERAND_OPERATORS)
                    or token_text(tok) == "mod"
                )
            )

        for i in range(limit):
            name = token_name(operand_toks[i])
            lower = name.lower() if name is not None else None
            local = locals_.get(lower) if lower else None
            if (
                lower is None
                or local is None
                or local.let_only
                or local.variant
                or i == eq - 1
                or _raw(_at(operand_toks, i - 1)) == "."
                or _raw(_at(operand_toks, i + 1)) == "."
                or state.get(lower) != "unset"
                or guarded_at(lower, span.start + operand_toks[i].start)
                or object_holding_default(local.as_type, member_ctx)
            ):
                continue
            verdict = object_let_assignment_verdict(local.as_type, member_ctx)
            # `CStr(x)`, `Len(x)`: a whole argument of a built-in that reads one
            # value (issue #415, measured in Excel 16.0). A Collection or Names
            # there does not compile, which is collection-operand's; a type with
            # no default member is object-default-value's, 438 or 91.
            argument = (
                (_raw(_at(operand_toks, i - 1)) or "") in ("(", ",")
                and (_raw(_at(operand_toks, i + 1)) or "") in (")", ",")
                and token_text(_at(operand_toks, builtin_name_before(operand_toks, i)))
                in ONE_VALUE_BUILTINS
                and verdict == "lets"
            )
            # `x(1)` passes the index to the default member: an Excel Range's
            # takes one, an Object is late bound, and a Collection's needs one.
            # A default that takes none, as Application's Name, does not
            # compile; a type with none is object-default-value's.
            type_ = normalize_type(local.as_type)
            close = (
                match_paren_from(operand_toks, i + 1)
                if _raw(_at(operand_toks, i + 1)) == "("
                else -1
            )
            indexed = (
                close > i + 2
                and _raw(_at(operand_toks, close + 1)) != "."
                and (
                    verdict == "argument"
                    or type_ == "object"
                    or (type_ == "range" and host != "Word" and host != "PowerPoint")
                )
            )
            operand_read = (
                _raw(_at(operand_toks, i + 1)) != "("
                and (is_operator(i - 1) or is_operator(i + 1))
                and verdict == "lets"
            )
            if not argument and not indexed and not operand_read:
                continue
            how = (
                "default member is indexed"
                if indexed
                else "value is read"
                if argument
                else "default member is read as an operand"
            )
            push(
                "objectVariableNotSet",
                f"Object variable '{operand_toks[i].raw_text}' is Nothing when its {how}. This will raise "
                "Run-time error '91': Object variable or With block variable not set.",
                Span(span.start + operand_toks[i].start, span.start + operand_toks[i].end),
            )
    passed_whole: Mapping[str, int] = locals_named_whole(
        source, stmt.span, locals_, _OBJECT_READ_ONLY_INTRINSICS
    )
    for name, span in _unset_object_member_accesses(source, stmt.span, locals_, state, member_ctx):
        # An access after a whole pass in the same statement, as in
        # `If TryGet(obj) Then obj.Name`, runs after the callee had its chance
        # to Set it. One before the pass, as in `Load(obj.Name)`, does not.
        pass_at = passed_whole.get(name.lower())
        if pass_at is not None and span.start > pass_at:
            continue
        if guarded_at(name.lower(), span.start):
            continue
        push(
            "objectVariableNotSet",
            f"Object variable '{name}' is Nothing before member access. This will raise Run-time error "
            "'91': Object variable or With block variable not set.",
            span,
        )
    # Nothing passed to a procedure of the module that reads a member of the
    # parameter first raises 91 there (issue #343, measured in Excel 16.0).
    for span in branches:
        for name, callee, read, start, end in _nothing_passed_to_member_read(
            statement_tokens(source, span), locals_, state, facts.member_first
        ):
            if guarded_at(name.lower(), span.start + start):
                continue
            push(
                "objectVariableNotSet",
                f"Object variable '{name}' is Nothing, and {callee} reads '{read}' from it first. This will "
                "raise Run-time error '91': Object variable or With block variable not set.",
                Span(span.start + start, span.start + end),
            )
    # Excel's Union takes no Nothing, first argument or any other (issue
    # #680, measured in Excel 16.0).
    if "union" in facts.excel_range_methods:
        for span in branches:
            for name, start, end in _nothing_passed_to_union(
                statement_tokens(source, span), locals_, state
            ):
                if not guarded_at(name.lower(), span.start + start):
                    push(
                        "objectVariableNotSet",
                        f"Object variable '{name}' is Nothing, and Union takes no Nothing. This will raise "
                        "Run-time error '5': Invalid procedure call or argument.",
                        Span(span.start + start, span.start + end),
                    )
    set_target = set_assignment_target(source, stmt.span)
    if set_target is not None:
        lower = set_target[0].lower()
        into = locals_.get(lower)
        if into is not None:
            state[lower] = _set_value_state(set_target[2], into, locals_, state, facts)
            return
    # A Let gives a Variant a value that is no object.
    let_assignment = bare_assignment_target(source, stmt.span)
    let_lower = let_assignment[0].lower() if let_assignment is not None else None
    let_variant = locals_.get(let_lower) if let_lower else None
    if let_lower and let_variant is not None and let_variant.variant:
        state[let_lower] = "unknown"
    for lower in passed_whole:
        if state.get(lower) == "unset":
            state[lower] = "unknown"
    # A Set in a single-line If's branch runs on one path only, so it moves an
    # unset object to 'unknown' the way a block If without Else does, not to
    # 'set' - as unallocated-dynamic-array-access reads a conditional ReDim.
    for branch in statement_and_branch_spans(stmt)[1:]:
        branch_target = set_assignment_target(source, branch)
        branch_lower = branch_target[0].lower() if branch_target is not None else None
        if branch_lower and branch_lower in locals_ and state.get(branch_lower) == "unset":
            state[branch_lower] = "unknown"
        branch_let = bare_assignment_target(source, branch)
        branch_let_lower = branch_let[0].lower() if branch_let is not None else None
        branch_variant = locals_.get(branch_let_lower) if branch_let_lower else None
        if (
            branch_let_lower
            and branch_variant is not None
            and branch_variant.variant
            and state.get(branch_let_lower) == "unset"
        ):
            state[branch_let_lower] = "unknown"


def _late_bound(type_: str) -> bool:
    normalized = normalize_type(type_)
    return (normalized if normalized is not None else "variant") in ("object", "variant")


def _set_value_state(
    value_tokens: Sequence[VbaToken],
    into: _LocalObjectVariable,
    locals_: Mapping[str, _LocalObjectVariable],
    state: Mapping[str, ObjectVariableState],
    facts: _ModuleObjectFacts,
) -> ObjectVariableState:
    """What a Set leaves in its target: Nothing from `Nothing`, from a local still
    Nothing, from a Function of the module that returns Nothing (issue #343), or
    from an Intersect of literal ranges that do not meet (issue #680); a local's
    own state from a local; otherwise an object."""
    if _set_assignment_value_is_nothing(value_tokens) or (
        "intersect" in facts.excel_range_methods and literal_intersect_is_nothing(value_tokens)
    ):
        return "unset"
    toks = [
        tok
        for tok in value_tokens
        if tok.kind is not TokenKind.COMMENT and tok.kind is not TokenKind.NEWLINE
    ]
    lower = (token_name(_at(toks, 0)) or "").lower()
    local = locals_.get(lower)
    if len(toks) == 1 and local is not None and not local.let_only:
        # A late-bound copy of a typed variable is runtime-member-not-found's,
        # which names the 91 with the 438 its members raise.
        copied = state.get(lower, "unknown")
        return (
            "unknown"
            if copied == "unset" and _late_bound(into.as_type) and not _late_bound(local.as_type)
            else copied
        )
    fn = facts.nothing_functions.get(lower)
    if len(toks) == 1:
        called = fn is not None and len(fn.params) == 0
    else:
        called = _raw(_at(toks, 1)) == "(" and match_paren_from(toks, 1) == len(toks) - 1
    return "unset" if fn is not None and called else "set"


def _nothing_passed_to_member_read(
    toks: Sequence[VbaToken],
    locals_: Mapping[str, _LocalObjectVariable],
    state: Mapping[str, ObjectVariableState],
    member_first: Mapping[str, Mapping[int, str]],
) -> list[tuple[str, str, str, int, int]]:
    """The locals still Nothing that a statement passes whole to a procedure of
    the module whose parameter there is read with a member first: `TakeC(o)`,
    `TakeC o`, `Call TakeC(o)`. Offsets are the statement's; each hit is
    (name, callee, read, start, end)."""
    out: list[tuple[str, str, str, int, int]] = []
    code = [tok for tok in toks if tok.kind is not TokenKind.COMMENT]
    for i in range(len(code)):
        reads = member_first.get((token_name(code[i]) or "").lower())
        before = _raw(_at(code, i - 1))
        if reads is None or before == "." or before == "!":
            continue
        # Parenthesized anywhere, or bare as the statement's first word.
        paren = match_paren_from(code, i + 1) if _raw(_at(code, i + 1)) == "(" else -1
        bare = paren < 0 and (i == 0 or (i == 1 and token_text(code[0]) == "call"))
        if paren < 0 and not bare:
            continue
        args = (
            split_top_level_token_groups(code, i + 2, ",", paren)
            if paren > 0
            else split_top_level_token_groups(code, i + 1, ",", len(code))
        )
        for k, read in reads.items():
            arg = args[k] if k < len(args) else None
            arg_name = token_name(arg[0]) if arg is not None and len(arg) == 1 else None
            lower = arg_name.lower() if arg_name is not None else None
            local = locals_.get(lower) if lower else None
            if (
                arg is None
                or lower is None
                or local is None
                or local.let_only
                or state.get(lower) != "unset"
            ):
                continue
            out.append((arg[0].raw_text, code[i].raw_text, read, arg[0].start, arg[0].end))
    return out


@dataclass(slots=True)
class _ObjectArrayElements:
    """The object arrays of a procedure and the elements its code names by a literal index."""

    # By lowercased name: (the array as declared, whether its bounds are fixed).
    arrays: dict[str, tuple[str, bool]] = field(default_factory=dict)
    # By state key, `a(0)`: (the array, the index).
    keys: dict[str, tuple[str, float]] = field(default_factory=dict)


def _iter_with_headers(source: str, body: Sequence[BodyNode]) -> Iterator[Span]:
    """The header line of every With block, an If's arms walked by branch."""
    stack: list[Iterator[BodyNode]] = [iter(body)]
    while stack:
        for node in stack[-1]:
            if isinstance(node, WithBlockNode):
                yield block_header_line_span(source, node.span)
            if isinstance(node, IfBlockNode):
                stack.append(itertools.chain.from_iterable(branch.body for branch in node.branches))
                break
            child = getattr(node, "body", None)
            if isinstance(child, list):
                stack.append(iter(child))
                break
        else:
            stack.pop()


def _object_array_elements(
    source: str,
    symbols: ModuleSymbols,
    proc: ProcedureNode,
    member_ctx: MemberCompletionContext,
    activity: ConditionalActivityTracker | None,
) -> _ObjectArrayElements:
    """The local arrays of an object type, not `As New`, and every `a(n)` the
    procedure writes with a whole-number literal n (issue #489, measured in
    Excel 16.0). Only those elements are followed."""
    elements = _ObjectArrayElements()
    proc_sym = procedure_symbol_for(symbols, proc)
    for child in (proc_sym.children if proc_sym is not None else None) or []:
        element_type = (
            _ARRAY_RETURN_RE.sub("", child.as_type, count=1) if child.as_type is not None else None
        )
        if (
            child.kind is VbaSymbolKind.LOCAL_VARIABLE
            and child.is_array
            and child.visibility is not SymbolVisibility.STATIC
            and not child.is_auto_instantiated
            and element_type
            and is_known_object_assignment_type(element_type, member_ctx)
        ):
            elements.arrays[child.name.lower()] = (child.name, child.array_bounds is not None)
    if not elements.arrays:
        return elements
    # `With a(0)` names an element too (issue #295).
    spans: list[Span] = list(_iter_with_headers(source, proc.body))

    def collect(stmt: LeafStatementNode) -> None:
        spans.append(stmt.span)

    for_each_statement(proc.body, collect, activity)
    for span in spans:
        toks = statement_tokens(source, span)
        i = 0
        while i + 3 < len(toks):
            name = token_name(toks[i])
            lower = name.lower() if name is not None else None
            if (
                lower
                and lower in elements.arrays
                and _raw(_at(toks, i - 1)) != "."
                and toks[i + 1].raw_text == "("
                and toks[i + 2].kind is TokenKind.INTEGER_LITERAL
                and toks[i + 3].raw_text == ")"
            ):
                index = js_number(_INTEGER_SUFFIX_RE.sub("", toks[i + 2].raw_text, count=1))
                if (
                    index == index
                    and index not in (float("inf"), float("-inf"))
                    and index == int(index)
                ):
                    elements.keys[f"{lower}({js_number_to_string(index)})"] = (lower, index)
            i += 1
    return elements


def _check_for_each_over_empty_object_array(
    source: str,
    member: ProcedureNode,
    symbols: ModuleSymbols,
    member_ctx: MemberCompletionContext,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """`For Each x In a` over a fixed array of objects the procedure never fills:
    x is Nothing on the first pass, so the body's first use of it through a
    member raises 91 (issue #489, measured in Excel 16.0)."""
    arrays = _object_array_elements(source, symbols, member, member_ctx, activity).arrays
    empty = {lower for lower, array in arrays.items() if array[1]}
    if not empty:
        return

    def visit_writes(stmt: LeafStatementNode) -> None:
        for span in statement_and_branch_spans(stmt):
            toks = statement_tokens_after_leading_label(source, span)
            for lower in list(empty):
                # Any write of an element, or a whole mention, may fill it.
                named = any(
                    (token_name(tok) or "").lower() == lower
                    and token_name(tok) is not None
                    and _raw(_at(toks, i - 1)) != "."
                    and (
                        _raw(_at(toks, i + 1)) != "("
                        or token_text(_at(toks, i - 1)) == "set"
                        or token_text(_at(toks, 0)) in ("redim", "erase")
                    )
                    for i, tok in enumerate(toks)
                )
                if named:
                    empty.discard(lower)

    for_each_statement(member.body, visit_writes, activity)
    for node in iter_body_nodes(member.body, lambda node: is_inactive_node(activity, node)):
        child_body = getattr(node, "body", None)
        if not isinstance(child_body, list):
            continue
        over = (
            node.source_expression.strip().lower()
            if isinstance(node, ForBlockNode) and node.each and node.source_expression
            else None
        )
        control = (
            node.control_variable.lower()
            if isinstance(node, ForBlockNode) and node.control_variable
            else None
        )
        if over and control and over in empty:
            for stmt in child_body:
                if is_inactive_node(activity, stmt):
                    continue
                if not is_leaf_statement(stmt) or len(statement_and_branch_spans(stmt)) > 1:
                    break
                toks = statement_tokens_after_leading_label(source, stmt.span)
                at = next(
                    (
                        i
                        for i, tok in enumerate(toks)
                        if (token_name(tok) or "").lower() == control
                        and token_name(tok) is not None
                        and _raw(_at(toks, i - 1)) != "."
                    ),
                    -1,
                )
                if at < 0:
                    continue
                if (
                    token_text(_at(toks, 0)) != "set"
                    and _raw(_at(toks, at + 1)) == "."
                    and token_name(_at(toks, at + 2))
                ):
                    push(
                        "objectVariableNotSet",
                        f"'{toks[at].raw_text}' takes each element of '{arrays[over][0]}', which the code never "
                        "sets, so it is Nothing on the first pass. This will raise Run-time error '91': Object "
                        "variable or With block variable not set.",
                        Span(stmt.span.start + toks[at].start, stmt.span.start + toks[at].end),
                    )
                break


def _element_keys_of(elements: _ObjectArrayElements, array: str) -> list[str]:
    """The element keys of an array, `a(0)` and the rest."""
    return [key for key, element in elements.keys.items() if element[0] == array]


def _names(tok: VbaToken | None, lower: str) -> bool:
    name = token_name(tok)
    return name is not None and name.lower() == lower


def _update_elements(
    toks: Sequence[VbaToken], elements: _ObjectArrayElements, state: dict[str, ObjectVariableState]
) -> None:
    """What a statement does to the followed elements, after its reads: `Set a(0)
    = ...` sets one, `Set a(i) = ...` may set any, a plain ReDim and Erase leave
    every one Nothing, ReDim Preserve keeps them, and any other whole mention of
    the array may change them."""
    head = token_text(_at(toks, 0))
    for array in elements.arrays:
        keys = _element_keys_of(elements, array)
        if not keys or not any(_names(tok, array) for tok in toks):
            continue
        if head == "set" and _names(_at(toks, 1), array) and _raw(_at(toks, 2)) == "(":
            close = match_paren_from(toks, 2)
            literal = (
                f"{array}({_literal_index_text(toks[3].raw_text)})"
                if close == 4 and toks[3].kind is TokenKind.INTEGER_LITERAL
                else None
            )
            value = [tok for tok in toks[close + 2 :] if tok.kind is not TokenKind.COMMENT]
            nothing = len(value) == 1 and token_text(value[0]) == "nothing"
            if literal is not None and literal in state:
                state[literal] = "unset" if nothing else "set"
            elif literal is None:
                for key in keys:
                    state[key] = "unknown"
            continue
        # A Set in a one-line If's branch may run or not.
        if head != "set" and any(
            token_text(tok) == "set" and _names(_at(toks, i + 1), array)
            for i, tok in enumerate(toks)
        ):
            for key in keys:
                if state.get(key) == "unset":
                    state[key] = "unknown"
            continue
        if head in ("redim", "erase"):
            if not (head == "redim" and token_text(_at(toks, 1)) == "preserve"):
                for key in keys:
                    state[key] = "unset"
            continue
        # `Fill a`, `b = a`: the whole array, which the callee or a copy may change.
        whole = any(
            _names(tok, array) and _raw(_at(toks, i - 1)) != "." and _raw(_at(toks, i + 1)) != "("
            for i, tok in enumerate(toks)
        )
        if whole:
            for key in keys:
                state[key] = "unknown"


def _element_touches(toks: Sequence[VbaToken], elements: _ObjectArrayElements) -> list[str]:
    """The followed elements a statement may change, for a block that runs it or not."""
    out: list[str] = []
    head = token_text(_at(toks, 0))
    for array in elements.arrays:
        if any(
            _names(tok, array)
            and _raw(_at(toks, i - 1)) != "."
            and (_raw(_at(toks, i + 1)) != "(" or head == "set")
            for i, tok in enumerate(toks)
        ) or (head in ("redim", "erase") and any(_names(tok, array) for tok in toks)):
            out.extend(_element_keys_of(elements, array))
    return out


def _unset_element_accesses(
    toks: Sequence[VbaToken],
    elements: _ObjectArrayElements,
    state: Mapping[str, ObjectVariableState],
) -> list[tuple[str, int, int]]:
    """`a(0).Count` with a(0) still Nothing, as (text, start, end). Offsets are the statement's."""
    out: list[tuple[str, int, int]] = []
    i = 0
    while i + 5 < len(toks):
        name = token_name(toks[i])
        lower = name.lower() if name is not None else None
        if (
            not lower
            or lower not in elements.arrays
            or _raw(_at(toks, i - 1)) == "."
            or toks[i + 1].raw_text != "("
            or toks[i + 2].kind is not TokenKind.INTEGER_LITERAL
            or toks[i + 3].raw_text != ")"
            or toks[i + 4].raw_text != "."
            or not token_name(toks[i + 5])
        ):
            i += 1
            continue
        key = f"{lower}({_literal_index_text(toks[i + 2].raw_text)})"
        if state.get(key) == "unset":
            out.append(
                ("".join(tok.raw_text for tok in toks[i : i + 4]), toks[i].start, toks[i + 3].end)
            )
        i += 1
    return out


# Intrinsics that read an object argument and never Set it.
_OBJECT_READ_ONLY_INTRINSICS: frozenset[str] = frozenset(
    {"typename", "vartype", "isobject", "isnull", "isempty", "ismissing", "objptr"}
)


def _local_object_variables_for(
    source: str,
    symbols: ModuleSymbols,
    proc: ProcedureNode,
    member_ctx: MemberCompletionContext,
) -> dict[str, _LocalObjectVariable]:
    """The tracked object variables, by lowercased name."""
    out: dict[str, _LocalObjectVariable] = {}
    proc_sym = procedure_symbol_for(symbols, proc)
    children = (proc_sym.children if proc_sym is not None else None) or []
    for child in children:
        # `DefObj O` then `Dim o` is an Object (issue #285).
        as_type: str | None = (
            child.as_type
            if child.as_type is not None
            else def_type_of(symbols, child.name)
            if child.kind is VbaSymbolKind.LOCAL_VARIABLE
            else None
        )
        if (
            child.kind is not VbaSymbolKind.LOCAL_VARIABLE
            or child.visibility is SymbolVisibility.STATIC
            or child.is_array is True
            # `Dim x As New Invoice` is instantiated on ANY access, including
            # the first one and including after `Set x = Nothing`, so it can
            # never be Nothing when a member is touched. Tracking it produced
            # error 91 warnings on code that runs.
            or child.is_auto_instantiated is True
            or not is_known_object_assignment_type(as_type, member_ctx)
            or not as_type
        ):
            continue
        out[child.name.lower()] = _LocalObjectVariable(name=child.name, as_type=as_type)
    # `DefObj A-Z` then `o = 5` with nothing declaring o: an Object local,
    # Nothing until a Set (issue #685, measured in Excel 16.0).
    implicit_locals: Mapping[int, AbstractSet[str]] | None = symbols.implicit_locals
    implicit = implicit_locals.get(proc.span.start) if implicit_locals is not None else None
    if implicit:
        env = type_environment_for(symbols, proc)
        for lower in implicit:
            if lower not in out and normalize_type(env.get(lower)) == "object":
                out[lower] = _LocalObjectVariable(name=lower, as_type="Object")
    # A Variant is followed only where the procedure sets it to Nothing.
    text = source[proc.span.start : proc.span.end]
    for child in children:
        type_ = normalize_type(child.as_type)
        if (
            child.kind is VbaSymbolKind.LOCAL_VARIABLE
            and child.visibility is not SymbolVisibility.STATIC
            and not child.is_array
            and (type_ is None or type_ == "variant")
            and _set_nothing_re(child.name).search(text)
        ):
            out[child.name.lower()] = _LocalObjectVariable(
                name=child.name, as_type="Variant", variant=True
            )
    # A module's object variable, where the procedure sets it to Nothing and
    # no local or parameter hides it (issue #618, measured in Excel 16.0).
    hidden = {child.name.lower() for child in children} | {proc.name.lower()}
    for child in symbols.root.children or []:
        if (
            child.kind is VbaSymbolKind.MODULE_VARIABLE
            and not child.is_array
            and not child.is_auto_instantiated
            and child.name.lower() not in hidden
            and child.as_type
            and is_known_object_assignment_type(child.as_type, member_ctx)
            and _set_nothing_re(child.name).search(text)
        ):
            out[child.name.lower()] = _LocalObjectVariable(
                name=child.name, as_type=child.as_type, module=True
            )
    result: str | None = return_assignment_type_for(proc)
    if (
        result
        and is_known_object_assignment_type(result, member_ctx)
        and proc.name.lower() not in out
    ):
        out[proc.name.lower()] = _LocalObjectVariable(name=proc.name, as_type=result, let_only=True)
    return out


def _unset_object_member_accesses(
    source: str,
    span: Span,
    locals_: Mapping[str, _LocalObjectVariable],
    state: Mapping[str, ObjectVariableState],
    member_ctx: MemberCompletionContext,
) -> list[tuple[str, Span]]:
    toks = statement_tokens(source, span)
    out: list[tuple[str, Span]] = []
    for i in range(len(toks) - 1):
        if toks[i + 1].raw_text != "." or _raw(_at(toks, i - 1)) == ".":
            continue
        name = token_name(toks[i])
        if not name:
            continue
        lower = name.lower()
        local = locals_.get(lower)
        if local is None or local.let_only or state.get(lower) != "unset":
            continue
        member = token_name(toks[i + 2]) if i + 2 < len(toks) else None
        if member and _has_definite_missing_member(
            source, span.start + toks[i + 1].end, member, member_ctx
        ):
            continue
        out.append((name, Span(span.start + toks[i].start, span.start + toks[i].end)))
    return out


def _has_definite_missing_member(
    source: str, dot_end_offset: int, member_name: str, member_ctx: MemberCompletionContext
) -> bool:
    surface = resolve_exhaustive_member_surface(source, dot_end_offset, member_ctx)
    return surface is not None and not surface.has_member(member_name)


def _with_dot_reaches(line: Sequence[VbaToken]) -> bool:
    for i, tok in enumerate(line):
        if tok.raw_text != "." or token_name(_at(line, i + 1)) is None:
            continue
        if i == 0:
            return True
        prev = line[i - 1]
        if (
            prev.kind is not TokenKind.IDENTIFIER
            and prev.kind is not TokenKind.BRACKETED_IDENTIFIER
            and prev.raw_text != ")"
            and prev.raw_text != "]"
            and token_text(prev) != "me"
        ):
            return True
    return False


def _unset_with_object_receiver(
    source: str,
    node: BodyNode,
    locals_: Mapping[str, _LocalObjectVariable],
    state: Mapping[str, ObjectVariableState],
    elements: _ObjectArrayElements,
) -> tuple[str, str | None, Span] | None:
    """`With o` on an object never set, or `With a(0)` on an element of a fixed
    array never set, whose body reaches it by a leading dot on a line that always
    runs. `With o` alone runs; the first `.Count` raises 91 (issue #295, measured
    in Excel 16.0). The hit is (name, the array for an element, span)."""
    header = block_header_line_span(source, node.span)
    toks = statement_tokens_after_leading_label(source, header)
    if token_text(_at(toks, 0)) != "with":
        return None
    name = token_name(_at(toks, 1))
    if not name:
        return None
    lower = name.lower()
    local = locals_.get(lower)
    found: tuple[str, str | None, Span] | None = None
    if len(toks) == 2 and local is not None and not local.let_only and state.get(lower) == "unset":
        found = (name, None, Span(header.start + toks[1].start, header.start + toks[1].end))
    elif (
        len(toks) == 5
        and toks[2].raw_text == "("
        and toks[3].kind is TokenKind.INTEGER_LITERAL
        and toks[4].raw_text == ")"
        and lower in elements.arrays
        and state.get(f"{lower}({_literal_index_text(toks[3].raw_text)})") == "unset"
    ):
        found = (
            "".join(tok.raw_text for tok in toks[1:]),
            elements.arrays[lower][0],
            Span(header.start + toks[1].start, header.start + toks[4].end),
        )
    body = getattr(node, "body", None)
    if found is None or not isinstance(body, list):
        return None
    reached = any(
        is_leaf_statement(child)
        and not (isinstance(child, StatementNode) and child.single_line_if_branches is not None)
        and _with_dot_reaches(statement_tokens_after_leading_label(source, child.span))
        for child in body
    )
    return found if reached else None


def _nothing_passed_to_union(
    toks: Sequence[VbaToken],
    locals_: Mapping[str, _LocalObjectVariable],
    state: Mapping[str, ObjectVariableState],
) -> list[tuple[str, int, int]]:
    """The locals still Nothing passed whole to Excel's Union: `Union(n, c)`,
    `Application.Union(c, n)`, as (name, start, end)."""
    out: list[tuple[str, int, int]] = []
    for i in range(len(toks) - 1):
        if (
            token_text(toks[i]) != "union"
            or toks[i + 1].raw_text != "("
            or range_method_owner(toks, i) is None
        ):
            continue
        close = match_paren_from(toks, i + 1)
        for arg in split_top_level_token_groups(toks, i + 2, ",", close) if close > i + 2 else []:
            named = [tok for tok in arg if tok.kind is not TokenKind.COMMENT]
            arg_name = token_name(named[0]) if len(named) == 1 else None
            lower = arg_name.lower() if arg_name is not None else None
            local = locals_.get(lower) if lower else None
            if (
                lower
                and local is not None
                and not local.let_only
                and not local.variant
                and state.get(lower) == "unset"
            ):
                out.append((named[0].raw_text, named[0].start, named[0].end))
    return out


def _set_assignment_value_is_nothing(value_tokens: Sequence[VbaToken]) -> bool:
    toks = [
        tok
        for tok in value_tokens
        if tok.kind is not TokenKind.COMMENT and tok.kind is not TokenKind.NEWLINE
    ]
    return len(toks) == 1 and token_text(toks[0]) == "nothing"
