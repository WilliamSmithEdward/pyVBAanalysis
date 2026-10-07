"""Rule family: assignment-statement rules.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/assignments.ts: constant
and procedure-name targets, scalar, array and member-access assignment types, Set
targets and their object types, missing return assignments, and the Mid-statement
literal target. The rules that read a statement structurally also read the
statements a single-line `If` carries.
"""

from __future__ import annotations

import json
import re
from collections.abc import Callable, Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass
from typing import Any

from ...completion.member_access import (
    MemberCompletion,
    MemberCompletionContext,
    is_late_bound_type_key,
    resolve_exact_member_completion,
    resolve_receiver_type_at,
    signature_declares_parameters,
)
from ...completion.member_access import (
    is_known_object_assignment_type as is_known_object_assignment_type_ctx,
)
from ...conditional import ConditionalActivityTracker, inactive_node_skip
from ...host.host_model import is_dispatch_only_host_type, resolve_host_enum
from ...js_compat import JS_WHITESPACE, js_number, js_number_to_string, js_trim
from ...lexer.token_helpers import match_paren_from, split_top_level_token_groups
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.expression_limits import MAX_EXPRESSION_DEPTH
from ...parser.nodes import (
    BodyNode,
    ForBlockNode,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    ProcKind,
    Span,
    StatementNode,
    iter_body_nodes,
)
from ...runtime.vba_runtime import resolve_runtime_function
from ...symbols.name_resolution import BareIdentifierContext, BareIdentifierResolutionScope
from ...symbols.symbol_model import (
    ModuleSymbols,
    SymbolVisibility,
    VbaProcedureSignature,
    VbaProjectClassMember,
    VbaProjectClassMembers,
    VbaSymbol,
    VbaSymbolKind,
    procedure_params_from_symbol,
)
from ...types.type_inference import (
    SCALAR_OBJECT_ASSIGNMENT_REASON,
    DeclaredValueShape,
    SourceDeclaredShape,
    create_object_assignment_type_resolver,
    create_project_interface_sharing_lookup,
    declaration_shape_environment_for,
    declared_shape_for_source_binding,
    declared_type_for_source_binding,
    def_type_of,
    known_local_literal_values_at,
    object_holding_default,
    procedure_symbol_for,
    read_only_host_default,
    sheets_from_collection_property,
    source_binding_type_resolvers,
    source_identifier_binding,
    type_environment_for,
    unreachable_statements_in,
)
from ...types.type_names import is_known_scalar_type, normalize_type, numeric_literal_bounds
from ..argument_inference import (
    SourceDeclaredTypeResolver,
    SourceQualifiedDeclaredTypeResolver,
    incompatibility_reason,
    infer_argument_type,
    nonnumeric_string_arithmetic_operand,
)
from ..assignment_coercion_type import create_assignment_coercion_type, array_element_identity, array_by_ref_identity
from ..setter_assignment import assignment_target_from_tokens, assignment_target_name, source_setter_assignment, invalid_setter_assignment_arity, project_setter_member
from ...symbols.class_member_facts import class_member_values
from ..call_extraction import (
    CallableTypeSignature,
    CallArguments,
    InferredArgumentType,
    extract_call,
    extract_qualified_call,
    CallableParamType,
    split_arg_slots,
    validate_arity,
    named_argument_slot,
    string_literal_value,
    unwrap_outer_parens,
)
from ..callable_signatures import (
    SourceNameScope,
    build_module_type_signatures,
    callable_signature_for_call,
    callable_type_signatures_for,
    is_member_statement_chain_through,
    runtime_callable_source_shadowed,
    source_name_scope_for,
    member_callable_signature,
)
from ..context import PushFn
from ..function_results import function_result_at, known_function_results
from ..held_objects import HeldObjects, held_objects_at
from ..known_locals import KnownLocalValue
from ..member_parameter_counts import member_parameter_counts
from ..null_operators import operator_yields_null
from ..straight_line_values import straight_line_assignments
from ..walker import (
    ProcedureStatementVisitor,
    active_module_members,
    bare_assignment_target,
    declared_name_span,
    first_executable_token_index,
    for_each_statement,
    set_assignment_target,
    statement_and_branch_spans,
    statement_tokens,
    statement_tokens_after_leading_label,
    strip_header_brackets,
    token_name,
    token_text,
    top_level_operator_index,
)
from .arrays import (
    FixedArrayBound,
    element_operand_starting_at,
    elements_written_in,
    known_array_shapes_at,
    module_option_base,
)
from .host_property_values import (
    formula_string_problem,
    host_property_value_problem,
    host_union_property_value_problem,
)
from .object_state import object_let_state_at
from .shared import name_mentions
from .type_of_is import object_assignment_incompatibility_reason, object_let_assignment_verdict

# JavaScript's `\s`, for the type-text patterns below.
_WS = "[" + re.escape(JS_WHITESPACE) + "]"
# `/\s*\(\s*\)\s*$/`: a trailing `()` and the space around it.
_TRAILING_PARENS_RE = re.compile(rf"{_WS}*\({_WS}*\){_WS}*\Z")
# `/\(\s*\)\s*$/`: whether a type text ends in `()`.
_ENDS_IN_PARENS_RE = re.compile(rf"\({_WS}*\){_WS}*\Z")


def _at(toks: Sequence[VbaToken], index: int) -> VbaToken | None:
    """`toks[index]` as JavaScript reads it: None out of range."""
    return toks[index] if 0 <= index < len(toks) else None


def _raw_at(toks: Sequence[VbaToken], index: int) -> str | None:
    """`toks[index]?.rawText`."""
    tok = _at(toks, index)
    return tok.raw_text if tok is not None else None


def _js_text(value: str | None) -> str:
    """A value in a JavaScript template string: undefined prints as "undefined"."""
    return "undefined" if value is None else value


def _json_string(value: str) -> str:
    """JSON.stringify of a string."""
    return json.dumps(value, ensure_ascii=False)


def _non_comment(toks: Sequence[VbaToken]) -> list[VbaToken]:
    return [tok for tok in toks if tok.kind is not TokenKind.COMMENT]


def _is_scalar_literal_token(tok: VbaToken) -> bool:
    """A literal that can never be an object reference: a number, string, date, True or False."""
    if tok.kind in (
        TokenKind.INTEGER_LITERAL,
        TokenKind.FLOAT_LITERAL,
        TokenKind.STRING_LITERAL,
        TokenKind.DATE_LITERAL,
    ):
        return True
    word = token_text(tok)
    return word in ("true", "false")


def check_const_assignment(
    source: str,
    symbols: ModuleSymbols,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    push: PushFn,
) -> ProcedureStatementVisitor:
    """Assigning to a constant is illegal. High-confidence form only: the
    left-hand side must be a bare identifier (no member access, no index) that
    resolves to a Const declared at module level or in the enclosing procedure."""

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        proc_sym = procedure_symbol_for(symbols, member)

        def visitor(stmt: LeafStatementNode) -> None:
            for span in statement_and_branch_spans(stmt):
                check_span(span)

        def check_span(span: Span) -> None:
            # `Set K = Nothing` assigns to the constant too (XLIDE issue #255).
            hit = bare_assignment_target(source, span)
            if hit is None:
                hit = set_assignment_target(source, span)
            if hit is None:
                return
            name, name_span, _value = hit
            binding = source_identifier_binding(
                symbols,
                proc_sym,
                project_visible_symbols,
                name,
                BareIdentifierContext.ASSIGNMENT_TARGET,
            )
            if binding.scope is BareIdentifierResolutionScope.AMBIGUOUS:
                return
            # An Enum member is a constant too (XLIDE issue #213): `eA = 2` is
            # "Assignment to constant not permitted".
            if any(
                d.kind in (VbaSymbolKind.CONSTANT, VbaSymbolKind.ENUM_MEMBER)
                for d in binding.definitions
            ):
                push("constAssignment", f"Cannot assign to constant '{name}'.", name_span)
                return
            target = _procedure_assignment_target(binding.definitions, proc_sym)
            if target is not None:
                push("assignmentToProcedureName", target(name), name_span)

        return visitor

    return factory


def _read_only_project_default(type_name: str, member_ctx: MemberCompletionContext) -> str | None:
    """The default member of a project class when it is a Property Get with no Property Let."""
    lower = js_trim(type_name).split(".")[-1].lower()
    cls = next(
        (
            candidate
            for candidate in member_ctx.project_class_members or ()
            if candidate.kind == "class" and candidate.name.lower() == lower
        ),
        None,
    )
    member = (
        next((candidate for candidate in cls.members if candidate.default_member), None)
        if cls is not None and cls.exhaustive is True
        else None
    )
    return (
        member.name
        if member is not None
        and member.kind == "property"
        and not member.let_accessor
        and member.writable is not True
        else None
    )


def _procedure_assignment_target(
    definitions: Sequence[VbaSymbol],
    proc_sym: VbaSymbol | None,
) -> Callable[[str], str] | None:
    """Why a procedure's name cannot be assigned to from outside it, measured in
    Excel 16.0 (XLIDE issue #213): a Sub's is "Expected Function or variable",
    and a Function's is "Function call on left-hand side of assignment must
    return Variant or Object" when it returns a type of VBA's own. Inside the
    Function the name is its return value and binds locally, so it never reaches
    here."""
    if len(definitions) != 1 or definitions[0] is proc_sym:
        return None
    definition = definitions[0]
    if definition.kind is VbaSymbolKind.SUB:
        return lambda name: (
            f"'{name}' is a Sub, which has no value to assign. This is a VBE compile error: "
            "Expected Function or variable."
        )
    returns = normalize_type(definition.as_type)
    # A Declare Function too (XLIDE issue #254): `GetTickCount = 5`.
    is_function = definition.kind is VbaSymbolKind.FUNCTION or (
        definition.kind is VbaSymbolKind.DECLARE and definition.declare_kind == "Function"
    )
    if (
        is_function
        and returns
        and returns != "variant"
        and is_known_scalar_type(returns)
        and not definition.is_array
    ):
        as_type = definition.as_type
        return lambda name: (
            f"'{name}' is a Function returning {as_type}, and a call cannot be assigned to. "
            "This is a VBE compile error: Function call on left-hand side of assignment must "
            "return Variant or Object."
        )
    return None


@dataclass(frozen=True, slots=True)
class _MemberAssignmentTarget:
    member: str
    label: str
    member_span: Span
    value_tokens: list[VbaToken]
    uses_set: bool
    # True for `wb.Name() = x`: the member is given arguments.
    with_arguments: bool
    has_arguments: bool
    argument_tokens: list[VbaToken]


def _member_assignment_target(source: str, span: Span) -> _MemberAssignmentTarget | None:
    """Port of memberAssignmentTarget: an `obj.Member = value` / `Set obj.Member = ...`
    LHS ending in `. Member` or `. Member(args)`. Returns None for bare or compound LHS."""
    toks = statement_tokens(source, span)
    i = first_executable_token_index(toks)
    if i >= len(toks):
        return None
    uses_set = token_text(toks[i]) == "set"
    if uses_set or token_text(toks[i]) == "let":
        i += 1
    eq = top_level_operator_index(toks[i:], "=")
    if eq < 0:
        return None
    equals_index = i + eq
    lhs = toks[i:equals_index]
    if len(lhs) < 2:
        return None
    # The member may be given arguments: `r.Address(False, False) = "B2"`.
    member_index = len(lhs) - 1
    with_arguments = lhs[member_index].raw_text == ")"
    if with_arguments:
        depth = 0
        while member_index >= 0:
            raw = lhs[member_index].raw_text
            depth += 1 if raw == ")" else -1 if raw == "(" else 0
            if depth == 0:
                break
            member_index -= 1
        member_index -= 1
    member_tok = _at(lhs, member_index)
    if member_tok is None or not token_name(member_tok) or _raw_at(lhs, member_index - 1) != ".":
        return None
    # A target is one receiver chain ending in the member. Anything else before
    # the `=` is another statement comparing the member: an ElseIf or Case
    # header, a single-line If's condition, a call given the comparison
    # (`Debug.Print w.Part = "a"`). ReDim's `ElseIf ReDimUI.SenderPart = "plus"
    # Then` compiles, and was reported as assigning to 'ElseIf
    # ReDimUI.SenderPart'.
    if not is_member_statement_chain_through(lhs, 0, member_index):
        return None
    if any(t.kind is TokenKind.OPERATOR and t.raw_text == "=" for t in lhs):
        return None
    member_name = token_name(member_tok)
    assert member_name is not None
    return _MemberAssignmentTarget(
        member=member_name,
        label=js_trim(source[span.start + lhs[0].start : span.start + lhs[-1].end]),
        member_span=Span(span.start + member_tok.start, span.start + member_tok.end),
        value_tokens=list(toks[equals_index + 1 :]),
        uses_set=uses_set,
        with_arguments=with_arguments,
        has_arguments=with_arguments and len(lhs) > member_index + 3,
        argument_tokens=list(lhs[member_index + 2:-1]) if with_arguments else [],
    )


@dataclass(frozen=True, slots=True)
class _ObjectFacts:
    is_object: bool
    verdict: str
    # objectHoldingDefault's { name, returns }.
    holding: Any
    read_only_default: str | None


@dataclass(frozen=True, slots=True)
class _NullSource:
    name: str
    span: Span
    returns: bool = False


@dataclass(frozen=True, slots=True)
class _NullExpression:
    text: str
    span: Span


def check_assignment_types(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    member_ctx: MemberCompletionContext,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """Scalar, array and member-access assignment type compatibility (`x = v`, `obj.M = v`)."""
    is_document_module = _project_type_name_lookup(member_ctx, "document", False)
    is_form_owner = _project_type_name_lookup(member_ctx, "userform", True)
    # Declared-type facts are stable within this rule invocation. Value and
    # object-state facts below still depend on the individual statement.
    object_types: dict[str, _ObjectFacts] = {}

    def object_facts_for(type_name: str) -> _ObjectFacts:
        facts = object_types.get(type_name)
        if facts is None:
            is_object = is_known_object_assignment_type_ctx(type_name, member_ctx)
            verdict = object_let_assignment_verdict(type_name, member_ctx) if is_object else "unknown"
            holding = (
                object_holding_default(type_name, member_ctx)
                if is_object and verdict != "noDefault"
                else None
            )
            read_only_default: str | None = None
            if is_object and not holding:
                read_only_default = _read_only_project_default(type_name, member_ctx)
                if read_only_default is None:
                    read_only_default = read_only_host_default(type_name, member_ctx)
            facts = _ObjectFacts(is_object, verdict, holding, read_only_default)
            object_types[type_name] = facts
        return facts

    # Base depends on this module/activity pass, not on a procedure or value.
    # Resolve it only when array folding needs it; zero is a cached result too.
    option_base: int | None = None

    def option_base_for() -> int:
        nonlocal option_base
        if option_base is None:
            option_base = module_option_base(mod, activity)
        return option_base

    # Null-choice shadowing checks only direct module names, not the broader
    # local/project runtime scope. At most three names are queried per pass.
    choice_module_names: dict[str, bool] = {}

    def choice_module_name_declared(lower: str) -> bool:
        if lower not in choice_module_names:
            choice_module_names[lower] = any(
                child.name.lower() == lower for child in symbols.root.children or ()
            )
        return choice_module_names[lower]

    # The Collection.Count guard considers any project surface named Collection.
    # Resolve this exact predicate only when queried, across the whole rule pass.
    project_collection: bool | None = None

    def project_declares_collection() -> bool:
        nonlocal project_collection
        if project_collection is None:
            project_collection = any(
                t.name.lower() == "collection" for t in member_ctx.project_class_members or ()
            )
        return project_collection

    module_signatures = build_module_type_signatures(symbols)
    # Enum assignment compatibility is a name query, not a full symbol scan
    # per assignment. Keep this index within the current rule pass.
    enum_names = {
        name
        for symbol in [*(symbols.root.children or ()), *(project_visible_symbols or ())]
        if symbol.kind is VbaSymbolKind.ENUM
        for name in (symbol.name.lower(), f"{symbol.module_name}.{symbol.name}".lower())
    }
    coercion_type = create_assignment_coercion_type(member_ctx, enum_names)
    member_coercion_type = create_assignment_coercion_type(member_ctx)
    def element_identity(type_: str | None) -> str:
        return array_element_identity(type_, member_ctx, coercion_type)

    def by_ref_identity(type_: str | None) -> str:
        return array_by_ref_identity(type_, member_ctx, enum_names)
    setter_names = {symbol.name.lower() for symbol in [*(symbols.root.children or ()), *(project_visible_symbols or ())] if symbol.kind is VbaSymbolKind.PROPERTY_LET}
    getter_names = {symbol.name.lower() for symbol in [*(symbols.root.children or ()), *(project_visible_symbols or ())] if symbol.kind in (VbaSymbolKind.PROPERTY_GET, VbaSymbolKind.FUNCTION)}
    own_getter_values: Mapping[str, str] | None = None

    def check_returned_object_default(type_: str, label: str, span: Span) -> bool:
        facts = object_facts_for(type_)
        if facts.holding or facts.read_only_default:
            name = facts.holding.name if facts.holding else facts.read_only_default
            push("invalidPropertyUse", f"Assignment through '{label}' reaches the default member {name} of {type_}, which has no writable Let contract. This is a VBE compile error: Invalid use of property.", span)
            return True
        if facts.verdict == "argument":
            push("argumentCount", f"Argument not optional: '{label}' returns {type_}, whose default member requires an index before a Let can reach it. This is a VBE compile error.", span)
            return True
        if facts.verdict == "noDefault":
            push("runtimeMemberNotFound", f"'{label}' returns {type_}, which has no default member to receive this Let assignment. This will raise Run-time error '438': Object doesn't support this property or method, or error '91' if the returned object is Nothing.", span)
            return True
        return False
    variant_array_functions = _array_only_variant_functions(source, mod, activity)

    def check_procedure(procedure: ProcedureNode) -> None:
        env = type_environment_for(symbols, procedure)
        shapes = declaration_shape_environment_for(symbols, procedure)
        source_names = source_name_scope_for(symbols, procedure, project_visible_symbols)
        proc_sym = procedure_symbol_for(symbols, procedure)
        # These value helpers inspect the same direct child declaration. Preserve
        # first-match semantics, but share queried names (and misses) per procedure.
        local_symbols: dict[str, VbaSymbol | None] = {}

        def local_symbol_named(lower: str) -> VbaSymbol | None:
            if lower not in local_symbols:
                children = proc_sym.children if proc_sym is not None and proc_sym.children else []
                local_symbols[lower] = next(
                    (child for child in children if child.name.lower() == lower), None
                )
            return local_symbols[lower]

        resolvers = source_binding_type_resolvers(symbols, proc_sym, project_visible_symbols)
        resolve_expression_type = resolvers.resolve_expression_type
        resolve_qualified_expression_type = resolvers.resolve_qualified_expression_type
        # What a Variant holds at a statement, and how often each name is
        # written anywhere: a Variant named once is never assigned, so Empty.
        shapes_at: Callable[[LeafStatementNode], Mapping[str, FixedArrayBound]] | None = None
        mentions: Mapping[str, int] | None = None
        reaching: Mapping[int, Mapping[str, Sequence[VbaToken]]] | None = None
        values_at: Callable[[LeafStatementNode], Mapping[str, KnownLocalValue]] | None = None
        unreachable: AbstractSet[int] | None = None
        written: AbstractSet[str] | None = None

        def shapes_at_stmt(stmt: LeafStatementNode) -> Mapping[str, FixedArrayBound]:
            nonlocal shapes_at
            if shapes_at is None:
                shapes_at = known_array_shapes_at(
                    source, symbols, procedure, activity, option_base_for()
                )
            return shapes_at(stmt)

        def name_mentions_once() -> Mapping[str, int]:
            nonlocal mentions
            if mentions is None:
                mentions = name_mentions(source, procedure, activity)
            return mentions

        def values_at_stmt(stmt: LeafStatementNode) -> Mapping[str, KnownLocalValue]:
            nonlocal values_at
            if values_at is None:
                values_at = known_local_literal_values_at(source, procedure, symbols, activity)
            return values_at(stmt)

        def plain_variant_local(lower: str) -> bool:
            local = local_symbol_named(lower)
            local_type = normalize_type(local.as_type) if local is not None else None
            return not (
                local is None
                or local.kind is not VbaSymbolKind.LOCAL_VARIABLE
                or local.visibility is SymbolVisibility.STATIC
                or local.is_array
                or (local_type is not None and local_type != "variant")
            )

        def array_value_at(stmt: LeafStatementNode, name: str) -> _ArrayValue | None:
            lower = name.lower()
            if not plain_variant_local(lower):
                return None
            shape = shapes_at_stmt(stmt).get(lower)
            if shape is not None:
                return _ArrayValue(
                    "string" if shape.origin == "Split(...)" else "variant",
                    f"'{name}', which holds an array from {shape.origin}",
                )
            return (
                _ArrayValue("empty", f"'{name}', which is never assigned and so is Empty")
                if name_mentions_once().get(lower) == 1
                else None
            )

        # `v = Null` then `s = v`: the Null a Variant local holds here, from
        # its last assignment in a straight line (XLIDE issue #239). A
        # single-line If's branch sees what held before the If, less whatever
        # the If touches.
        def null_held_at(
            stmt: LeafStatementNode, span: Span, value_tokens: Sequence[VbaToken]
        ) -> _NullSource | None:
            nonlocal reaching
            value = _non_comment(value_tokens)
            # `n = F()` with F a Function of the module returning Null (XLIDE issue #448).
            called = (
                function_result_at(
                    value, 0, known_function_results(source, mod, activity), procedure, symbols
                )
                if value
                else None
            )
            if called is not None and called.end == len(value) - 1:
                at = Span(span.start + value[0].start, span.start + value[called.end].end)
                return (
                    _NullSource(source[at.start : at.end], at, True)
                    if called.result.kind == "null"
                    else None
                )
            name = token_name(value[0]) if len(value) == 1 else None
            if not name:
                return None
            lower = name.lower()
            if not plain_variant_local(lower):
                return None
            if reaching is None:
                reaching = straight_line_assignments(source, procedure.body, activity)
            reached = reaching.get(id(stmt))
            held_raw = reached.get(lower) if reached is not None else None
            held = _non_comment(held_raw) if held_raw is not None else None
            return (
                _NullSource(name, Span(span.start + value[0].start, span.start + value[0].end))
                if held is not None and len(held) == 1 and token_text(held[0]) == "null"
                else None
            )

        # `n = 1 + Null`, `s = "a" + v` with v holding Null: arithmetic, `+`,
        # unary minus, Not, a comparison and Abs give Null when an operand is
        # Null, and so do Xor and Eqv; And, Or and Imp are judged only with
        # Null on every side; `&` never does (XLIDE issues #324 and #556,
        # measured in Excel 16.0).
        def null_expression_at(
            stmt: LeafStatementNode, span: Span, value_tokens: Sequence[VbaToken]
        ) -> _NullExpression | None:
            value = unwrap_outer_parens(_non_comment(value_tokens))
            if len(value) < 2:
                return None

            def holds_null(tok: VbaToken) -> bool:
                return token_text(tok) == "null" or null_held_at(stmt, span, [tok]) is not None

            if not operator_yields_null(value, holds_null):
                return None
            text = re.sub(r" ?([()]) ?", r"\1", " ".join(tok.raw_text for tok in value))
            return _NullExpression(
                text, Span(span.start + value[0].start, span.start + value[-1].end)
            )

        # `s = "b"` then `n = s`: the String a local or an array element is
        # known to hold here, from its last assignment in a straight line.
        def known_string_at(
            stmt: LeafStatementNode,
            span: Span,
            value_tokens: Sequence[VbaToken],
            expected: str,
        ) -> InferredArgumentType | None:
            nonlocal unreachable, written
            value = unwrap_outer_parens(_non_comment(value_tokens))
            if not value:
                return None
            if unreachable is None:
                unreachable = unreachable_statements_in(source, procedure, symbols, activity)
            if id(stmt) in unreachable:
                return None
            value_span = Span(span.start + value[0].start, span.start + value[-1].end)
            value_text = source[value_span.start : value_span.end]
            label = f"'{value_text}', which holds"
            # `n = F()` with F a Function of the module returning "abc" (XLIDE issue #448).
            called = function_result_at(
                value, 0, known_function_results(source, mod, activity), procedure, symbols
            )
            if called is not None and called.end == len(value) - 1:
                return (
                    InferredArgumentType(
                        type_="String",
                        label=f"'{value_text}', which returns {_json_string(called.result.value)}",
                        span=value_span,
                        string_value=called.result.value,
                    )
                    if called.result.kind == "string"
                    else None
                )
            if len(value) == 1:
                name = token_name(value[0])
                lower = name.lower() if name is not None else None
                known = values_at_stmt(stmt).get(lower) if lower else None
                if known is not None and known.kind == "string" and not known.content_mutated:
                    assert isinstance(known.value, str)
                    return InferredArgumentType(
                        type_="String",
                        label=f"{label} {_json_string(known.value)}",
                        span=value_span,
                        string_value=known.value,
                    )
                # A `String * 3` local named nowhere else holds three Chr(0),
                # which convert to no number, Boolean or date (XLIDE issue #451,
                # measured in Excel 16.0).
                local = local_symbol_named(lower) if lower else None
                length = (
                    int(local.fixed_length)
                    if local is not None
                    and local.kind is VbaSymbolKind.LOCAL_VARIABLE
                    and local.visibility is not SymbolVisibility.STATIC
                    and not local.is_array
                    and re.fullmatch(r"[0-9]+", local.fixed_length or "") is not None
                    and local.fixed_length is not None
                    else None
                )
                if (
                    length is not None
                    and length >= 1
                    and lower is not None
                    and name_mentions_once().get(lower) == 1
                ):
                    return InferredArgumentType(
                        type_="String",
                        label=(
                            f"'{value[0].raw_text}', a String * {length} never assigned, "
                            f"which holds {length} Chr(0)"
                        ),
                        span=value_span,
                        string_value="\x00" * length,
                    )
                return None
            # `"a" & "b"`, `Left("abc", 1)`, `o & 5` (XLIDE issue #405). A Date
            # written as text converts back to a Date, so that target is left alone.
            known_at = values_at_stmt(stmt)
            spelled = _spelled_text(value, known_at.get, env, source_names, True)
            if spelled is not None:
                if spelled.stand_in and normalize_type(expected) == "date":
                    return None
                if spelled.stand_in:
                    spelled_label = f"{label} a Date written as text"
                elif spelled.named is not None:
                    spelled_label = f"{label} {spelled.named}"
                else:
                    spelled_label = f"{label} {_json_string(spelled.text)}"
                return InferredArgumentType(
                    type_="String",
                    label=spelled_label,
                    span=value_span,
                    string_value=spelled.text,
                )
            # `v = Array("1", "b")` then `n = v(1)` (XLIDE issue #260).
            if written is None:
                written = elements_written_in(source, procedure, activity)
            first_name = token_name(value[0])
            if (first_name.lower() if first_name is not None else "") in written:
                return None
            element = element_operand_starting_at(
                value, 0, shapes_at_stmt(stmt), option_base_for()
            )
            return (
                InferredArgumentType(
                    type_="String",
                    label=f"{label} {_json_string(element.value)}",
                    span=value_span,
                    string_value=element.value,
                )
                if element is not None
                and element.last == len(value) - 1
                and isinstance(element.value, str)
                else None
            )

        # `a(0) = New Collection` with a an array As Collection: a Let into the
        # element reaches the default member Item, which needs an argument, as
        # it does into a variable (XLIDE issue #306, measured in Excel 16.0).
        def check_element_let(span: Span) -> None:
            element = _array_element_target(
                source, span, symbols, proc_sym, project_visible_symbols
            )
            if element is None or element.uses_set:
                return
            declared = declared_type_for_source_binding(
                symbols,
                proc_sym,
                project_visible_symbols,
                element.name,
                BareIdentifierContext.ASSIGNMENT_TARGET,
            )
            expected = declared.as_type if declared.resolved else None
            facts = object_facts_for(expected) if expected else None
            if facts and facts.read_only_default:
                push("readonlyMemberAssignment", f"Assignment to '{element.label}' reaches the default member {facts.read_only_default} of {expected}, a Property Get with no Property Let. This is a VBE compile error: Invalid use of property.", element.span)
                return
            if expected and object_facts_for(expected).verdict == "argument":
                error = (
                    "Argument not optional"
                    if normalize_type(expected) == "collection"
                    else "Invalid use of property"
                )
                push(
                    "setRequired",
                    f"Assignment to '{element.label}' requires Set: the default member of "
                    f"{expected} takes an argument, so a Let cannot reach it. This is a VBE "
                    f"compile error: {error}.",
                    element.span,
                )

        def resolve_target_shape(name: str) -> SourceDeclaredShape:
            return declared_shape_for_source_binding(
                symbols, proc_sym, project_visible_symbols, name,
                BareIdentifierContext.ASSIGNMENT_TARGET,
            )

        def resolve_source_shape(name: str) -> SourceDeclaredShape:
            return declared_shape_for_source_binding(
                symbols, proc_sym, project_visible_symbols, name,
                BareIdentifierContext.EXPRESSION,
            )

        def infer_value(tokens: Sequence[VbaToken], offset: int) -> InferredArgumentType | None:
            return infer_argument_type(tokens, offset, env, module_signatures, source_names,
                source=source, member_ctx=member_ctx, resolve_expression_type=resolve_expression_type,
                resolve_qualified_expression_type=resolve_qualified_expression_type)

        def check_bare_getter(span: Span) -> bool:
            nonlocal own_getter_values
            if not getter_names:
                return False
            tokens = statement_tokens_after_leading_label(source, span)
            first = first_executable_token_index(tokens)
            if token_text(_at(tokens, first)) == "let":
                first += 1
            root = token_name(_at(tokens, first))
            if not root or root.lower() not in getter_names:
                return False
            equals = top_level_operator_index(tokens, "=")
            target = assignment_target_from_tokens(tokens[:equals + 1]) if equals >= 0 else None
            named = assignment_target_name(target) if target else None
            if target is None or named is None or named[0] != 0:
                return False
            binding = source_identifier_binding(symbols, proc_sym, project_visible_symbols, root, BareIdentifierContext.ASSIGNMENT_TARGET)
            if binding.scope is BareIdentifierResolutionScope.AMBIGUOUS or any(definition.kind in (VbaSymbolKind.PROPERTY_LET, VbaSymbolKind.PROPERTY_SET) for definition in binding.definitions):
                return False
            getter = next((definition for definition in binding.definitions if definition.kind in (VbaSymbolKind.PROPERTY_GET, VbaSymbolKind.FUNCTION)), None)
            if getter is None or getter is proc_sym:
                return False
            declared = getter.as_type or (def_type_of(symbols, getter.name) if getter.module_name.lower() == symbols.module_name.lower() else "Variant")
            if getter.kind is VbaSymbolKind.FUNCTION and not is_known_object_assignment_type_ctx(declared, member_ctx):
                return False
            parameters = [CallableParamType(name=param.name, type_=param.type_, optional=param.optional, param_array=param.param_array) for param in procedure_params_from_symbol(getter)]
            name_span = Span(span.start + target[0].start, span.start + target[0].end)
            if not named[1] and any(not param.optional and not param.param_array for param in parameters):
                push("argumentCount", f"Argument not optional: '{root}' requires a getter argument.", name_span)
                return True
            result_indexed = named[1] and len(target) > 3 and not parameters
            if not result_indexed and _invalid_getter_argument_count(source, root, name_span, parameters, target[2:-1] if named[1] else [], span.start):
                return True
            if declared and _getter_may_return_object(declared, member_ctx) and not result_indexed and check_returned_object_default(declared, root, name_span):
                return True
            if getter.kind is VbaSymbolKind.FUNCTION or normalize_type(declared) != "variant":
                return False
            if getter.module_name.lower() == symbols.module_name.lower():
                if own_getter_values is None:
                    own_getter_values = class_member_values(source, symbols.root.children or [])
                value = own_getter_values.get(getter.name.lower())
            else:
                member = project_setter_member(member_ctx, getter.module_name, getter.name)
                value = member.known_value if member else None
            if value not in ("scalar", "empty"):
                return False
            push("variantValueMisuse", f"'{root}' has only a Property Get returning a Variant that holds no object. The Let writes through its returned value, which cannot receive a property assignment. This will raise Run-time error '424': Object required.", name_span)
            return True

        def check_bare_setter(span: Span, stmt: LeafStatementNode) -> bool:
            if not setter_names:
                return False
            tokens = statement_tokens_after_leading_label(source, span)
            first = first_executable_token_index(tokens)
            if token_text(_at(tokens, first)) == "if":
                return False
            if token_text(_at(tokens, first)) == "let":
                first += 1
            root = token_name(_at(tokens, first))
            if not root or root.lower() not in setter_names:
                return False
            equals = top_level_operator_index(tokens, "=")
            target = assignment_target_from_tokens(tokens[:equals + 1]) if equals >= 0 else None
            named = assignment_target_name(target) if target else None
            if target is None or named is None or named[0] != 0:
                return False
            name = token_name(target[0])
            if not name:
                return False
            binding = source_identifier_binding(symbols, proc_sym, project_visible_symbols, name, BareIdentifierContext.ASSIGNMENT_TARGET)
            if binding.scope is BareIdentifierResolutionScope.AMBIGUOUS:
                return False
            setter = next((definition for definition in binding.definitions if definition.kind is VbaSymbolKind.PROPERTY_LET), None)
            params = [child for child in setter.children or [] if child.kind is VbaSymbolKind.PARAMETER] if setter else []
            parameter = params[-1] if params else None
            declared = parameter.as_type or (def_type_of(symbols, parameter.name) if parameter.module_name.lower() == symbols.module_name.lower() else None) if parameter else None
            if declared is None and setter:
                member = project_setter_member(member_ctx, setter.module_name, setter.name)
                source_params = (member.procedure_params or {}).get("propertyLet") if member and member.let_accessor else None
                declared = (source_params[-1].type_ if source_params else None) or (member.write_type if member and member.let_accessor else None)
            value = tokens[equals + 1:]
            actual = infer_value(value, span.start)
            if parameter and parameter.is_array:
                _check_array_setter_value(name, declared, value, span.start, actual, coercion_type, resolve_expression_type, resolve_qualified_expression_type, push, source_names, project_declares_collection, lambda type_: object_let_assignment_verdict(type_, member_ctx), lambda name: array_value_at(stmt, name), by_ref_identity)
                return True
            if not declared:
                return False
            expected = coercion_type(declared)
            if not is_known_scalar_type(normalize_type(expected) or ""):
                return False
            problem = _array_assignment_problem(name, actual.span if actual else span, value, span.start, DeclaredValueShape(expected, False, False), lambda name: resolve_source_shape(name).shape, lambda name: array_value_at(stmt, name), lambda _: actual, source_names, member_ctx=member_ctx, source=source)
            if problem:
                push(problem.code, problem.message, problem.span)
                return True
            reason = incompatibility_reason(expected, actual) if actual else None
            if reason and actual:
                push("assignmentTypeMismatch", f"Assignment to '{name}' expects {declared}, but got {actual.label}. {reason}", actual.span)
            return True

        def check_assignment_span(span: Span, stmt: LeafStatementNode) -> None:
            setter = source_setter_assignment(source, span, symbols, proc_sym, project_visible_symbols, member_ctx)
            if setter and invalid_setter_assignment_arity(setter, source, push):
                return
            if check_bare_setter(span, stmt) or check_bare_getter(span):
                return
            assignment = bare_assignment_target(source, span)
            if assignment is None:
                check_element_let(span)
                return
            name, name_span, value_tokens = assignment
            target_type = declared_type_for_source_binding(
                symbols,
                proc_sym,
                project_visible_symbols,
                name,
                BareIdentifierContext.ASSIGNMENT_TARGET,
            )
            # `Dim a()` with no As clause is an array of Variant (XLIDE issue #222).
            untyped_array = False
            if not target_type.as_type:
                shape = resolve_target_shape(name)
                untyped_array = shape.resolved and shape.shape is not None and shape.shape.is_array
            declared_expected: str | None
            if target_type.resolved:
                declared_expected = target_type.as_type
                if declared_expected is None:
                    declared_expected = None if untyped_array else def_type_of(symbols, name)
            else:
                declared_expected = env.get(name.lower())
            if declared_expected is None and untyped_array:
                declared_expected = "Variant"
            # A variable As an Enum is a Long: `x = "abc"` raises 13 and
            # `x = 3000000000#` 6 (XLIDE issue #436, measured in Excel 16.0).
            expected = coercion_type(declared_expected) if declared_expected else None
            # `Sheet1 = 5` compiles as a Let through the document's default
            # member, and a Worksheet or Workbook has none (XLIDE issue #225).
            if not expected and not target_type.resolved and is_document_module(name):
                # Word refuses `ThisDocument = 5` while compiling (XLIDE issue #228).
                model = member_ctx.model
                if model is not None and model.get("hostName") == "Word":
                    push(
                        "setRequiresObject",
                        f"'{name}' names the document module itself, which no assignment can "
                        "replace. This is a VBE compile error: Invalid use of property.",
                        name_span,
                    )
                    return
                push(
                    "setRequired",
                    f"'{name}' names the document module itself: a Let reaches it through its "
                    "default member, which a document does not have, and a Set cannot replace it "
                    "either. This will raise Run-time error '438': Object doesn't support this "
                    "property or method.",
                    name_span,
                )
                return
            if not expected:
                return
            resolved_target_shape = resolve_target_shape(name)
            target_shape = resolved_target_shape.shape if resolved_target_shape.resolved else shapes.get(name.lower())
            object_type = None if target_shape is not None and target_shape.is_array else object_facts_for(expected)
            if object_type is not None and object_type.is_object:
                # The VBE compiles a bare `=` to an object variable as a Let
                # through the type's default member (XLIDE issue #107): `r = 5`
                # writes the Range's Value. What is reported is what the
                # default member makes of it.
                verdict = object_type.verdict
                # `x = 5` on a Word Paragraph: its default member Range holds an
                # object, which a Let cannot write (XLIDE issue #462, measured in
                # Word 16.0). A DAO Recordset's Fields holds an object too (#464).
                holding = object_type.holding
                if holding:
                    push(
                        "invalidPropertyUse",
                        f"Assignment to '{name}' reaches the default member {holding.name} of "
                        f"{expected}, which holds an object ({holding.returns}), so a Let cannot "
                        "write it. This is a VBE compile error: Invalid use of property.",
                        name_span,
                    )
                    return
                # A class whose default member is a Property Get with no Let:
                # `c = 5` does not compile (XLIDE issue #256, measured in Excel
                # 16.0). So does a Word Document, whose Name takes none (#438).
                read_only_default = object_type.read_only_default
                if read_only_default:
                    push(
                        "readonlyMemberAssignment",
                        f"Assignment to '{name}' reaches the default member {read_only_default} "
                        f"of {expected}, a Property Get with no Property Let. This is a VBE "
                        "compile error: Invalid use of property.",
                        name_span,
                    )
                    return
                if verdict == "argument":
                    # A Collection says "Argument not optional"; Excel's
                    # collections, Hyperlinks and Workbooks among them, say
                    # "Invalid use of property" (XLIDE issue #221, measured).
                    error = (
                        "Argument not optional"
                        if normalize_type(expected) == "collection"
                        else "Invalid use of property"
                    )
                    push(
                        "setRequired",
                        f"Assignment to '{name}' requires Set: the default member of {expected} "
                        "takes an argument, so a Let cannot reach it. This is a VBE compile "
                        f"error: {error}.",
                        name_span,
                    )
                elif verdict == "noDefault":
                    # While the object is Nothing the Let raises 91, and 438 only
                    # once it holds one (XLIDE issue #193). The object-state walk
                    # says which; this rule owns the report either way, since the
                    # fix is the Set.
                    state = object_let_state_at(
                        source, mod, procedure, symbols, member_ctx, activity, name_span.start
                    )
                    lower = name.lower()
                    declared = next(
                        (
                            child
                            for child in (
                                proc_sym.children if proc_sym is not None and proc_sym.children else []
                            )
                            if child.name.lower() == lower
                        ),
                        None,
                    )
                    if declared is None:
                        declared = next(
                            (
                                child
                                for child in symbols.root.children or ()
                                if child.name.lower() == lower
                            ),
                            None,
                        )
                    if state == "unset":
                        run_error = (
                            "It is still Nothing here, so this will raise Run-time error '91': "
                            "Object variable or With block variable not set."
                        )
                    elif state == "set" or (declared is not None and declared.is_auto_instantiated):
                        run_error = (
                            "This will raise Run-time error '438': Object doesn't support this "
                            "property or method."
                        )
                    else:
                        run_error = (
                            "This will raise Run-time error '438': Object doesn't support this "
                            "property or method, or '91' while it is Nothing."
                        )
                    push(
                        "setRequired",
                        f"Assignment to '{name}' requires Set: {expected} has no default member "
                        f"for a Let to reach. {run_error}",
                        name_span,
                    )
                return
            array_source = _array_assignment_to_scalar_source(
                name, value_tokens, span.start, expected, shapes,
                resolve_target_shape, resolve_source_shape,
            )
            if array_source is not None:
                src_name, src_span = array_source
                push(
                    "arrayAssignmentToScalar",
                    f"Array variable '{src_name}' cannot be assigned to scalar '{name}'. "
                    "Assign an array element or use a Variant/array target.",
                    src_span,
                )
                return
            # A dynamic Byte array takes a String whole - `b = "abc"` copies the
            # string's bytes, and a String takes the array back (XLIDE issue
            # #105, measured in Excel 16.0). The element type is not what the
            # value is checked against there.
            resolved_target_shape = resolve_target_shape(name)
            target_shape = (
                resolved_target_shape.shape
                if resolved_target_shape.resolved
                else shapes.get(name.lower())
            )

            def source_shape(source_name: str) -> DeclaredValueShape | None:
                resolved = resolve_source_shape(source_name)
                return resolved.shape if resolved.resolved else shapes.get(source_name.lower())

            def scalar_type(tokens: list[VbaToken]) -> InferredArgumentType | None:
                return infer_argument_type(
                    tokens, span.start, env, module_signatures, source_names,
                    source=source, member_ctx=member_ctx,
                    resolve_expression_type=resolve_expression_type,
                    resolve_qualified_expression_type=resolve_qualified_expression_type,
                )

            def returns_variant_array(callee: str) -> bool:
                lower_callee = callee.lower()
                children = proc_sym.children if proc_sym is not None and proc_sym.children else []
                return lower_callee in variant_array_functions and not any(
                    child.name.lower() == lower_callee for child in children
                )

            array_problem = _array_assignment_problem(
                name,
                name_span,
                value_tokens,
                span.start,
                target_shape,
                source_shape,
                lambda value_name: array_value_at(stmt, value_name),
                scalar_type,
                source_names,
                returns_variant_array,
                member_ctx=member_ctx,
                source=source,
                array_element_identity=element_identity,
            )
            if array_problem is not None:
                push(array_problem.code, array_problem.message, array_problem.span)
                return
            if (
                target_shape is not None
                and target_shape.is_array
                and normalize_type(target_shape.as_type) == "byte"
            ):
                return
            string_arithmetic = nonnumeric_string_arithmetic_operand(
                expected, value_tokens, span.start
            )
            if string_arithmetic is not None:
                push(
                    "stringArithmeticCoercion",
                    f"Assignment to '{name}' expects {expected}, but this numeric expression "
                    f"contains {string_arithmetic.label}. This will raise Run-time error '13': "
                    "Type mismatch.",
                    string_arithmetic.span,
                )
                return
            actual = scalar_type(value_tokens)
            scalar_target = is_known_scalar_type(normalize_type(expected) or "")
            null_call = _null_from_choice(value_tokens, choice_module_name_declared)
            if null_call is not None and scalar_target:
                push(
                    "assignmentTypeMismatch",
                    f"Assignment to '{name}' expects {expected}, but {null_call.why}, so it "
                    "returns Null. Null cannot be coerced to this scalar type. This will raise "
                    "Run-time error '94': Invalid use of Null.",
                    Span(span.start + null_call.first.start, span.start + null_call.last.end),
                )
                return
            null_source = null_held_at(stmt, span, value_tokens)
            if null_source is not None and scalar_target:
                held = "returns Null" if null_source.returns else "holds Null here"
                push(
                    "assignmentTypeMismatch",
                    f"Assignment to '{name}' expects {expected}, but '{null_source.name}' {held}. "
                    "Null cannot be coerced to this scalar type. This will raise Run-time error "
                    "'94': Invalid use of Null.",
                    null_source.span,
                )
                return
            null_value = null_expression_at(stmt, span, value_tokens) if scalar_target else None
            if null_value is not None:
                push(
                    "assignmentTypeMismatch",
                    f"Assignment to '{name}' expects {expected}, but '{null_value.text}' is Null: "
                    "an operator on Null gives Null. Null cannot be coerced to this scalar type. "
                    "This will raise Run-time error '94': Invalid use of Null.",
                    null_value.span,
                )
                return
            known_string = (
                known_string_at(stmt, span, value_tokens, expected) if scalar_target else None
            )
            if known_string is not None:
                string_reason = incompatibility_reason(expected, known_string)
                if string_reason:
                    push(
                        "assignmentTypeMismatch",
                        f"Assignment to '{name}' expects {expected}, but {known_string.label} "
                        f"here. {string_reason.replace('This string literal', 'This string', 1)}",
                        known_string.span,
                    )
                return
            if actual is None:
                return
            reason = incompatibility_reason(expected, actual)
            if not reason:
                return
            # A constant out of the target's range overflows; its type is no
            # mismatch: `b = vbTrue` stores -1 in a Byte (XLIDE issue #326).
            bounds = (
                numeric_literal_bounds(normalize_type(expected) or "")
                if actual.numeric_constant_name is not None and actual.numeric_value is not None
                else None
            )
            if bounds is not None and actual.numeric_value is not None:
                article = "an" if re.match(r"[AEIOU]", bounds.label) else "a"
                message = (
                    f"Assignment to '{name}' stores {actual.numeric_constant_name}, which is "
                    f"{js_number_to_string(actual.numeric_value)}, in {article} {bounds.label}, "
                    f"whose range is {bounds.min} to {bounds.max}. This will raise Run-time "
                    "error '6': Overflow."
                )
            else:
                message = (
                    f"Assignment to '{name}' expects {expected}, but got {actual.label}. {reason}"
                )
            push("assignmentTypeMismatch", message, actual.span)

        def visit(stmt: LeafStatementNode) -> None:
            for span in statement_and_branch_spans(stmt):
                check_assignment_span(span, stmt)

        for_each_statement(procedure.body, visit, activity)
        check_member_assignment_types(
            source,
            procedure,
            env,
            module_signatures,
            source_names,
            member_ctx,
            activity,
            push,
            project_declares_collection,
            is_form_owner,
            resolve_expression_type,
            resolve_qualified_expression_type,
            symbols,
            coercion_type=member_coercion_type,
            check_returned_object_default=check_returned_object_default,
            array_value_at=array_value_at,
            project_visible_symbols=project_visible_symbols,
            array_by_ref_identity=by_ref_identity,
        )

    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            check_procedure(member)


@dataclass(frozen=True, slots=True)
class _NullChoice:
    why: str
    first: VbaToken
    last: VbaToken


def _getter_may_return_object(type_: str | None, ctx: MemberCompletionContext) -> bool:
    if _ENDS_IN_PARENS_RE.search(type_ or ""):
        return False
    normalized = normalize_type(type_)
    return normalized in (None, "variant", "object") or is_known_object_assignment_type_ctx(type_, ctx)


def _invalid_getter_argument_count(source: str, name: str, span: Span, params: Sequence[CallableParamType], tokens: Sequence[VbaToken], base: int) -> bool:
    split = split_arg_slots(tokens, base) if tokens else None
    invalid = False

    def report(rule: str, message: str, span: Span, data: Any = None) -> None:
        nonlocal invalid
        invalid = True

    validate_arity(source, CallableTypeSignature(name, list(params)), CallArguments(name=name, name_span=span, slots=split.slots if split else [], slot_spans=split.spans if split else [], slice_start=base), report)
    return invalid


def _check_array_setter_value(label: str, expected: str | None, tokens: Sequence[VbaToken], base: int,
    actual: InferredArgumentType | None, coercion: Callable[[str], str],
    resolve_type: SourceDeclaredTypeResolver | None, resolve_qualified_type: SourceQualifiedDeclaredTypeResolver | None,
    push: PushFn, source_names: SourceNameScope, project_collection: Callable[[], bool],
    object_verdict: Callable[[str | None], str], variant_value: Callable[[str], _ArrayValue | None],
    by_ref_identity: Callable[[str | None], str] = lambda value: _element_type(value),
) -> None:
    raw = [token for token in tokens if token.kind not in (TokenKind.COMMENT, TokenKind.NEWLINE)]
    if not raw:
        return
    value = unwrap_outer_parens(raw)
    expected_type = normalize_type(expected)
    actual_type = normalize_type(actual.type_ if actual else None)
    typed_array = actual is not None and _ENDS_IN_PARENS_RE.search(actual.type_) is not None
    expected_array_type = by_ref_identity(expected)
    actual_array_type = by_ref_identity(actual.type_ if actual else None)
    same_storage = ("longlong" if actual_array_type == "longptr" else actual_array_type) == ("longlong" if expected_array_type == "longptr" else expected_array_type)
    wrong_element = typed_array and expected_array_type != actual_array_type and not (is_known_scalar_type(expected_array_type) and is_known_scalar_type(actual_array_type) and same_storage)
    indexed = len(value) > 1 and value[1].raw_text == "(" and match_paren_from(value, 1) == len(value) - 1
    name = token_name(value[0]) if len(value) == 1 or indexed else None
    qualified = len(value) == 3 and value[1].raw_text == "." and token_name(value[0]) and token_name(value[2])
    declared = resolve_type(name) if name and resolve_type else resolve_qualified_type(value[0].raw_text, value[2].raw_text) if qualified and resolve_qualified_type else None
    span = Span(base + raw[0].start, base + raw[-1].end)
    kinds = (VbaSymbolKind.LOCAL_VARIABLE, VbaSymbolKind.MODULE_VARIABLE, VbaSymbolKind.PARAMETER)
    array_variable = declared is not None and declared.is_array and declared.kind in kinds
    if array_variable and (not indexed or len(value) == 3):
        push("arrayTargetAssignment", f"Assignment to '{label}' passes a whole array to a Property Let value parameter. This is a VBE compile error: Can't assign to array.", span)
        return
    string_element = array_variable and indexed and len(value) > 3 and declared is not None and normalize_type(declared.as_type) == "string"
    if expected_type == "byte" and ((actual_type == "string" and not typed_array) or string_element):
        return
    if len(value) == 2 and token_text(value[0]) == "new" and normalize_type(value[1].raw_text) == "collection" and not project_collection():
        push("argumentCount", f"Assignment to '{label}' reads the Collection's default member Item, which requires an index. This is a VBE compile error: Argument not optional.", span)
        return
    variant_variable = declared is not None and declared.resolved and not declared.is_array and declared.kind in kinds and (normalize_type(declared.as_type) or "variant") == "variant"
    produced = _array_produced_by(value, source_names)
    if expected_type == "byte" and (variant_variable or produced):
        def shape(name: str) -> DeclaredValueShape | None:
            binding = resolve_type(name) if resolve_type else None
            return DeclaredValueShape(binding.as_type, bool(binding.is_array), False) if binding and binding.resolved else None

        problem = _array_assignment_problem(label, span, value, base, DeclaredValueShape("Byte", True, False), shape, variant_value, lambda _: actual, source_names)
        if problem:
            push(problem.code, problem.message, problem.span)
        return
    if wrong_element or (array_variable and indexed and len(value) > 3) or variant_variable or (actual and not typed_array and is_known_scalar_type(normalize_type(coercion(actual.type_)) or "")) or produced or (actual and not typed_array and object_verdict(actual.type_) == "noDefault"):
        push("argumentShapeMismatch", f"Assignment to '{label}' passes {actual.label if actual else 'an array element'}, but the Property Let value parameter requires an array. This is a VBE compile error: Type mismatch: array or user-defined type expected.", span)


def _null_from_choice(
    value_tokens: Sequence[VbaToken],
    module_name_declared: Callable[[str], bool],
) -> _NullChoice | None:
    """A Choose, Switch or IIf whose literal arguments make it return Null
    (XLIDE issue #243, measured in Excel 16.0): `Choose(5, 1, 2)` and
    `Choose(0, "a")` name no choice, `Switch(False, 1)` finds no True
    condition, and `IIf(False, 1, Null)` takes its Null. A module procedure of
    the same name is the module's."""
    toks = _non_comment(value_tokens)
    start = 2 if token_text(_at(toks, 0)) == "vba" and _raw_at(toks, 1) == "." else 0
    fn = token_text(_at(toks, start))
    if (
        fn not in ("choose", "switch", "iif")
        or _raw_at(toks, start + 1) != "("
        or match_paren_from(toks, start + 1) != len(toks) - 1
    ):
        return None
    if start == 0 and module_name_declared(fn):
        return None
    args = split_top_level_token_groups(toks, start + 2, ",", len(toks) - 1)

    def literal(arg: Sequence[VbaToken]) -> float | None:
        word = token_text(arg[0]) if len(arg) == 1 else ""
        if word in ("true", "false"):
            return -1 if word == "true" else 0
        signed = len(arg) == 2 and arg[0].raw_text == "-"
        number = arg[1] if signed else arg[0] if len(arg) == 1 else None
        value = (
            js_number(re.sub(r"[%&^]\Z", "", number.raw_text))
            if number is not None and number.kind is TokenKind.INTEGER_LITERAL
            else None
        )
        if value is None or value != value or value in (float("inf"), float("-inf")):
            return None
        return -value if signed else value

    shown = "".join(tok.raw_text for tok in toks[start:]).replace(",", ", ")
    first, last = toks[0], toks[-1]
    if fn == "choose":
        index = literal(args[0]) if len(args) > 1 else None
        choices = len(args) - 1
        return (
            _NullChoice(
                f"{shown} names no choice: its index {js_number_to_string(index)} is not 1 to "
                f"{choices}",
                first,
                last,
            )
            if index is not None and (index < 1 or index > choices)
            else None
        )
    if fn == "switch":
        conditions = [arg for i, arg in enumerate(args) if i % 2 == 0]
        return (
            _NullChoice(f"{shown} has no condition that is True", first, last)
            if len(args) % 2 == 0 and all(literal(arg) == 0 for arg in conditions)
            else None
        )
    condition = literal(args[0]) if len(args) == 3 else None
    chosen = None if condition is None else args[2 if condition == 0 else 1]
    return (
        _NullChoice(f"{shown} takes the branch that is Null", first, last)
        if chosen is not None and len(chosen) == 1 and token_text(chosen[0]) == "null"
        else None
    )


@dataclass(frozen=True, slots=True)
class _ArrayValue:
    """An array value, by its element type, or an Empty Variant."""

    element: str
    text: str


@dataclass(frozen=True, slots=True)
class _ArrayProblem:
    code: str  # "arrayTargetAssignment" | "assignmentTypeMismatch"
    message: str
    span: Span


def _array_assignment_problem(
    name: str,
    name_span: Span,
    value_tokens: Sequence[VbaToken],
    base_offset: int,
    target_shape: DeclaredValueShape | None,
    source_shape: Callable[[str], DeclaredValueShape | None],
    variant_value: Callable[[str], _ArrayValue | None],
    scalar_type: Callable[[list[VbaToken]], InferredArgumentType | None],
    source_names: SourceNameScope,
    returns_variant_array: Callable[[str], bool] = lambda _name: False,
    member_ctx: MemberCompletionContext | None = None,
    source: str | None = None,
    array_failure_phase: str = "compile",
    array_element_identity: Callable[[str | None], str] | None = None,
) -> _ArrayProblem | None:
    """What an array target takes, and what an array value goes into (XLIDE
    issue #194, each measured in Excel 16.0):

     - A fixed array takes no assignment, and a dynamic array takes no scalar:
       `a = b` into `Dim a(1)`, `a = "abc"`, `a = 5`, `a = Join(...)` do not
       compile ("Can't assign to array"). A Byte array takes a String.
     - A dynamic array takes an array variable of its own element type only:
       Long() from Integer(), and Variant() from Long(), do not compile.
     - Array() gives Variant() and Split and Filter give String(), so either
       into another element type raises 13: `Dim a() As String: a = Array("x")`.
       So does an Empty Variant, and so does an array into a scalar.

    `name`, `name_span` and `value_tokens` are upstream's `assignment` object."""
    value = unwrap_outer_parens(_non_comment(value_tokens))
    if not value:
        return None
    value_span = Span(base_offset + value[0].start, base_offset + value[-1].end)
    shown = value[0].raw_text if len(value) == 1 else "this value"
    value_name = token_name(value[0]) if len(value) == 1 else None
    named = source_shape(value_name) if value_name else None
    # Only into an array: into a scalar, Empty would run.
    called = (
        _ArrayValue("variant", f"{value[0].raw_text}(...), which returns Array(...) or Empty,")
        if target_shape is not None
        and target_shape.is_array
        and _is_whole_call(value)
        and returns_variant_array(value[0].raw_text)
        else None
    )
    produced = _array_produced_by(value, source_names)
    if produced is None:
        produced = called
    if produced is None and value_name and not (named is not None and named.is_array):
        produced = variant_value(value_name)
    identity = array_element_identity or _element_type
    last = value[-1]
    member_name = token_name(last)
    member = resolve_exact_member_completion(source, member_name, base_offset + last.end, member_ctx) if source and member_ctx and member_name and len(value) >= 2 and value[-2].raw_text == "." else None
    target_type = identity(target_shape.as_type if target_shape is not None else None)
    if target_shape is not None and target_shape.is_array:
        element_text = _TRAILING_PARENS_RE.sub(
            "", target_shape.as_type if target_shape.as_type is not None else "Variant", count=1
        )
        elements = f"an array of {element_text}"

        def cannot(what: str) -> _ArrayProblem:
            return _ArrayProblem(
                "arrayTargetAssignment",
                f"Can't assign to array: '{name}' is {elements}, and {what}. This is a VBE "
                "compile error.",
                name_span,
            )

        if target_shape.is_fixed_array:
            return cannot("a fixed-size array takes no assignment whole")
        if named is not None and named.is_array:
            return (
                None
                if identity(named.as_type) == target_type
                else cannot(
                    f"'{value_name}' is an array of "
                    f"{named.as_type if named.as_type is not None else 'Variant'}"
                )
            )
        # A Function declared to return a typed array is held to the same
        # rule as an array variable: `a = StrArr()` into Long() or Variant()
        # does not compile (XLIDE issue #222, measured in Excel 16.0).
        if member is not None and member.is_array:
            element = member.returns
            return None if identity(element) == target_type else cannot(f"'{member.owner}.{member.name}' is an array of {element or 'Variant'}")
        returned = scalar_type(value) if produced is None else None
        if returned is not None and _ENDS_IN_PARENS_RE.search(returned.type_):
            element = _TRAILING_PARENS_RE.sub("", returned.type_, count=1)
            return (
                None
                if identity(element) == target_type
                else cannot(f"{returned.label} returns an array of {element}")
            )
        if produced is not None:
            if produced.element == target_type:
                return None
            what = (
                produced.text
                if produced.element == "empty"
                else f"{produced.text} holds "
                f"{'String' if produced.element == 'string' else 'Variant'} elements"
            )
            return _ArrayProblem(
                "assignmentTypeMismatch",
                f"Assignment to '{name}' expects {elements}, but {what}. This will raise "
                "Run-time error '13': Type mismatch.",
                value_span,
            )
        # A scalar is a literal, or a VBA function that returns one: an
        # expression's inferred type can miss an array (a UDT field, a Function
        # returning Byte()).
        literal = len(value) == 1 and value[0].kind in (
            TokenKind.STRING_LITERAL,
            TokenKind.INTEGER_LITERAL,
            TokenKind.FLOAT_LITERAL,
        )
        runtime_call = _scalar_runtime_call(value, source_names)
        scalar = scalar_type(value) if literal or runtime_call else None
        scalar_kind = normalize_type(scalar.type_ if scalar is not None else None)
        if (
            scalar is not None
            and scalar_kind
            and not _ENDS_IN_PARENS_RE.search(scalar.type_)
            and is_known_scalar_type(scalar_kind)
            and not (target_type == "byte" and scalar_kind == "string")
        ):
            return cannot(f"{shown} is a {scalar.type_}, not an array")
        return None
    if target_shape is not None and is_known_scalar_type(target_type):
        typed_call = scalar_type(value) if value[-1].raw_text == ")" else None
        array_type = named.as_type if named is not None and named.is_array else member.returns if member is not None and member.is_array else typed_call.type_ if typed_call else None
        if target_type == "string" and _element_type(array_type) == "byte":
            return None
        if (named is not None and (named.is_array or _ENDS_IN_PARENS_RE.search(named.as_type or ""))) or (member is not None and (member.is_array or _ENDS_IN_PARENS_RE.search(member.returns or ""))) or (produced is None and typed_call and _ENDS_IN_PARENS_RE.search(typed_call.type_)):
            phase = "This is a VBE compile error." if array_failure_phase == "compile" else "This will raise Run-time error '13': Type mismatch."
            return _ArrayProblem("arrayAssignmentToScalar" if array_failure_phase == "compile" else "assignmentTypeMismatch", f"Type mismatch: assignment to '{name}' expects {target_shape.as_type}, but this value is a whole typed array. {phase}", value_span)
    if (
        produced is not None
        and produced.element != "empty"
        and target_shape is not None
        and is_known_scalar_type(target_type)
    ):
        return _ArrayProblem(
            "assignmentTypeMismatch",
            f"Assignment to '{name}' expects {_js_text(target_shape.as_type)}, but "
            f"{produced.text} is an array. This will raise Run-time error '13': Type mismatch.",
            value_span,
        )
    return None


def _array_only_variant_functions(
    source: str,
    mod: ModuleNode,
    activity: ConditionalActivityTracker | None,
) -> set[str]:
    """The module's Functions declared As Variant whose every assignment to their
    own name is `Array(...)`: each returns an array of Variant, or Empty when no
    assignment runs. Into an array of another element type both raise 13 (XLIDE
    issue #222, measured in Excel 16.0: `a = VarArr()` into Long())."""
    out: set[str] = set()
    for member in active_module_members(mod, activity):
        if (
            not isinstance(member, ProcedureNode)
            or member.proc_kind is not ProcKind.FUNCTION
            or member.type_suffix
        ):
            continue
        returns = normalize_type(member.return_type)
        if returns is not None and returns != "variant":
            continue
        lower = member.name.lower()
        only_arrays = True

        def names_self(tok: VbaToken) -> bool:
            tok_name = token_name(tok)
            return tok_name is not None and tok_name.lower() == lower

        def visit(stmt: LeafStatementNode) -> None:
            nonlocal only_arrays
            if not only_arrays:
                return
            for span in statement_and_branch_spans(stmt):
                toks = statement_tokens(source, span)
                # A one-line If's own span runs to its branches, which come next
                # as spans of their own; only its condition is read here.
                branches = isinstance(stmt, StatementNode) and stmt.single_line_if_branches
                if branches and span is stmt.span:
                    then = next(
                        (index for index, tok in enumerate(toks) if token_text(tok) == "then"), -1
                    )
                    toks = toks[:then] if then >= 0 else toks
                    if any(names_self(tok) for tok in toks):
                        only_arrays = False
                        return
                    continue
                if not any(names_self(tok) for tok in toks):
                    continue
                target = bare_assignment_target(source, span)
                # A one-line If's Then branch runs to its Else.
                significant = _non_comment(target[2]) if target is not None else []
                value = [
                    tok
                    for index, tok in enumerate(significant)
                    if not (index == len(significant) - 1 and token_text(tok) == "else")
                ]
                is_array_call = (
                    token_text(_at(value, 0)) == "array"
                    and _raw_at(value, 1) == "("
                    and match_paren_from(value, 1) == len(value) - 1
                )
                reads_self = any(names_self(tok) for tok in value)
                if target is None or target[0].lower() != lower or not is_array_call or reads_self:
                    only_arrays = False
                    return

        for_each_statement(member.body, visit, activity)
        if only_arrays:
            out.add(lower)
    return out


def _project_type_name_lookup(
    member_ctx: MemberCompletionContext,
    kind: str,
    case_sensitive: bool,
) -> Callable[[str], bool]:
    """Query one project kind without repeating scans, preserving early matches."""
    names: set[str] = set()
    next_index = 0

    def lookup(name: str) -> bool:
        nonlocal next_index
        query = name if case_sensitive else name.lower()
        if query in names:
            return True
        surfaces = member_ctx.project_class_members or ()
        # Metadata is stable within this rule pass. Resume after the last examined
        # surface; a successful query need not inspect the remaining project.
        while next_index < len(surfaces):
            surface = surfaces[next_index]
            next_index += 1
            if surface.kind != kind:
                continue
            declared = surface.name if case_sensitive else surface.name.lower()
            names.add(declared)
            if declared == query:
                return True
        return False

    return lookup


def _is_whole_call(value: Sequence[VbaToken]) -> bool:
    """Whether the value is one call and nothing more: `F()`, `F(1, 2)`."""
    return (
        token_name(_at(value, 0)) is not None
        and _raw_at(value, 1) == "("
        and match_paren_from(value, 1) == len(value) - 1
    )


def _element_type(as_type: str | None) -> str:
    """An array's element type, normalized: "Byte()" and "Byte" are both byte."""
    normalized = normalize_type(
        _TRAILING_PARENS_RE.sub("", as_type, count=1) if as_type is not None else None
    )
    return normalized if normalized is not None else "variant"


def _scalar_runtime_call(value: Sequence[VbaToken], source_names: SourceNameScope) -> bool:
    """Whether the value is one call to a VBA runtime function that returns a scalar: `Join(...)`."""
    index = 0
    if token_text(_at(value, 0)) == "vba" and _raw_at(value, 1) == ".":
        index = 2
    name = token_name(_at(value, index))
    paren = index + 1
    if _raw_at(value, paren) == "$":
        paren += 1
    if not name or _raw_at(value, paren) != "(" or match_paren_from(value, paren) != len(value) - 1:
        return False
    if index == 0 and runtime_callable_source_shadowed(name, source_names):
        return False
    runtime = resolve_runtime_function(name)
    returns = normalize_type(runtime.returns if runtime is not None else None)
    return returns is not None and is_known_scalar_type(returns)


def _array_produced_by(value: Sequence[VbaToken], source_names: SourceNameScope) -> _ArrayValue | None:
    """The array a call returns, by element type: Array() is Variant(), Split and Filter are String()."""
    index = 0
    if token_text(_at(value, 0)) == "vba" and _raw_at(value, 1) == ".":
        index = 2
    callee = token_text(_at(value, index))
    if _raw_at(value, index + 1) != "(" or match_paren_from(value, index + 1) != len(value) - 1:
        return None
    if index == 0 and runtime_callable_source_shadowed(value[0].raw_text, source_names):
        return None
    text = "".join(tok.raw_text for tok in value[: index + 1]) + "(...)"
    if callee == "array":
        return _ArrayValue("variant", text)
    if callee in ("split", "filter"):
        return _ArrayValue("string", text)
    return None


def _array_assignment_to_scalar_source(
    name: str,
    value_tokens: list[VbaToken],
    base_offset: int,
    expected_type: str,
    shapes: Mapping[str, DeclaredValueShape],
    resolve_target_shape: Callable[[str], SourceDeclaredShape],
    resolve_source_shape: Callable[[str], SourceDeclaredShape],
) -> tuple[str, Span] | None:
    target_scalar_type = _known_scalar_assignment_target_type(
        name, expected_type, shapes, resolve_target_shape
    )
    if not target_scalar_type:
        return None
    if len(value_tokens) != 1:
        return None
    tok = value_tokens[0]
    source_name = token_name(tok)
    if not source_name:
        return None
    resolved = resolve_source_shape(source_name)
    source_shape = resolved.shape if resolved.resolved else shapes.get(source_name.lower())
    if source_shape is None or not source_shape.is_array:
        return None
    # VBA special case (MS-VBAL Let-statement rules): a Byte array is directly
    # assignable to a String scalar - the idiomatic encoding-conversion pattern
    # (`s = bytes`). Only Byte element types are exempt; every other element
    # type remains a compile error.
    if target_scalar_type == "string" and normalize_type(source_shape.as_type) == "byte":
        return None
    return (source_name, Span(base_offset + tok.start, base_offset + tok.end))


def _known_scalar_assignment_target_type(
    name: str,
    expected_type: str,
    shapes: Mapping[str, DeclaredValueShape],
    resolve_shape: Callable[[str], SourceDeclaredShape],
) -> str | None:
    resolved = resolve_shape(name)
    target_shape = resolved.shape if resolved.resolved else shapes.get(name.lower())
    if target_shape is not None and target_shape.is_array:
        return None
    as_type = target_shape.as_type if target_shape is not None else None
    normalized = normalize_type(as_type if as_type is not None else expected_type)
    return normalized if normalized and is_known_scalar_type(normalized) else None


def check_missing_return_assignments(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None,
    activity: ConditionalActivityTracker | None,
    module_name: str | None,
    implemented_interfaces: AbstractSet[str] | None,
    push: PushFn,
) -> None:
    """A Function/Property Get with no return assignment silently returns the default."""
    module_signatures = callable_type_signatures_for(symbols, project_procedures)
    is_interface = (
        module_name is not None
        and implemented_interfaces is not None
        and module_name.lower() in implemented_interfaces
    )
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        if member.proc_kind not in (ProcKind.FUNCTION, ProcKind.PROPERTY_GET):
            continue
        if not member.closed:
            continue
        if _procedure_has_return_assignment(source, member, activity, module_signatures):
            continue
        if _return_is_not_expected(source, member, activity, is_interface):
            continue
        proc_label = "Property Get" if member.proc_kind is ProcKind.PROPERTY_GET else "Function"
        push(
            "missingReturnAssignment",
            f"{proc_label} '{member.name}' has no return assignment; VBA will return the default "
            f"value. Assign to '{member.name}' before exit if a value is intended.",
            declared_name_span(source, member.span, member.name),
        )


def _procedure_has_return_assignment(
    source: str,
    proc: ProcedureNode,
    activity: ConditionalActivityTracker | None,
    module_signatures: Mapping[str, CallableTypeSignature],
) -> bool:
    lower = proc.name.lower()
    found = False

    def assigns_in(span: Span) -> bool:
        bare = bare_assignment_target(source, span)
        if bare is not None and bare[0].lower() == lower:
            return True
        set_target = set_assignment_target(source, span)
        if set_target is not None and set_target[0].lower() == lower:
            return True
        if _return_assigned_by_statement_form(source, span, lower):
            return True
        call = extract_call(source, span)
        qualified = None if call else extract_qualified_call(source, span, module_signatures)
        effective = call or qualified
        return effective is not None and _call_passes_name_to_by_ref_param(
            effective, lower, module_signatures
        )

    def visit(stmt: LeafStatementNode) -> None:
        nonlocal found
        if found:
            return
        # The branch spans cover a single-line If, whose statements the walk
        # itself does not reach (XLIDE issue #46).
        for span in statement_and_branch_spans(stmt):
            if assigns_in(span) or _assigns_own_field(source, span, lower):
                found = True
                return

    for_each_statement(proc.body, visit, activity)
    # `For Count3 = 1 To 3` assigns the return variable as its counter (XLIDE issue
    # #115): the loop is a block, not a statement the walk above visits.
    return found or _for_loop_assigns(proc.body, lower, activity)


def _for_loop_assigns(
    body: Sequence[BodyNode], lower: str, activity: ConditionalActivityTracker | None
) -> bool:
    return any(
        isinstance(node, ForBlockNode)
        and node.control_variable is not None
        and node.control_variable.lower() == lower
        for node in iter_body_nodes(body, inactive_node_skip(activity))
    )


_TYPE_SUFFIX_CHARS = frozenset("$%&!#@")


def _return_assigned_by_statement_form(source: str, span: Span, lower: str) -> bool:
    """The statement forms besides `Name = value` that assign a Function's return
    variable (XLIDE issue #115, each measured in Excel 16.0): `For Name = 1 To 3`
    leaves the counter's final value, `ReDim Name(2)` sizes an array return,
    `Line Input #f, Name`, `Input #f, Name` and `Get #f, 1, Name` read into it, and
    `Name$ = "hi"` names it with its type-declaration character."""
    toks = statement_tokens(source, span)
    i = first_executable_token_index(toks)

    def at(index: int) -> VbaToken | None:
        return toks[index] if 0 <= index < len(toks) else None

    def is_name(tok: VbaToken | None) -> bool:
        name = token_name(tok)
        return name is not None and name.lower() == lower

    head = token_text(at(i))
    if head == "for":
        return is_name(at(i + 2)) if token_text(at(i + 1)) == "each" else is_name(at(i + 1))
    if head == "redim":
        k = i + 1
        if token_text(at(k)) == "preserve":
            k += 1
        depth = 0
        while k < len(toks):
            raw = toks[k].raw_text
            if raw == "(":
                depth += 1
            elif raw == ")":
                depth -= 1
            elif (
                depth == 0
                and is_name(toks[k])
                and (k == i + 1 or token_text(toks[k - 1]) == "preserve" or toks[k - 1].raw_text == ",")
            ):
                return True
            k += 1
        return False
    if head in ("line", "input", "get"):
        # Everything after the file number is a target (Line Input / Input) or the
        # third slot is (Get #f, rec, var); a plain name in one of them is the
        # assignment.
        rest = toks[i + 1 :]
        return any(is_name(tok) and index > 0 and rest[index - 1].raw_text == "," for index, tok in enumerate(rest))
    # `Name$ = value`: the suffix is glued to the name and the `=` follows.
    name_tok = at(i)
    suffix = at(i + 1)
    equals = at(i + 2)
    return (
        is_name(name_tok)
        and name_tok is not None
        and suffix is not None
        and suffix.start == name_tok.end
        and len(suffix.raw_text) == 1
        and suffix.raw_text in _TYPE_SUFFIX_CHARS
        and equals is not None
        and equals.raw_text == "="
    )


def _assigns_own_field(source: str, span: Span, lower: str) -> bool:
    """True for `Name.Field = value`, which fills in a UDT or object return field by
    field.

    `utc_DateToSystemTime.utc_wYear = ...` IS the return assignment, and reading
    only bare `Name =` counted 16 such functions in the test corpus as never
    assigning anything.
    """
    toks = statement_tokens(source, span)
    i = first_executable_token_index(toks)
    if i < len(toks) and toks[i].raw_text.lower() == "let":
        i += 1
    name = toks[i] if i < len(toks) else None
    nxt = toks[i + 1] if i + 1 < len(toks) else None
    if name is None or name.raw_text.lower() != lower:
        return False
    if nxt is None or nxt.raw_text != ".":
        return False
    return any(
        tok.kind is TokenKind.OPERATOR and tok.raw_text == "=" for tok in toks[i + 2 :]
    )


_DECLARATION_ONLY_RE = re.compile(r"^(Dim|Const|Static|ReDim)\b", re.IGNORECASE | re.ASCII)
_ERR_RAISE_RE = re.compile(r"\bErr\s*\.\s*Raise\b", re.IGNORECASE)
_ERROR_STATEMENT_RE = re.compile(r"^Error\s", re.IGNORECASE)


def _return_is_not_expected(
    source: str,
    proc: ProcedureNode,
    activity: ConditionalActivityTracker | None,
    is_interface: bool,
) -> bool:
    """True for a procedure that is not expected to assign a return: an empty body
    in a CLASS another module implements (its members are stated for the
    implementer to fill in) or one whose work is to raise. An empty Function
    anywhere else is unfinished code and still reports.
    """
    executable = 0
    raises = False

    def visit(stmt: LeafStatementNode) -> None:
        nonlocal executable, raises
        if raises:
            return
        text = js_trim(source[stmt.span.start : stmt.span.end])
        if not text or text.startswith("'") or _DECLARATION_ONLY_RE.match(text):
            return
        executable += 1
        if _ERR_RAISE_RE.search(text) or _ERROR_STATEMENT_RE.match(text):
            raises = True

    for_each_statement(proc.body, visit, activity)
    return (executable == 0 and is_interface) or raises


def _call_passes_name_to_by_ref_param(
    call: CallArguments, lower_name: str, module_signatures: Mapping[str, CallableTypeSignature]
) -> bool:
    sig = callable_signature_for_call(call, module_signatures)
    if sig is None:
        return False
    positional_index = 0
    for slot in call.slots:
        named = named_argument_slot(slot)
        if named is not None:
            param = next(
                (p for p in sig.params if strip_header_brackets(p.name).lower() == named[0].lower()),
                None,
            )
            value_slot = named[1]
        else:
            param = sig.params[min(positional_index, len(sig.params) - 1)] if sig.params else None
            positional_index += 1
            value_slot = slot
        if param is None or not param.by_ref or not _single_slot_name_equals(value_slot, lower_name):
            continue
        return True
    return False


def _single_slot_name_equals(slot: list[VbaToken], lower_name: str) -> bool:
    toks = [t for t in slot if t.kind is not TokenKind.COMMENT and t.kind is not TokenKind.NEWLINE]
    return len(toks) == 1 and (token_name(toks[0]) or "").lower() == lower_name


_VARIANT_OR_OBJECT_RE = re.compile(r"(?:Variant|Object)\Z", re.IGNORECASE)


def _host_read_only_assignment_error(
    target: MemberCompletion,
    uses_set: bool,
    member_ctx: MemberCompletionContext,
) -> str | None:
    """The compile error the VBE gives an assignment to a read-only host property,
    or None when the assignment compiles or the models cannot say which error it
    is. Measured in Excel, Word and PowerPoint 16.0 (XLIDE issue #198):

    - A property of type Variant or Object takes either statement: the value goes
      to whatever the property returns when the code runs.
    - A Let to a scalar property is "Can't assign to read-only property", except
      on Excel's dispatch-only interfaces (Range, Shape, Font, ...), where it is
      "Wrong number of arguments or invalid property assignment", or "Assignment
      to constant not permitted" when the property takes parameters
      (Range.Address).
    - A Set to an object property is "Invalid use of property". A Let to one goes
      to the returned object's default member, so it is not decided here.
    - A Set to a scalar property gives the Let's error on a dispatch-only
      interface and "Invalid use of property" on a dual one, but "Type mismatch"
      when a dual property takes parameters (Word's Range.XML), and the Word,
      PowerPoint and Office models do not record a property's parameters. So a
      Set is judged on Excel's types only, whose parameters the model has.
    """
    if target.kind != "property":
        return None
    declared = js_trim(target.declared_type) if target.declared_type is not None else ""
    if not declared or _VARIANT_OR_OBJECT_RE.match(declared):
        return None
    scalar = (
        is_known_scalar_type(normalize_type(declared) or "")
        or resolve_host_enum(declared, member_ctx.model) is not None
    )
    dispatch_only = is_dispatch_only_host_type(target.owner, member_ctx.model)
    with_parameters = signature_declares_parameters(target.signature)
    if not scalar:
        return "Invalid use of property" if uses_set and target.returns and not with_parameters else None
    if dispatch_only:
        return (
            "Assignment to constant not permitted"
            if with_parameters
            else "Wrong number of arguments or invalid property assignment"
        )
    if not uses_set:
        return "Can't assign to read-only property"
    return "Invalid use of property" if target.owner.startswith("Excel.") else None


def _late_bound_receiver(source: str, offset: int, member_ctx: MemberCompletionContext) -> bool:
    """Whether the receiver of the member ending at `offset` binds only at run time."""
    receiver = resolve_receiver_type_at(source, offset, member_ctx)
    return receiver is None or is_late_bound_type_key(receiver)


def _push_object_assignment_mismatch(
    push: PushFn,
    label: str,
    expected: str | None,
    actual: InferredArgumentType | None,
    reason: str,
    fallback: Span,
    scalar_error: str,
) -> None:
    """A Set whose value cannot be the target's type. A scalar value is refused
    when the module compiles: "Type mismatch" into a variable (`Set r = 5`),
    "Object required" through a Property Set (`Set h.Item = 5`). An object of the
    wrong class compiles, and raises 13 when the Set runs, so under On Error
    Resume Next it is handled (XLIDE issue #202, measured in Excel 16.0)."""
    actual_label = _js_text(actual.label if actual is not None else None)
    span = actual.span if actual is not None else fallback
    if reason == SCALAR_OBJECT_ASSIGNMENT_REASON:
        push(
            "setRequiresObject",
            f"Set assigns an object to '{label}', which expects {_js_text(expected)}, but "
            f"{actual_label} is not an object. This is a VBE compile error: {scalar_error}.",
            span,
        )
        return
    push(
        "assignmentObjectTypeMismatch",
        f"Object assignment to '{label}' expects {_js_text(expected)}, but got {actual_label}. "
        f"{reason} This will raise Run-time error '13': Type mismatch.",
        span,
    )


_COLLECTION_COUNT_RE = re.compile(r"([A-Za-z][A-Za-z0-9_]*)\.count\Z", re.IGNORECASE)
_MSFORMS_PREFIX_RE = re.compile(r"MSForms\.", re.IGNORECASE)


def check_member_assignment_types(
    source: str,
    member: ProcedureNode,
    env: Mapping[str, str],
    module_signatures: Mapping[str, CallableTypeSignature],
    source_names: SourceNameScope,
    member_ctx: MemberCompletionContext,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    project_declares_collection: Callable[[], bool],
    is_form_owner: Callable[[str], bool],
    resolve_expression_type: SourceDeclaredTypeResolver | None = None,
    resolve_qualified_expression_type: SourceQualifiedDeclaredTypeResolver | None = None,
    symbols: ModuleSymbols | None = None,
    *,
    coercion_type: Callable[[str], str] | None = None,
    check_returned_object_default: Callable[[str, str, Span], bool] | None = None,
    array_value_at: Callable[[LeafStatementNode, str], _ArrayValue | None] = lambda _stmt, _name: None,
    project_visible_symbols: Sequence[VbaSymbol] | None = None,
    array_by_ref_identity: Callable[[str | None], str] = lambda value: _element_type(value),
) -> None:
    """Port of checkMemberAssignmentTypes: `obj.Member = value` type compatibility.

    The project-class checks read source-backed members, which carry
    writability; a host member is judged by its read-only contract and the
    values the host refuses."""
    project_classes = len(member_ctx.project_class_members or ()) > 0
    coerce = coercion_type or create_assignment_coercion_type(member_ctx)
    returned_default = check_returned_object_default or (lambda _type, _label, _span: False)
    values_at: Callable[[LeafStatementNode], Mapping[str, KnownLocalValue]] | None = None
    boolean_names: set[str] | None = None

    def infer(tokens: list[VbaToken], offset: int) -> InferredArgumentType | None:
        return infer_argument_type(
            tokens, offset, env, module_signatures, source_names,
            source=source, member_ctx=member_ctx,
            resolve_expression_type=resolve_expression_type,
            resolve_qualified_expression_type=resolve_qualified_expression_type,
        )

    def check_statement(span: Span, stmt: LeafStatementNode) -> None:
        # A local known to hold a number, for a host property's limits (XLIDE
        # issue #346). A Boolean holding True is no -1 there: Excel takes True
        # where it refuses -1, and refuses False as it does 0 (XLIDE issue #630,
        # measured in Excel 16.0).
        def known(tokens: Sequence[VbaToken]) -> float | None:
            nonlocal values_at, boolean_names
            value = _non_comment(tokens)
            value_name = token_name(value[0]) if len(value) == 1 else None
            lower = value_name.lower() if value_name is not None else None
            held: KnownLocalValue | None = None
            if lower and symbols is not None:
                if values_at is None:
                    values_at = known_local_literal_values_at(source, member, symbols, activity)
                held = values_at(stmt).get(lower)
            if held is None or held.kind != "number" or lower is None:
                return None
            # Any direct child with this exact declared type counts, including a
            # later duplicate. Held values still come from the current statement.
            if boolean_names is None:
                boolean_names = set()
                assert symbols is not None
                proc_sym = procedure_symbol_for(symbols, member)
                for child in (proc_sym.children if proc_sym is not None and proc_sym.children else []):
                    if child.as_type is not None and child.as_type.lower() == "boolean":
                        boolean_names.add(child.name.lower())
            boolean = lower in boolean_names
            if isinstance(held.value, str):
                return None
            return None if boolean and held.value != 0 else held.value

        assignment = _member_assignment_target(source, span)
        if symbols:
            setter = source_setter_assignment(source, span, symbols, procedure_symbol_for(symbols, member), None, member_ctx)
            if setter and invalid_setter_assignment_arity(setter, source, lambda *_: None):
                return
        if assignment is None:
            return
        value = _non_comment(assignment.value_tokens)
        value_span = (
            Span(span.start + value[0].start, span.start + value[-1].end) if value else None
        )
        # `c.Count = 2` on a local Collection: Count is a Long Function (XLIDE
        # issue #305, measured in Excel 16.0).
        collection_count = _COLLECTION_COUNT_RE.match(assignment.label)
        if (
            collection_count is not None
            and not assignment.with_arguments
            and normalize_type(env.get(collection_count.group(1).lower())) == "collection"
            and not project_declares_collection()
        ):
            push(
                "readonlyMemberAssignment",
                f"Cannot assign to '{assignment.label}': a Collection's Count is a Function "
                "returning Long. This is a VBE compile error: Function call on left-hand side "
                "of assignment must return Variant or Object.",
                assignment.member_span,
            )
            return
        target = resolve_exact_member_completion(
            source, assignment.member, assignment.member_span.end, member_ctx
        )
        if target and target.kind == "method" and not target.sub and not assignment.uses_set and not _late_bound_receiver(source, assignment.member_span.end, member_ctx) and _getter_may_return_object(target.returns or target.declared_type, member_ctx):
            total, _ = member_parameter_counts(target.signature)
            result_indexed = assignment.with_arguments and total == 0 and assignment.has_arguments
            params = member_callable_signature(target).params if target.signature or target.procedure_params else None
            if not result_indexed and params is not None and _invalid_getter_argument_count(source, target.name, assignment.member_span, params, assignment.argument_tokens, span.start):
                return
            returned = target.returns or target.declared_type
            if not result_indexed and returned and returned_default(returned, assignment.label, assignment.member_span):
                return
        if target is not None and target.access == "read-only" and target.writable is None:
            vbe_error = _host_read_only_assignment_error(target, assignment.uses_set, member_ctx)
            if vbe_error and not _late_bound_receiver(source, assignment.member_span.end, member_ctx):
                push(
                    "readonlyMemberAssignment",
                    f"Cannot assign to read-only property '{assignment.label}'. This is a VBE "
                    f"compile error: {vbe_error}.",
                    assignment.member_span,
                )
            elif target.kind == "property" and target.owner.lower() == "excel.range" and target.name.lower() in {"height", "width", "left", "top", "text", "countlarge", "hasarray", "hasformula"} and not _late_bound_receiver(source, assignment.member_span.end, member_ctx):
                alternative = " Use RowHeight to change it." if target.name.lower() == "height" else " Use ColumnWidth to change it." if target.name.lower() == "width" else ""
                push("hostReadonlyValueAssignment", f"Cannot assign to read-only property '{assignment.label}': it returns a value and has no setter. This assignment fails when it runs.{alternative}", assignment.member_span)
            return
        # `Range("A1").Formula = "=SUM(B1"`: a formula Excel cannot parse (XLIDE
        # issue #276). A warning: a cell formatted as Text takes it.
        formula = (
            formula_string_problem(target, assignment.value_tokens)
            if target is not None and not assignment.uses_set and not assignment.with_arguments
            else None
        )
        if formula and value_span is not None:
            push("formulaStringUnparsed", formula, value_span)
            return
        # `Range("A1").Font.Size = 500`: a value the host refuses (XLIDE issue #204).
        if (
            target is not None
            and target.writable is None
            and not assignment.uses_set
            and not assignment.with_arguments
        ):
            problem = host_property_value_problem(target, assignment.value_tokens, known)
            if problem and value_span is not None:
                push("hostPropertyValueOutOfRange", problem, value_span)
                return
        # `ActiveSheet.Visible = "abc"`: a receiver of several host types, each
        if target and target.kind == "property" and target.access == "read/write" and target.writable is None and not assignment.uses_set and not assignment.with_arguments and target.declared_type and (is_known_scalar_type(normalize_type(target.declared_type) or "") or resolve_host_enum(target.declared_type, member_ctx.model)) and not _late_bound_receiver(source, assignment.member_span.end, member_ctx):
            actual = infer(assignment.value_tokens, span.start)
            host_expected = "Long" if resolve_host_enum(target.declared_type, member_ctx.model) else target.declared_type
            host_array_problem = _array_assignment_problem(assignment.label, assignment.member_span, assignment.value_tokens, span.start, DeclaredValueShape(host_expected, False, False), lambda name: declared_shape_for_source_binding(symbols, procedure_symbol_for(symbols, member), project_visible_symbols, name, BareIdentifierContext.EXPRESSION).shape if symbols else None, lambda name: array_value_at(stmt, name), lambda _: actual, source_names, member_ctx=member_ctx, source=source, array_failure_phase="runtime")
            if host_array_problem:
                push(host_array_problem.code, host_array_problem.message, host_array_problem.span)
                return
            reason = incompatibility_reason(target.declared_type, actual) if actual else None
            if reason and actual:
                push("assignmentTypeMismatch", f"Assignment to '{assignment.label}' expects {target.declared_type}, but got {actual.label}. {reason}", actual.span)
                return

        # `ActiveSheet.Visible = "abc"`: a receiver of several host types, each
        # refusing the value alike (XLIDE issue #416).
        if (
            (target is None or "." not in target.owner)
            and (target is None or target.writable is None)
            and not assignment.uses_set
            and not assignment.with_arguments
        ):
            resolved = resolve_receiver_type_at(source, assignment.member_span.start, member_ctx)
            parts = (
                resolved[len("union:") :].split("|")
                if resolved is not None and resolved.startswith("union:")
                else []
            )
            union_problem = (
                host_union_property_value_problem(parts, assignment.member, assignment.value_tokens)
                if parts
                else None
            )
            if union_problem and value_span is not None:
                push("hostPropertyValueOutOfRange", union_problem, value_span)
                return
        # `Set f.T1 = Nothing`: a form's control is no variable to Set (XLIDE
        # issue #226, measured in Excel 16.0: "Invalid use of property").
        if (
            assignment.uses_set
            and not assignment.with_arguments
            and target is not None
            and target.writable is None
            and _MSFORMS_PREFIX_RE.match(target.returns or "")
            and is_form_owner(target.owner)
        ):
            push(
                "setRequiresObject",
                f"'{assignment.label}' is a control on the form {target.owner}, which no Set can "
                "replace. This is a VBE compile error: Invalid use of property.",
                assignment.member_span,
            )
            return
        # The project-class checks read a bare property target only. A Type's
        # array field takes an array, or a String As Byte: typeMembers.ts
        # judges it (XLIDE issue #417).
        indexed_accessor = target is not None and (target.set_accessor if assignment.uses_set else (target.let_accessor or (target.writable is False and target.signature is not None and (signature_declares_parameters(target.signature) or normalize_type(target.returns or target.declared_type) != "string"))))
        if (
            not project_classes
            or (assignment.with_arguments and not indexed_accessor)
            or target is None
            or target.writable is None
            or (target.is_array and not target.write_is_array)
        ):
            return
        if target.writable is False:
            if not assignment.uses_set and _getter_may_return_object(target.returns or target.declared_type, member_ctx):
                total, required = member_parameter_counts(target.signature)
                if (assignment.with_arguments and total == 0 and assignment.has_arguments) or (not assignment.with_arguments and required > 0):
                    return
                returned = target.returns or target.declared_type
                getter_params = (target.procedure_params or {}).get("propertyGet")
                if getter_params is not None and _invalid_getter_argument_count(source, target.name, assignment.member_span, member_callable_signature(target).params, assignment.argument_tokens, span.start):
                    return
                if returned and returned_default(returned, assignment.label, assignment.member_span):
                    return
                if normalize_type(returned) in (None, "variant") and target.known_value in ("scalar", "empty"):
                    push("variantValueMisuse", f"'{assignment.label}' has only a Property Get returning a Variant that holds no object. The Let writes through its returned value, which cannot receive a property assignment. This will raise Run-time error '424': Object required.", assignment.member_span)
                return
            push(
                "readonlyMemberAssignment",
                f"Cannot assign to read-only property '{assignment.label}'.",
                assignment.member_span,
            )
            return
        if not assignment.uses_set and target.write_is_array:
            actual = infer(assignment.value_tokens, span.start)
            _check_array_setter_value(assignment.label, target.write_type, assignment.value_tokens, span.start, actual, coerce, resolve_expression_type, resolve_qualified_expression_type, push, source_names, project_declares_collection, lambda type_: object_let_assignment_verdict(type_, member_ctx), lambda name: array_value_at(stmt, name), array_by_ref_identity)
            return
        declared_expected = target.write_type if target.write_type is not None else target.returns
        expected = coerce(declared_expected) if declared_expected else None
        if assignment.uses_set:
            # A Property Set takes the Set whatever the Let and Get are typed:
            # `Set c.M = New Collection` runs beside a Long Let (XLIDE issue #414).
            if (
                expected
                and is_known_scalar_type(normalize_type(expected) or "")
                and not target.set_accessor
            ):
                push(
                    "setRequiresObject",
                    f"Set assignment requires an object-valued target, but '{assignment.label}' "
                    f"expects {expected}.",
                    assignment.member_span,
                )
                return
            # `Set h.Item = x` needs a Property Set; with only a Property Let the
            # VBE refuses it, "Invalid use of property" (XLIDE issue #107).
            if target.let_accessor and not target.set_accessor:
                push(
                    "setRequiresObject",
                    f"Set assignment to '{assignment.label}' needs a Property Set, but the property "
                    "declares only a Property Let. This is a VBE compile error: Invalid use of property.",
                    assignment.member_span,
                )
                return
            actual = infer(assignment.value_tokens, span.start)
            reason = object_assignment_incompatibility_reason(expected, actual, member_ctx)
            if reason:
                _push_object_assignment_mismatch(
                    push, assignment.label, expected, actual, reason, assignment.member_span,
                    "Object required",
                )
            return
        # A bare `=` to a project property calls its Property Let, whatever the
        # value's type: `h.Item = New Collection` compiles with `Property Let
        # Item(ByVal v As Object)` (XLIDE issue #107). Only a property with a Set
        # and no Let refuses it: "Invalid use of property".
        if target.set_accessor and not target.let_accessor:
            push(
                "setRequired",
                f"Assignment to '{assignment.label}' requires Set: the property declares a Property "
                "Set and no Property Let. This is a VBE compile error: Invalid use of property.",
                assignment.member_span,
            )
            return
        if not target.let_accessor and is_known_object_assignment_type_ctx(expected, member_ctx):
            push(
                "setRequired",
                f"Object assignment to '{assignment.label}' requires Set because it expects {expected}.",
                assignment.member_span,
            )
            return
        if not expected or not is_known_scalar_type(normalize_type(expected) or ""):
            return  # a Let of an object or unknown type: nothing provable about the value
        actual = infer(assignment.value_tokens, span.start)
        array_problem = _array_assignment_problem(assignment.label, assignment.member_span, assignment.value_tokens, span.start, DeclaredValueShape(expected, False, False), lambda name: declared_shape_for_source_binding(symbols, procedure_symbol_for(symbols, member), project_visible_symbols, name, BareIdentifierContext.EXPRESSION).shape if symbols else None, lambda name: array_value_at(stmt, name), lambda _: actual, source_names, member_ctx=member_ctx, source=source)
        if array_problem:
            push(array_problem.code, array_problem.message, array_problem.span)
            return
        string_arithmetic = nonnumeric_string_arithmetic_operand(
            expected, assignment.value_tokens, span.start
        )
        if string_arithmetic is not None:
            push(
                "stringArithmeticCoercion",
                f"Assignment to '{assignment.label}' expects {expected}, but this numeric expression "
                f"contains {string_arithmetic.label}. This will raise Run-time error '13': "
                "Type mismatch.",
                string_arithmetic.span,
            )
            return
        actual = infer(assignment.value_tokens, span.start)
        if actual is None:
            return
        reason = incompatibility_reason(expected, actual)
        if not reason:
            return
        push(
            "assignmentTypeMismatch",
            f"Assignment to '{assignment.label}' expects {declared_expected}, but got {actual.label}. {reason}",
            actual.span,
        )

    # This rule reads a statement structurally - what precedes its first `=` is
    # the target - so it takes a single-line If's branches as statements of their
    # own. Read whole, `If ok Then w.Part = 1` had the target `If ok Then w.Part`,
    # and `If w.Part = 1 Then Exit Sub`, which assigns nothing, had `If w.Part`.
    def visit(stmt: LeafStatementNode) -> None:
        for span in statement_and_branch_spans(stmt):
            check_statement(span, stmt)

    for_each_statement(member.body, visit, activity)


@dataclass(frozen=True, slots=True)
class _ElementTarget:
    name: str
    label: str
    span: Span
    value_tokens: list[VbaToken]
    uses_set: bool


def _array_element_target(
    source: str,
    span: Span,
    symbols: ModuleSymbols,
    proc_sym: VbaSymbol | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
) -> _ElementTarget | None:
    """`v(0) = x` or `Set v(0) = x` where v is a declared array: the element as an
    assignment target, with the array's name, the span from the name to the
    closing parenthesis, and the value."""
    toks = statement_tokens(source, span)
    i = first_executable_token_index(toks)
    head = token_text(_at(toks, i))
    if head in ("set", "let"):
        i += 1
    name = token_name(_at(toks, i))
    if not name or _raw_at(toks, i + 1) != "(":
        return None
    close = match_paren_from(toks, i + 1)
    if close < 0 or _raw_at(toks, close + 1) != "=" or close + 2 >= len(toks):
        return None
    shape = declared_shape_for_source_binding(
        symbols, proc_sym, project_visible_symbols, name, BareIdentifierContext.ASSIGNMENT_TARGET
    )
    if not shape.resolved or shape.shape is None or shape.shape.is_array is not True:
        return None
    return _ElementTarget(
        name=name,
        label="".join(tok.raw_text for tok in toks[i : close + 1]),
        span=Span(span.start + toks[i].start, span.start + toks[close].end),
        value_tokens=list(toks[close + 2 :]),
        uses_set=head == "set",
    )


_CONTROL_CLASS_RE = re.compile(
    r"(?:msforms\.)?(textbox|label|listbox|combobox|checkbox|optionbutton|togglebutton|"
    r"commandbutton|frame|multipage|tabstrip|scrollbar|spinbutton|image)\Z",
    re.IGNORECASE,
)


def check_set_assignments(
    source: str,
    symbols: ModuleSymbols,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    member_ctx: MemberCompletionContext,
    push: PushFn,
    activity: ConditionalActivityTracker | None = None,
) -> ProcedureStatementVisitor:
    """`Set x = ...` where x is a declared scalar requires an object variable; a Set to
    an object target of a provably-incompatible object type is reported too."""
    is_document_module = _project_type_name_lookup(member_ctx, "document", False)
    is_project_class = _project_type_name_lookup(member_ctx, "class", False)
    resolve_object_type = create_object_assignment_type_resolver(member_ctx)
    share_interfaces = create_project_interface_sharing_lookup(member_ctx)
    # Form metadata is stable within this rule invocation; query only the names
    # actually used, retaining the first matching control and missing results.
    form_resolved = False
    current_form: VbaProjectClassMembers | None = None
    controls: dict[str, VbaProjectClassMember | None] = {}

    def form_control(lower: str) -> VbaProjectClassMember | None:
        nonlocal form_resolved, current_form
        if not member_ctx.me_project_type:
            return None
        if not form_resolved:
            form_name = member_ctx.me_project_type.lower()
            current_form = next(
                (
                    t
                    for t in member_ctx.project_class_members or ()
                    if t.kind == "userform" and t.name.lower() == form_name
                ),
                None,
            )
            form_resolved = True
        if current_form is None:
            return None
        if lower not in controls:
            controls[lower] = next(
                (
                    m
                    for m in current_form.members
                    if m.name.lower() == lower and _MSFORMS_PREFIX_RE.match(m.returns or "")
                ),
                None,
            )
        return controls[lower]

    module_signatures = build_module_type_signatures(symbols)
    module_declared_names: set[str] | None = None

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        env = type_environment_for(symbols, member)
        source_names = source_name_scope_for(symbols, member, project_visible_symbols)
        proc_sym = procedure_symbol_for(symbols, member)
        procedure_declared_names: set[str] | None = None
        resolvers = source_binding_type_resolvers(symbols, proc_sym, project_visible_symbols)
        resolve_expression_type = resolvers.resolve_expression_type
        resolve_qualified_expression_type = resolvers.resolve_qualified_expression_type
        # What an Object or Variant local holds at a statement (XLIDE issue #246).
        held_at: Callable[[BodyNode], HeldObjects] | None = None

        def visitor(stmt: LeafStatementNode) -> None:
            for span in statement_and_branch_spans(stmt):
                check_set_span(span, stmt)

        def check_set_span(span: Span, stmt: LeafStatementNode) -> None:
            nonlocal procedure_declared_names, module_declared_names, held_at
            # `Set v(0) = x` into an element of a declared array is judged as a
            # Set into a variable of its element type (XLIDE issue #306).
            element = _array_element_target(
                source, span, symbols, proc_sym, project_visible_symbols
            )
            set_target = set_assignment_target(source, span)
            if set_target is not None:
                target_name, target_span, value = set_target
                is_element = False
            elif element is not None and element.uses_set:
                target_name, target_span, value = element.name, element.span, element.value_tokens
                is_element = True
            else:
                return
            target_declared_type = declared_type_for_source_binding(
                symbols,
                proc_sym,
                project_visible_symbols,
                target_name,
                BareIdentifierContext.ASSIGNMENT_TARGET,
            )
            # `Set Sheet1 = Nothing`: a document's name is no variable to Set
            # (XLIDE issue #225, measured in Excel 16.0).
            if not target_declared_type.resolved and is_document_module(target_name):
                push(
                    "setRequiresObject",
                    f"'{target_name}' names the document module itself, which no Set can "
                    "replace. This is a VBE compile error: Invalid use of property.",
                    target_span,
                )
                return
            # `Set Answer = Nothing` inside the form: a control is no variable to
            # Set (XLIDE issue #315, measured in Excel 16.0).
            lower_target = target_name.lower()
            # Only a form context needs this exact direct-declaration shadow check.
            # Keep it local to this invocation: project names and enum members in
            # the broader runtime shadow scope do not count as declarations here.
            declared_here = False
            if member_ctx.me_project_type:
                if procedure_declared_names is None:
                    procedure_declared_names = {
                        symbol.name.lower()
                        for symbol in (proc_sym.children if proc_sym is not None and proc_sym.children else [])
                    }
                declared_here = lower_target in procedure_declared_names
                if not declared_here:
                    if module_declared_names is None:
                        module_declared_names = {
                            symbol.name.lower() for symbol in symbols.root.children or ()
                        }
                    declared_here = lower_target in module_declared_names
            control = None if declared_here else form_control(lower_target)
            if control is not None:
                push(
                    "setRequiresObject",
                    f"'{target_name}' is a control on this form, which no Set can replace. This "
                    "is a VBE compile error: Invalid use of property.",
                    target_span,
                )
                return
            expected = (
                target_declared_type.as_type
                if target_declared_type.resolved
                else env.get(target_name.lower())
            )
            target_type = normalize_type(expected)
            # `Set t = Prompt` with t As MSForms.TextBox and Prompt a Label on this
            # form (XLIDE issue #315, measured in Excel 16.0: 13).
            control_match = _CONTROL_CLASS_RE.match(js_trim(expected) if expected is not None else "")
            control_class = control_match.group(1).lower() if control_match is not None else None
            # Both target extractors slice significant statement tokens.
            value_name = token_name(value[0]) if len(value) == 1 else None
            value_control = (
                form_control(value_name.lower())
                if control_class and value_name and value_name.lower() not in env
                else None
            )
            value_class = (
                value_control.returns[len("MSForms.") :].lower()
                if value_control is not None and value_control.returns is not None
                else None
            )
            if value_control is not None and value_class and value_class != control_class:
                push(
                    "assignmentObjectTypeMismatch",
                    f"Object assignment to '{target_name}' expects {_js_text(expected)}, but "
                    f"'{value[0].raw_text}' is a control of class {value_control.returns}. This "
                    "will raise Run-time error '13': Type mismatch.",
                    Span(span.start + value[0].start, span.start + value[0].end),
                )
                return
            # `Set v = 5` is refused whatever v is: a literal is never an object
            # reference ("Object required", XLIDE issue #125, measured in Excel 16.0).
            if (
                (not target_type or target_type == "variant")
                and len(value) == 1
                and _is_scalar_literal_token(value[0])
            ):
                push(
                    "setRequiresObject",
                    f"Set assigns an object reference, but {value[0].raw_text} is a literal value. "
                    "This is a VBE compile error: Object required.",
                    Span(span.start + value[0].start, span.start + value[0].end),
                )
                return
            if not target_type or not is_known_scalar_type(target_type):
                if not resolve_object_type(expected):
                    return
                actual = infer_argument_type(
                    value, span.start, env, module_signatures, source_names,
                    source=source, member_ctx=member_ctx,
                    resolve_expression_type=resolve_expression_type,
                    resolve_qualified_expression_type=resolve_qualified_expression_type,
                )
                shown = actual
                reason = object_assignment_incompatibility_reason(
                    expected, actual, member_ctx, resolve_object_type, share_interfaces
                )
                value_span = (
                    Span(span.start + value[0].start, span.start + value[-1].end) if value else None
                )
                # `Set o = New Flat1` then `Set c = o`: the class an Object holds
                # is checked as the Set runs (XLIDE issue #246, measured in Excel 16.0).
                if not reason and len(value) == 1 and value_name and value_span is not None:
                    if held_at is None:
                        held_at = held_objects_at(source, member, symbols, activity)
                    held = held_at(stmt).classes.get(value_name.lower())
                    if held:
                        shown = InferredArgumentType(
                            type_=held,
                            label=f"'{value[0].raw_text}', which holds a {held} here",
                            span=value_span,
                        )
                        reason = object_assignment_incompatibility_reason(
                            expected, shown, member_ctx, resolve_object_type, share_interfaces
                        )
                # `Set c = ActiveSheet`: a Worksheet or a Chart, never a Collection
                # or a class of the project (XLIDE issue #306, measured in Excel
                # 16.0: 13).
                if (
                    not reason
                    and len(value) == 1
                    and value_span is not None
                    and token_text(value[0]) == "activesheet"
                    and "activesheet" not in source_names.runtime_shadows
                    and (
                        target_type == "collection"
                        or (target_type is not None and is_project_class(target_type))
                    )
                ):
                    shown = InferredArgumentType(
                        type_="Object",
                        label="ActiveSheet, a Worksheet or a Chart",
                        span=value_span,
                    )
                    reason = f"ActiveSheet holds a sheet, never a {_js_text(expected)}."
                sheets = (
                    None
                    if reason
                    else sheets_from_collection_property(value, expected, source_names, member_ctx)
                )
                if sheets is not None and value_span is not None:
                    shown = InferredArgumentType(
                        type_="Excel.Sheets",
                        label=f"'{sheets.text}', which returns a Sheets object",
                        span=value_span,
                    )
                    reason = (
                        "Excel's Worksheets and Charts properties return a Sheets object, never "
                        f"a {sheets.collection} one."
                    )
                if reason:
                    _push_object_assignment_mismatch(
                        push,
                        element.label if is_element and element is not None else target_name,
                        expected,
                        shown,
                        reason,
                        target_span,
                        "Type mismatch",
                    )
                return
            if is_element:
                return  # an element of a scalar array, which the VBE has not been asked about
            push(
                "setRequiresObject",
                f"Set assignment requires an object variable, but '{target_name}' is declared as "
                f"{_js_text(expected)}.",
                target_span,
            )

        return visitor

    return factory


_TYPE_CHAR_SUFFIX = re.compile(r"[$%&!#@]\Z")


def _mid_base_word(tok: VbaToken | None) -> str:
    """Suffix-stripped, lower-cased word for a token (keyword or identifier)."""
    if tok is None:
        return ""
    text = token_name(tok)
    if text is None:
        text = tok.raw_text
    return _TYPE_CHAR_SUFFIX.sub("", text.lower())


def check_mid_statement_literal_target(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """The target of a `Mid`/`Mid$`/`MidB`/`MidB$` replacement statement (MS-VBAL
    5.4.3.4) must be a writable String variable; a literal target is a compile
    error. The rule stays silent for a module that names `mid`/`midb` in any
    declaration form."""
    if _module_shadows_mid_intrinsic(symbols) or _module_redim_declares_mid_intrinsic(
        source, mod, activity
    ):
        return

    def visit(stmt: LeafStatementNode) -> None:
        hit = _mid_statement_literal_target_violation(source, stmt.span)
        if hit is not None:
            span, message = hit
            push("midStatementLiteralTarget", message, span)

    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            for_each_statement(member.body, visit, activity)


def _module_shadows_mid_intrinsic(symbols: ModuleSymbols) -> bool:
    """True when a module declares any symbol that shadows the Mid/MidB intrinsic."""
    return any(
        _TYPE_CHAR_SUFFIX.sub("", sym.name.lower()) in ("mid", "midb") for sym in symbols.all
    )


def _module_redim_declares_mid_intrinsic(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None
) -> bool:
    """True when a ReDim implicitly declares an array named mid/midb (absent from symbols)."""
    found = False

    def visit(stmt: LeafStatementNode) -> None:
        nonlocal found
        if found:
            return
        toks = statement_tokens_after_leading_label(source, stmt.span)
        if _mid_base_word(_at(toks, 0)) != "redim":
            return
        start = 2 if _mid_base_word(_at(toks, 1)) == "preserve" else 1
        for group in split_top_level_token_groups(toks, start, ","):
            if _mid_base_word(_at(group, 0)) in ("mid", "midb"):
                found = True
                return

    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            for_each_statement(member.body, visit, activity)
            if found:
                return True
    return False


_MID_LITERAL_KINDS = (TokenKind.STRING_LITERAL, TokenKind.INTEGER_LITERAL, TokenKind.FLOAT_LITERAL)


def _mid_statement_literal_target_violation(source: str, span: Span) -> tuple[Span, str] | None:
    toks = statement_tokens_after_leading_label(source, span)
    if not toks:
        return None
    # Strip a trailing type-character so both lexings of `Mid$` are handled: a
    # single `Mid$` token, or `Mid` followed by a separate `$` token (below).
    if _mid_base_word(toks[0]) not in ("mid", "midb"):
        return None
    paren_index = 1
    if _raw_at(toks, paren_index) == "$":
        paren_index = 2
    if _raw_at(toks, paren_index) != "(":
        return None
    close = match_paren_from(toks, paren_index)
    if close <= paren_index + 1:
        return None  # empty or unbalanced argument list
    # The Mid replacement-statement form: the matching `)` is followed by `=`.
    if _raw_at(toks, close + 1) != "=":
        return None
    arg_toks = toks[paren_index + 1 : close]
    slots = split_top_level_token_groups(arg_toks, 0, ",")
    target = slots[0] if slots else None
    # A number is no more a target than a string is: `Mid(5, 1) = "x"` is a
    # "Syntax error" as well (XLIDE issue #213, measured in Excel 16.0).
    if not target or len(target) != 1 or target[0].kind not in _MID_LITERAL_KINDS:
        return None  # target is not exactly one literal
    what = "string literal" if target[0].kind is TokenKind.STRING_LITERAL else "number"
    return (
        Span(span.start + target[0].start, span.start + target[0].end),
        "The target of a Mid statement must be a writable String variable, not a "
        f"{what}. Assigning into a literal is a compile error.",
    )


@dataclass(frozen=True, slots=True)
class _SpelledText:
    """The text a String expression spells out. `stand_in` marks a Date written as
    text, a Date literal or local or CStr of one, whose text the locale decides:
    it is never a number or a Boolean, so the text stands in for it only to say
    that (XLIDE issue #405, measured in Excel 16.0).

    `named` says what the text is, where the host or the locale decides the
    words: a month name, a type name, a cell address. Such a word is never a
    number, a Boolean or a date, and the text stands in for it (XLIDE issue #457,
    measured in Excel 16.0)."""

    text: str
    stand_in: bool
    named: str | None = None


_KnownLookup = Callable[[str], KnownLocalValue | None]

_CELL_RE = re.compile(r"[A-Z]{1,3}[1-9][0-9]*\Z")
_CELL_PARTS_RE = re.compile(r"([A-Z]+)([0-9]+)\Z")


def _literal_range_address(part: Sequence[VbaToken]) -> str | None:
    """`Range("c1").Address` as Excel gives it with no arguments: "$C$1", or "$A$1:$B$2"."""
    toks = _non_comment(part)
    if (
        len(toks) != 6
        or token_text(toks[0]) != "range"
        or toks[1].raw_text != "("
        or toks[2].kind is not TokenKind.STRING_LITERAL
        or toks[3].raw_text != ")"
        or toks[4].raw_text != "."
        or token_text(toks[5]) != "address"
    ):
        return None
    cells = string_literal_value(toks[2].raw_text).upper().split(":")
    if len(cells) > 2 or not all(_CELL_RE.match(cell) for cell in cells):
        return None
    return ":".join(_CELL_PARTS_RE.sub(r"$\1$\2", cell, count=1) for cell in cells)


MONTH_NAMES = (
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
)
DAY_NAMES = ("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
STRCONV_CASES: Mapping[str, int] = {"vbuppercase": 1, "vblowercase": 2, "vbpropercase": 3}

_DIGITS_RE = re.compile(r"[0-9]+\Z")
_PROPER_CASE_RE = re.compile(r"(^|[^a-z0-9])([a-z])")


def _exact(text: str) -> _SpelledText:
    return _SpelledText(text, False)


def _fixed_text_part(
    part: Sequence[VbaToken],
    known: _KnownLookup,
    env: Mapping[str, str],
    source_names: SourceNameScope,
    nesting: int,
) -> _SpelledText | None:
    """A call or property whose String result is fixed, or is a word no locale
    reads as a number, a Boolean or a date (XLIDE issue #457, each measured in
    Excel 16.0 into a Long, a Double, a Boolean and a Date): Chr, Hex, Oct, Space,
    String and StrConv over literals; MonthName, WeekdayName, Format of a Date
    literal as a month or day name, and TypeName; a Range's Address;
    Application.Name and PathSeparator. `Hex(9)` is "9", which converts."""

    def named(text: str, what: str) -> _SpelledText:
        return _SpelledText(text, False, what)

    first = token_text(_at(part, 0))
    if len(part) == 3 and first == "application" and part[1].raw_text == ".":
        member = token_text(part[2])
        if member == "name":
            return named("Microsoft Excel", "the application name")
        if member == "pathseparator":
            return named("\\", "the path separator")
        return None
    # `Range("A1").Address`, `Cells(1, 2).Address(False, False)`. The text
    # stands in with no digit: a string with one may be a number somewhere.
    if first in ("range", "cells") and _raw_at(part, 1) == "(":
        close = match_paren_from(part, 1)
        tail = part[close + 1 :]
        address = (
            len(tail) >= 2
            and tail[0].raw_text == "."
            and token_text(tail[1]) == "address"
            and (
                len(tail) == 2
                or (tail[2].raw_text == "(" and match_paren_from(tail, 2) == len(tail) - 1)
            )
        )
        return named("address", "a cell address") if close > 0 and address else None
    open_index = 2 if _raw_at(part, 1) == "$" else 1
    fn_name = token_name(_at(part, 0))
    fn = fn_name.lower() if fn_name is not None else None
    # `Split(Range("C1").Address, "$")(1)`, the column letter "C" (XLIDE issue #457).
    if (
        fn == "split"
        and _raw_at(part, open_index) == "("
        and not runtime_callable_source_shadowed(fn, source_names)
    ):
        close = match_paren_from(part, open_index)
        index = part[close + 1 :]
        groups = split_top_level_token_groups(part, open_index + 1, ",", close)
        text_group = groups[0] if groups else None
        separator = groups[1] if len(groups) > 1 else None
        rest = groups[2:]
        address_text = _literal_range_address(text_group) if text_group else None
        at = (
            js_number(index[1].raw_text)
            if len(index) == 3
            and index[0].raw_text == "("
            and index[1].kind is TokenKind.INTEGER_LITERAL
            and index[2].raw_text == ")"
            else None
        )
        if (
            address_text
            and separator is not None
            and len(separator) == 1
            and separator[0].kind is TokenKind.STRING_LITERAL
            and not rest
            and at is not None
        ):
            sep = string_literal_value(separator[0].raw_text)
            pieces = address_text.split(sep) if sep else None
            element = (
                pieces[int(at)]
                if pieces is not None and at == at and at.is_integer() and 0 <= at < len(pieces)
                else None
            )
            return None if element is None else _exact(element)
        return None
    if (
        not fn
        or runtime_callable_source_shadowed(fn, source_names)
        or _raw_at(part, open_index) != "("
        or match_paren_from(part, open_index) != len(part) - 1
    ):
        return None
    args = split_top_level_token_groups(part, open_index + 1, ",", len(part) - 1)

    def whole(k: int) -> int | None:
        arg = args[k] if k < len(args) else None
        return (
            int(arg[0].raw_text)
            if arg is not None
            and len(arg) == 1
            and arg[0].kind is TokenKind.INTEGER_LITERAL
            and _DIGITS_RE.match(arg[0].raw_text)
            else None
        )

    n = whole(0)
    if fn == "typename":
        return named("Integer", "a type name") if len(args) == 1 else None
    if fn == "monthname":
        return (
            named(MONTH_NAMES[n - 1], "a month name")
            if n is not None and 1 <= n <= 12 and len(args) <= 2
            else None
        )
    if fn == "weekdayname":
        return (
            named(DAY_NAMES[n - 1], "a day name")
            if n is not None and 1 <= n <= 7 and len(args) <= 3
            else None
        )
    if fn == "format":
        pattern = (
            string_literal_value(args[1][0].raw_text).lower()
            if len(args) == 2 and len(args[1]) == 1 and args[1][0].kind is TokenKind.STRING_LITERAL
            else None
        )
        date = len(args[0]) == 1 and args[0][0].kind is TokenKind.DATE_LITERAL
        if date and pattern in ("mmm", "mmmm"):
            return named("January", "a month name")
        if date and pattern in ("ddd", "dddd"):
            return named("Sunday", "a day name")
        return None
    if fn == "chr":
        return _exact(chr(n)) if len(args) == 1 and n is not None and 32 <= n <= 126 else None
    if fn in ("hex", "oct"):
        return (
            _exact(format(n, "X" if fn == "hex" else "o").upper())
            if len(args) == 1 and n is not None and n <= 2147483647
            else None
        )
    if fn == "space":
        return _exact(" " * n) if len(args) == 1 and n is not None and n <= 1000 else None
    if fn == "string":
        fill = (
            string_literal_value(args[1][0].raw_text)
            if len(args) > 1 and len(args[1]) == 1 and args[1][0].kind is TokenKind.STRING_LITERAL
            else ""
        )
        return (
            _exact(fill[0] * n)
            if len(args) == 2 and n is not None and n <= 1000 and fill != ""
            else None
        )
    if fn == "strconv":
        subject = (
            _spelled_text(args[0], known, env, source_names, False, nesting + 1)
            if len(args) == 2
            else None
        )
        kind: int | None = None
        if len(args) > 1 and len(args[1]) == 1:
            kind = whole(1)
            if kind is None:
                kind = STRCONV_CASES.get(token_text(args[1][0]))
        if subject is None or subject.stand_in or subject.named is not None:
            return None
        lower = subject.text.lower()
        if kind == 1:
            return _exact(subject.text.upper())
        if kind == 2:
            return _exact(lower)
        if kind == 3:
            return _exact(
                _PROPER_CASE_RE.sub(lambda m: m.group(1) + m.group(2).upper(), lower)
            )
        return None
    return None


# The text functions folded over known text: `Left("abc", 1)` is "a".
TEXT_FUNCTIONS: AbstractSet[str] = frozenset(
    {"cstr", "left", "right", "mid", "ucase", "lcase", "trim", "ltrim", "rtrim"}
)

_UNSIGNED_INTEGER_RE = re.compile(r"[0-9]+[%&]?\Z")


def _spelled_text(
    toks: Sequence[VbaToken],
    known: _KnownLookup,
    env: Mapping[str, str],
    source_names: SourceNameScope,
    whole: bool = False,
    nesting: int = 0,
) -> _SpelledText | None:
    """`"a" & "b"`, `Left("abc", 1)`, `o & 5` with o a Boolean known to be True:
    each part of a `&` chain a literal, a known local, or a text function over
    those. A whole number is written in digits and a Boolean as True or False."""
    # Recover conservatively on unfinished or deeply nested editor input, as the
    # expression parser and other recursive folders do. Siblings share the
    # current depth; only descending into a call argument consumes it.
    if nesting >= MAX_EXPRESSION_DEPTH:
        return None
    text = ""
    stand_in = False
    parts = split_top_level_token_groups(toks, 0, "&")
    # A number on its own is no text: only `&` makes one of it. `d = -657435` is
    # a Long, which Overflow judges (XLIDE issue #329). Inside `CStr(40000)` it is
    # the text the function makes (XLIDE issue #624).
    lone = unwrap_outer_parens(parts[0]) if whole and len(parts) == 1 else []
    if (
        0 < len(lone) <= 2
        and lone[-1].kind is TokenKind.INTEGER_LITERAL
        and (len(lone) == 1 or lone[0].raw_text == "-")
    ):
        return None
    for part in parts:
        spelled = _spelled_part(unwrap_outer_parens(part), known, env, source_names, nesting)
        if spelled is None:
            return None
        # A month name or address stands in only for itself: beside other text,
        # as in `"1 " & MonthName(1)`, it may make a date.
        if spelled.named is not None:
            return spelled if len(parts) == 1 else None
        text += spelled.text
        stand_in = stand_in or spelled.stand_in
    return _SpelledText(text, stand_in)


def _spelled_part(
    part: list[VbaToken],
    known: _KnownLookup,
    env: Mapping[str, str],
    source_names: SourceNameScope,
    nesting: int,
) -> _SpelledText | None:
    if len(part) == 2 and part[0].raw_text == "-" and _UNSIGNED_INTEGER_RE.match(part[1].raw_text):
        return _exact("-" + js_number_to_string(js_number(re.sub(r"[%&]\Z", "", part[1].raw_text))))
    if len(part) == 1:
        tok = part[0]
        if tok.kind is TokenKind.STRING_LITERAL:
            return _exact(string_literal_value(tok.raw_text))
        if tok.kind is TokenKind.INTEGER_LITERAL and _UNSIGNED_INTEGER_RE.match(tok.raw_text):
            return _exact(js_number_to_string(js_number(re.sub(r"[%&]\Z", "", tok.raw_text))))
        if tok.kind is TokenKind.DATE_LITERAL:
            return _SpelledText(tok.raw_text, True)
        word = token_text(tok)
        if word in ("true", "false"):
            return _exact("True" if word == "true" else "False")
        tok_name = token_name(tok)
        lower = tok_name.lower() if tok_name is not None else None
        if not lower:
            return None
        local_type = normalize_type(env.get(lower))
        if local_type == "date":
            return _SpelledText(f"[{tok.raw_text}]", True)
        value = known(lower)
        if value is not None and value.kind == "string" and not value.content_mutated:
            assert isinstance(value.value, str)
            return _exact(value.value)
        if value is not None and value.kind == "number" and local_type == "boolean":
            return _exact("False" if value.value == 0 else "True")
        if (
            value is not None
            and value.kind == "number"
            and local_type in ("byte", "integer", "long")
            and not isinstance(value.value, str)
            and float(value.value).is_integer()
        ):
            return _exact(js_number_to_string(value.value))
        return None
    fixed = _fixed_text_part(part, known, env, source_names, nesting)
    if fixed is not None:
        return fixed
    # `Left$` lexes as Left and a `$`.
    open_index = 2 if _raw_at(part, 1) == "$" else 1
    fn_name = token_name(_at(part, 0))
    fn = fn_name.lower() if fn_name is not None else None
    if (
        not fn
        or fn not in TEXT_FUNCTIONS
        or runtime_callable_source_shadowed(fn, source_names)
        or _raw_at(part, open_index) != "("
        or match_paren_from(part, open_index) != len(part) - 1
    ):
        return None
    args = split_top_level_token_groups(part, open_index + 1, ",", len(part) - 1)
    subject = _spelled_text(args[0], known, env, source_names, False, nesting + 1)
    # Only CStr passes a Date written as text on: Left of it depends on the locale.
    if subject is None or ((subject.stand_in or subject.named is not None) and fn != "cstr"):
        return None

    def count(k: int) -> int | None:
        arg = args[k] if k < len(args) else None
        return (
            int(arg[0].raw_text)
            if arg is not None and len(arg) == 1 and _DIGITS_RE.match(arg[0].raw_text)
            else None
        )

    s = subject.text
    if fn == "cstr":
        return subject if len(args) == 1 else None
    if fn == "ucase":
        return _exact(s.upper()) if len(args) == 1 else None
    if fn == "lcase":
        return _exact(s.lower()) if len(args) == 1 else None
    if fn == "trim":
        return _exact(s.strip(" ")) if len(args) == 1 else None
    if fn == "ltrim":
        return _exact(s.lstrip(" ")) if len(args) == 1 else None
    if fn == "rtrim":
        return _exact(s.rstrip(" ")) if len(args) == 1 else None
    if fn in ("left", "right"):
        n = count(1) if len(args) == 2 else None
        if n is None:
            return None
        return _exact(s[:n] if fn == "left" else s[max(0, len(s) - n) :])
    if fn == "mid":
        start = count(1)
        length = count(2) if len(args) == 3 else len(s) if len(args) == 2 else None
        if start is None or start < 1 or length is None:
            return None
        return _exact(s[start - 1 : start - 1 + length])
    return None
