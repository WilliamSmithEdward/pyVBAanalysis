"""Source Let/Set assignment index contracts, from XLIDE 11.1.0."""

from collections.abc import Sequence
from dataclasses import dataclass, field, replace

from ..completion.member_access import MemberCompletionContext, resolve_exact_member_completion
from ..identity_cache import IdentityLru
from ..lexer.token_helpers import token_name, tokens_without_leading_line_number
from ..lexer.token_kinds import TokenKind, VbaToken
from ..parser.nodes import ExprKind, Span
from ..parser.parse_expression import parse_expression
from ..symbols.name_resolution import BareIdentifierContext, BareIdentifierResolutionInput, BareIdentifierResolutionScope, resolve_bare_identifier_binding
from ..symbols.symbol_model import ModuleSymbols, VbaProjectClassMember, VbaProcedureParam, VbaSymbol, VbaSymbolKind, procedure_params_from_symbol
from .call_extraction import ArgSplit, CallableParamType, CallableTypeSignature, CallArguments, split_arg_slots, validate_arity
from .callable_signatures import is_by_ref_procedure_param
from .context import PushFn
from .model import VbaDiagnosticData
from .walker import statement_tokens_after_leading_label, top_level_operator_index


def assignment_target_name(target: Sequence[VbaToken]) -> tuple[int, bool] | None:
    index = len(target) - 1
    indexed = index >= 0 and target[index].raw_text == ")"
    if indexed:
        depth = 0
        while index >= 0:
            depth += 1 if target[index].raw_text == ")" else -1 if target[index].raw_text == "(" else 0
            index -= 1
            if depth == 0:
                break
    return (index, indexed) if index >= 0 else None


def assignment_target_from_tokens(statement: Sequence[VbaToken]) -> list[VbaToken] | None:
    tokens = tokens_without_leading_line_number(statement)
    if not tokens or tokens[-1].raw_text != "=":
        return None
    if tokens[0].raw_text.lower() == "if":
        branch = next((i for i in range(len(tokens) - 1, -1, -1) if tokens[i].kind is TokenKind.KEYWORD and tokens[i].raw_text.lower() in ("then", "else")), -1)
        if branch >= 0:
            tokens = tokens[branch + 1:]
    tokens = tokens[:-1]
    if tokens and tokens[0].raw_text.lower() == "let":
        tokens = tokens[1:]
    if not tokens or (tokens[0].kind not in (TokenKind.IDENTIFIER, TokenKind.BRACKETED_IDENTIFIER) and tokens[0].raw_text != "." and tokens[0].raw_text.lower() not in ("me", "thisworkbook")):
        return None
    syntax = [replace(tok, kind=TokenKind.IDENTIFIER) if tok.kind is TokenKind.KEYWORD and tok.raw_text.lower() == "me" else tok for tok in tokens]
    parsed = parse_expression(syntax)
    return tokens if parsed.expr is not None and parsed.end_index == len(tokens) and not parsed.diagnostics and parsed.expr.EXPR_KIND in (ExprKind.IDENTIFIER, ExprKind.MEMBER_ACCESS, ExprKind.INDEX) else None


_STANDARD_MEMBERS = IdentityLru()
_SYMBOL_NAMES = IdentityLru()
_SURFACE_NAMES = IdentityLru()


def project_setter_member(ctx: MemberCompletionContext, module: str, name: str) -> VbaProjectClassMember | None:
    surfaces = ctx.project_class_members
    if surfaces is None:
        return None
    members = _STANDARD_MEMBERS.get(surfaces)
    if members is None:
        members = {f"{surface.module_name}.{member.name}".lower(): member for surface in surfaces if surface.kind == "standardModule" for member in surface.members}
        _STANDARD_MEMBERS.put(members, surfaces)
    result: VbaProjectClassMember | None = members.get(f"{module}.{name}".lower())
    return result


def _symbol_names(symbols: Sequence[VbaSymbol]) -> tuple[set[str], set[str]]:
    cached = _SYMBOL_NAMES.get(symbols)
    if cached is None:
        all_names, indexed = set(), set()
        for symbol in symbols:
            if symbol.kind in (VbaSymbolKind.PROPERTY_LET, VbaSymbolKind.PROPERTY_SET):
                all_names.add(symbol.name.lower())
                if sum(child.kind is VbaSymbolKind.PARAMETER for child in symbol.children or []) > 1:
                    indexed.add(symbol.name.lower())
        cached = _SYMBOL_NAMES.put((all_names, indexed), symbols)
    return cached  # type: ignore[no-any-return]


def _may_be_setter(symbols: ModuleSymbols, visible: Sequence[VbaSymbol] | None, ctx: MemberCompletionContext, name: str, indexed: bool) -> bool:
    slot = 0 if indexed else 1
    if name in _symbol_names(symbols.root.children or ())[slot] or (visible is not None and name in _symbol_names(visible)[slot]):
        return True
    surfaces = ctx.project_class_members
    if surfaces is None:
        return False
    names = _SURFACE_NAMES.get(surfaces)
    if names is None:
        all_names, indexed_names = set(), set()
        for surface in surfaces:
            for member in surface.members:
                params = [value for kind, value in (member.procedure_params or {}).items() if kind in ("propertyLet", "propertySet")]
                if params:
                    all_names.add(member.name.lower())
                    if any(len(value) > 1 for value in params):
                        indexed_names.add(member.name.lower())
        names = _SURFACE_NAMES.put((all_names, indexed_names), surfaces)
    return name in names[slot]


@dataclass(slots=True)
class SourceSetterAssignment(CallArguments):
    index_params: list[CallableParamType] = field(default_factory=list)
    indexed: bool = False
    legacy_no_index: bool = False


def source_setter_assignment(source: str, span: Span, symbols: ModuleSymbols, procedure: VbaSymbol | None, visible: Sequence[VbaSymbol] | None, ctx: MemberCompletionContext) -> SourceSetterAssignment | None:
    tokens = tokens_without_leading_line_number(statement_tokens_after_leading_label(source, span))
    if not tokens or tokens[0].raw_text.lower() == "if":
        return None
    uses_set = tokens[0].raw_text.lower() == "set"
    if uses_set:
        tokens = tokens[1:]
    equals = top_level_operator_index(tokens, "=")
    if equals < 0:
        return None
    candidate = assignment_target_name(tokens[:equals])
    name = token_name(tokens[candidate[0]]) if candidate else None
    if candidate is None or not name or not _may_be_setter(symbols, visible, ctx, name.lower(), candidate[1]):
        return None
    target = assignment_target_from_tokens(tokens[:equals + 1])
    named = assignment_target_name(target) if target else None
    if target is None or named is None:
        return None
    index, indexed = named
    token = target[index]
    name = token_name(token)
    if not name:
        return None
    kind = VbaSymbolKind.PROPERTY_SET if uses_set else VbaSymbolKind.PROPERTY_LET
    params: list[VbaProcedureParam] | None
    legacy = False
    qualifier: str | None
    if index == 0:
        binding = resolve_bare_identifier_binding(BareIdentifierResolutionInput(current_module=symbols, enclosing_procedure=procedure, project_visible_symbols=visible or (), name=name, context=BareIdentifierContext.ASSIGNMENT_TARGET))
        if binding.scope is BareIdentifierResolutionScope.AMBIGUOUS:
            return None
        setter = next((definition for definition in binding.definitions if definition.kind is kind), None)
        if setter is None:
            return None
        params = procedure_params_from_symbol(setter, include_passing=True)
        if setter.module_name.lower() == symbols.module_name.lower():
            params = [replace(param, type_=param.type_ or (symbols.def_types or {}).get(param.name[:1].lower())) for param in params]
        else:
            member = project_setter_member(ctx, setter.module_name, name)
            params = (member.procedure_params or {}).get(kind.value, params) if member else params
        qualifier = setter.module_name
    else:
        member_completion = resolve_exact_member_completion(source, name, span.start + token.end, ctx)
        params = (member_completion.procedure_params or {}).get(kind.value) if member_completion else None
        qualifier = member_completion.owner if member_completion else None
        legacy = not uses_set and member_completion is not None and member_completion.signature is None and params is not None and len(params) == 1
    if not params:
        return None
    index_params = [CallableParamType(name=param.name, type_=param.type_, optional=bool(param.optional), param_array=bool(param.param_array), is_array=param.is_array, by_ref=is_by_ref_procedure_param(param.by_ref, param.by_val, param.param_array)) for param in params[:-1]]
    split = ArgSplit([], []) if not indexed or len(target) == index + 3 else split_arg_slots(target[index + 2:-1], span.start)
    return SourceSetterAssignment(name=name, qualifier=qualifier, indexed=indexed, name_span=Span(span.start + token.start, span.start + token.end), index_params=index_params, slots=split.slots, slot_spans=split.spans, slice_start=span.start, legacy_no_index=legacy)


def invalid_setter_assignment_arity(assignment: SourceSetterAssignment, source: str, push: PushFn) -> bool:
    if assignment.legacy_no_index and assignment.slots:
        return True
    invalid = False

    def report(rule: str, message: str, span: Span, data: VbaDiagnosticData | None = None) -> None:
        nonlocal invalid
        invalid = True
        placeholder = data.missing_required_argument_placeholder if data else None
        if data is not None and placeholder and not assignment.indexed and not assignment.slots:
            data = replace(data, missing_required_argument_placeholder=replace(placeholder, edit=replace(placeholder.edit, new_text=f"({placeholder.edit.new_text.strip()})")))
        if not assignment.slots and any(not param.optional and not param.param_array for param in assignment.index_params):
            message = f"Argument not optional: property '{assignment.name}' requires an index."
        push(rule, message, span, data)

    validate_arity(source, CallableTypeSignature(assignment.name, assignment.index_params), assignment, report)
    return invalid
