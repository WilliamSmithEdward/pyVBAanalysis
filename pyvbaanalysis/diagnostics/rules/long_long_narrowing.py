"""Rule: a LongLong or LongPtr narrowed implicitly in 64-bit VBA.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/longLongNarrowing.ts
(XLIDE issue #298, each measured in 64-bit Excel 16.0: a compile error, "Type
mismatch").

In 64-bit Office a LongLong, and so a LongPtr, converts to no narrower
whole-number type on its own: stored in a Long, Integer, Byte or Currency, passed
ByVal to such a parameter, used as an array index or bound, as a For counter's
bound, or as a VBA function's whole-number argument such as Mid's Start. VarPtr,
ObjPtr and StrPtr return a LongPtr, so `n = VarPtr(x)`, how 32-bit code keeps a
pointer, stops compiling. CLng(q) converts; a Double, Variant or String takes the
value; `q / 1` is a Double. The rule follows the Win64 compiler constant: in
32-bit Office LongPtr is a Long and there is no LongLong.
"""

from __future__ import annotations

from collections.abc import Callable, Mapping, Sequence

from ...conditional import ConditionalActivityTracker
from ...conditional.conditional_compilation import (
    ConditionalCompilationEnvironment,
    compiler_constants_with_defaults,
)
from ...lexer.token_helpers import match_paren_from, split_top_level_token_groups
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    ForBlockNode,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    Span,
    iter_body_nodes,
)
from ...symbols.symbol_model import ModuleSymbols, VbaProcedureSignature, VbaSymbol
from ...types.type_inference import procedure_symbol_for, type_environment_for
from ...types.type_names import normalize_type
from ..call_extraction import CallArguments, CallableTypeSignature, extract_call, is_named_slot
from ..callable_signatures import (
    callable_type_signatures_for,
    expression_calls,
    runtime_callable_source_shadowed,
    source_name_scope_for,
)
from ..context import PushFn, statement_tokens
from ..walker import (
    active_module_members,
    bare_assignment_target,
    block_header_line_span,
    for_each_statement,
    statement_and_branch_spans,
    token_name,
    token_text,
)
from .runtime_values import whole_number_arguments

_NARROW = frozenset({"long", "integer", "byte", "currency"})
_WIDE = frozenset({"longlong", "longptr"})
_WHOLE = frozenset({"long", "integer", "byte", "longlong", "longptr"})
_POINTER_FUNCTIONS = frozenset({"varptr", "objptr", "strptr"})
_OPERATORS = frozenset({"+", "-", "*", "\\", "mod"})


def check_long_long_narrowing(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    conditional_compilation: ConditionalCompilationEnvironment | None,
    host: str | None,
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    win64 = compiler_constants_with_defaults(conditional_compilation).get("win64")
    if isinstance(win64, bool):
        on = win64
    elif isinstance(win64, (int, float)):
        on = win64 != 0
    else:
        on = False
    if (host is not None and host.lower() == "vb6") or not on:
        return
    signatures = callable_type_signatures_for(symbols, project_procedures)
    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            _check_procedure(
                source, member, symbols, signatures, host, project_visible_symbols, activity, push
            )


def _check_procedure(
    source: str,
    member: ProcedureNode,
    symbols: ModuleSymbols,
    signatures: Mapping[str, CallableTypeSignature],
    host: str | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    env = type_environment_for(symbols, member)
    source_names = source_name_scope_for(symbols, member, project_visible_symbols)
    proc_sym = procedure_symbol_for(symbols, member)
    arrays = {
        child.name.lower()
        for child in ((proc_sym.children if proc_sym is not None else None) or [])
        if child.is_array
    }

    def wide(toks: Sequence[VbaToken]) -> bool:
        try:
            return _wide_expression(
                [tok for tok in toks if tok.kind is not TokenKind.COMMENT],
                env,
                lambda name: runtime_callable_source_shadowed(name, source_names),
            )
        except RecursionError:
            # Port-only: parentheses nested past Python's frame limit, where
            # upstream recurses on a deeper stack. Unknown, so nothing is said.
            return False

    def report(toks: Sequence[VbaToken], base: int, where: str) -> None:
        value = [tok for tok in toks if tok.kind is not TokenKind.COMMENT]
        push(
            "longLongNarrowing",
            f"{' '.join(tok.raw_text for tok in value)} is a LongLong in 64-bit VBA, and {where} takes "
            "no LongLong without a conversion such as CLng. This is a VBE compile error: Type mismatch.",
            Span(base + value[0].start, base + value[-1].end),
        )

    def check_span(span: Span) -> None:
        toks = statement_tokens(source, span)
        target = bare_assignment_target(source, span)
        if target is not None and len(target[2]) > 0:
            target_name, _target_span, value_tokens = target
            lower = target_name.lower()
            type_ = normalize_type(member.return_type if lower == member.name.lower() else env.get(lower))
            if type_ and type_ in _NARROW and lower not in arrays and wide(value_tokens):
                report(value_tokens, span.start, f"'{target_name}', {_article(type_)} {_capitalize(type_)},")
        # `a(q)` and `ReDim a(q)`: an index or a bound.
        for i, tok in enumerate(toks):
            name = token_name(tok)
            lower_name = name.lower() if name else None
            if (
                not lower_name
                or lower_name not in arrays
                or _raw_at(toks, i + 1) != "("
                or _raw_at(toks, i - 1) == "."
            ):
                continue
            close = match_paren_from(toks, i + 1)
            for slot in split_top_level_token_groups(toks, i + 2, ",", close) if close > 0 else []:
                bounds = split_top_level_token_groups(slot, 0, "to")
                for bound in bounds if len(bounds) > 0 else [slot]:
                    if len(bound) > 0 and wide(bound):
                        report(
                            bound,
                            span.start,
                            "an array bound" if token_text(_at(toks, 0)) == "redim" else "an array index",
                        )
        # Calls: a ByVal whole-number parameter of a procedure, and a VBA
        # function's whole-number argument.
        calls: list[CallArguments] = []
        first_call = extract_call(source, span)
        if first_call is not None:
            calls.append(first_call)
        calls.extend(expression_calls(source, span, signatures, source_names))
        seen: set[int] = set()
        for call in calls:
            if call.name_span.start in seen or any(is_named_slot(slot) for slot in call.slots):
                continue
            seen.add(call.name_span.start)
            base = call.slice_start
            signature = None if call.qualifier else signatures.get(call.name.lower())
            if signature is None:
                continue
            params = signature.params
            for k, slot in enumerate(call.slots):
                param = params[k] if k < len(params) else None
                param_type = normalize_type(param.type_ if param is not None else None)
                if (
                    param is not None
                    and not param.by_ref
                    and param_type
                    and param_type in _NARROW
                    and len(slot) > 0
                    and wide(slot)
                ):
                    report(
                        slot,
                        base,
                        f"the ByVal {_capitalize(param_type)} '{param.name}' of '{signature.name}'",
                    )
        # `Mid$("abc", q, 1)`: a VBA function's whole-number argument, `$` or not.
        for i, tok in enumerate(toks):
            name = token_name(tok)
            open_at = i + 2 if _raw_at(toks, i + 1) == "$" else i + 1
            if (
                not name
                or _raw_at(toks, open_at) != "("
                or _raw_at(toks, i - 1) == "."
                or name.lower() in signatures
                or runtime_callable_source_shadowed(name, source_names)
            ):
                continue
            close = match_paren_from(toks, open_at)
            slots = split_top_level_token_groups(toks, open_at + 1, ",", close) if close > open_at else []
            for slot in whole_number_arguments(name, slots, host) or []:
                if wide(slot):
                    report(slot, span.start, f"{name}'s whole-number argument")

    def visit(stmt: LeafStatementNode) -> None:
        for span in statement_and_branch_spans(stmt):
            check_span(span)

    for_each_statement(member.body, visit, activity)

    # A For counter's bounds: `For i = 1 To q` with i a Long.
    def skip(node: BodyNode) -> bool:
        return (activity is not None and activity.is_inactive(node.span)) or not isinstance(
            getattr(node, "body", None), list
        )

    for node in iter_body_nodes(member.body, skip):
        if (
            not isinstance(node, ForBlockNode)
            or node.each
            or not node.control_variable
            or (normalize_type(env.get(node.control_variable.lower())) or "") not in _NARROW
        ):
            continue
        header = block_header_line_span(source, node.span)
        toks = statement_tokens(source, header)
        eq = _find(toks, lambda tok: tok.raw_text == "=")
        to = _find(toks, lambda tok: token_text(tok) == "to")
        step = _find(toks, lambda tok: token_text(tok) == "step")
        ranges = [(eq + 1, to), (to + 1, step if step > 0 else len(toks))]
        if step > 0:
            ranges.append((step + 1, len(toks)))
        for start, end in ranges:
            part = _js_slice(toks, start, end)
            if eq > 0 and to > eq and len(part) > 0 and wide(part):
                report(part, header.start, f"the Long counter '{node.control_variable}'")


def _wide_expression(
    toks: Sequence[VbaToken], env: Mapping[str, str], shadowed: Callable[[str], bool]
) -> bool:
    """Whether an expression is a whole number that is a LongLong: a LongLong or
    LongPtr name, a `^` literal or a pointer function, with whole-number operands
    and +, -, *, \\ or Mod between them."""
    saw_wide = False
    i = 0
    expect_operand = True
    while i < len(toks):
        tok = toks[i]
        word = token_text(tok)
        if expect_operand:
            if tok.raw_text == "-" or tok.raw_text == "+":
                i += 1
                continue
            if tok.raw_text == "(":
                close = match_paren_from(toks, i)
                if close < 0:
                    return False
                inner = toks[i + 1 : close]
                inner_wide = _wide_expression(inner, env, shadowed)
                if not inner_wide and not _whole_expression(inner, env):
                    return False
                saw_wide = saw_wide or inner_wide
                i = close + 1
            elif tok.kind is TokenKind.INTEGER_LITERAL:
                saw_wide = saw_wide or tok.raw_text.endswith("^")
                i += 1
            elif (
                token_name(tok)
                and word in _POINTER_FUNCTIONS
                and _raw_at(toks, i + 1) == "("
                and _raw_at(toks, i - 1) != "."
                and not shadowed(tok.raw_text)
            ):
                close = match_paren_from(toks, i + 1)
                if close < 0:
                    return False
                saw_wide = True
                i = close + 1
            elif (
                token_name(tok)
                and _raw_at(toks, i + 1) != "("
                and _raw_at(toks, i + 1) != "."
                and _raw_at(toks, i - 1) != "."
            ):
                type_ = normalize_type(env.get(word))
                if not type_ or type_ not in _WHOLE:
                    return False
                saw_wide = saw_wide or type_ in _WIDE
                i += 1
            else:
                return False
            expect_operand = False
            continue
        if (word or tok.raw_text) not in _OPERATORS:
            return False
        expect_operand = True
        i += 1
    return saw_wide and not expect_operand


def _whole_expression(toks: Sequence[VbaToken], env: Mapping[str, str]) -> bool:
    """Whether an expression is made of whole-number names and literals only."""
    return len(toks) > 0 and all(
        tok.kind is TokenKind.INTEGER_LITERAL
        or (token_text(tok) or tok.raw_text) in _OPERATORS
        or tok.raw_text in ("(", ")")
        or (normalize_type(env.get(token_text(tok))) or "") in _WHOLE
        for tok in toks
    )


def _article(type_: str) -> str:
    return "an" if type_[:1].lower() in ("a", "e", "i", "o", "u") else "a"


def _capitalize(type_: str) -> str:
    return type_[:1].upper() + type_[1:]


def _find(toks: Sequence[VbaToken], test: Callable[[VbaToken], bool]) -> int:
    return next((i for i, tok in enumerate(toks) if test(tok)), -1)


def _js_slice(toks: Sequence[VbaToken], start: int, end: int) -> list[VbaToken]:
    """Array.prototype.slice: a negative index counts from the end."""
    length = len(toks)
    lo = max(length + start, 0) if start < 0 else min(start, length)
    hi = max(length + end, 0) if end < 0 else min(end, length)
    return list(toks[lo:hi])


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    return toks[i] if 0 <= i < len(toks) else None


def _raw_at(toks: Sequence[VbaToken], i: int) -> str | None:
    tok = _at(toks, i)
    return tok.raw_text if tok is not None else None
