"""Rule: an array resized or erased while VBA holds it locked (XLIDE issue #283,
each measured in Excel 16.0: error 10, "This array is fixed or temporarily
locked").

Ported from xlide_vscode/src/analyzer/diagnostics/rules/lockedArrays.ts. A For
Each over an array, a With on one of its elements, and an element passed ByRef
each lock the array until they end. Inside the loop or the With, `Erase a`,
`ReDim a(5)`, `ReDim Preserve a(5)` and, for a Variant holding the array,
`v = Array(9)` raise 10. So does a call that passes `a(0)` ByRef beside `a`
itself to a procedure of the module that erases or ReDims that array parameter.
Only a statement the block runs every time is judged: one inside an If, a
Select or another loop, or after an Exit, may not run.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Iterator, Mapping, Sequence
from collections.abc import Set as AbstractSet

from ...conditional import ConditionalActivityTracker
from ...js_compat import js_trim
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    ConditionalDirectiveNode,
    ForBlockNode,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    Span,
    StatementNode,
    VariableGroupNode,
    WithBlockNode,
    is_leaf_statement,
)
from ...symbols.symbol_model import ModuleSymbols, SymbolVisibility, VbaSymbolKind
from ...types.type_inference import procedure_symbol_for
from ...types.type_names import normalize_type
from ..call_extraction import extract_call, is_named_slot
from ..context import PushFn
from ..walker import (
    active_module_members,
    block_header_line_span,
    for_each_statement,
    is_inactive_node,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)
from .arrays import FixedArrayBound, known_array_shapes_at, module_option_base

# Statement heads after which a body's later statements may not run.
_LEAVING_HEADS = frozenset({"exit", "goto", "gosub", "return", "end", "resume", "on", "stop", "error"})

# `/^[a-z_]\w*$/` on a lowercased name, `\w` as JavaScript reads it.
_SIMPLE_NAME_RE = re.compile(r"[a-z_][A-Za-z0-9_]*")


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    return toks[i] if 0 <= i < len(toks) else None


def _lower_name(tok: VbaToken | None) -> str | None:
    name = token_name(tok)
    return name.lower() if name else None


def _child_body(node: BodyNode) -> list[BodyNode] | None:
    body = getattr(node, "body", None)
    return body if isinstance(body, list) else None


def check_locked_arrays(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    procedures: dict[str, ProcedureNode | None] = {}
    for member in active_module_members(mod, activity):
        if isinstance(member, ProcedureNode):
            lower = member.name.lower()
            procedures[lower] = None if lower in procedures else member
    option_base = module_option_base(mod, activity)
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        _check_procedure(source, member, symbols, activity, push, procedures, option_base)


def _check_procedure(
    source: str,
    member: ProcedureNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    procedures: Mapping[str, ProcedureNode | None],
    option_base: int,
) -> None:
    proc_sym = procedure_symbol_for(symbols, member)
    locals_ = [
        child
        for child in ((proc_sym.children if proc_sym is not None else None) or [])
        if child.kind is VbaSymbolKind.LOCAL_VARIABLE and child.visibility is not SymbolVisibility.STATIC
    ]
    dynamic = {child.name.lower() for child in locals_ if child.is_array and child.array_bounds is None}
    variants = {
        child.name.lower()
        for child in locals_
        if not child.is_array and (normalize_type(child.as_type) or "variant") == "variant"
    }
    if not dynamic and not variants:
        return
    shapes_at: list[Callable[[LeafStatementNode], Mapping[str, FixedArrayBound]]] = []

    # A Variant locks only while it holds an array with an element to step through.
    def holds_elements(node: BodyNode, lower: str) -> bool:
        if lower in dynamic:
            return True
        if not shapes_at:
            shapes_at.append(known_array_shapes_at(source, symbols, member, activity, option_base))
        shape = shapes_at[0](node).get(lower)  # type: ignore[arg-type]
        return shape is not None and len(shape.dims) > 0 and all(dim.upper >= dim.lower for dim in shape.dims)

    # Upstream recurses into each block's body; the bodies are frames on a stack
    # here, visited in the same order.
    stack: list[Iterator[BodyNode]] = [iter(member.body)]
    while stack:
        for node in stack[-1]:
            child = _child_body(node)
            if is_inactive_node(activity, node) or child is None:
                if is_leaf_statement(node) and not is_inactive_node(activity, node):
                    _check_element_pass(node.span, dynamic, procedures, source, activity, push)
                continue
            if isinstance(node, ForBlockNode) and node.each:
                lower = js_trim(node.source_expression).lower() if node.source_expression else None
                if (
                    lower
                    and _SIMPLE_NAME_RE.fullmatch(lower)
                    and (lower in dynamic or lower in variants)
                    and holds_elements(node, lower)
                ):
                    _report_unlocks(source, child, lower, lower in variants, "the For Each over it", activity, push)
            if isinstance(node, WithBlockNode):
                header = statement_tokens_after_leading_label(source, block_header_line_span(source, node.span))
                lower = _lower_name(_at(header, 1))
                if (
                    token_text(_at(header, 0)) == "with"
                    and lower
                    and lower in dynamic
                    and len(header) > 2
                    and header[2].raw_text == "("
                    and header[-1].raw_text == ")"
                ):
                    _report_unlocks(source, child, lower, False, "the With on its element", activity, push)
            stack.append(iter(child))
            break
        else:
            stack.pop()


def _every_time(
    source: str, body: Sequence[BodyNode], activity: ConditionalActivityTracker | None
) -> list[tuple[BodyNode, list[VbaToken]]]:
    """The statements a block runs every time, up to the first that may leave it."""
    out: list[tuple[BodyNode, list[VbaToken]]] = []
    for node in body:
        if is_inactive_node(activity, node) or isinstance(node, (ConditionalDirectiveNode, VariableGroupNode)):
            continue
        if not is_leaf_statement(node):
            leaves = [False]

            def scan(stmt: LeafStatementNode, leaves: list[bool] = leaves) -> None:
                if token_text(_at(statement_tokens_after_leading_label(source, stmt.span), 0)) in _LEAVING_HEADS:
                    leaves[0] = True

            for_each_statement([node], scan, activity)
            if leaves[0]:
                return out
            continue
        toks = statement_tokens_after_leading_label(source, node.span)
        if token_text(_at(toks, 0)) in _LEAVING_HEADS:
            return out
        if not (isinstance(node, StatementNode) and node.single_line_if_branches):
            out.append((node, toks))
    return out


def _report_unlocks(
    source: str,
    body: Sequence[BodyNode],
    lower: str,
    variant: bool,
    lock: str,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """What in a locked array's block resizes or erases it, each reported."""
    for node, toks in _every_time(source, body, activity):
        head = token_text(_at(toks, 0))
        at: VbaToken | None = None
        what = ""
        if head == "erase":
            at = next(
                (
                    tok
                    for i, tok in enumerate(toks)
                    if i > 0 and _lower_name(tok) == lower and toks[i - 1].raw_text != "."
                ),
                None,
            )
            what = "Erase cannot free it"
        elif head == "redim":
            target = _at(toks, 2 if token_text(_at(toks, 1)) == "preserve" else 1)
            at = target if _lower_name(target) == lower else None
            what = "ReDim cannot resize it"
        elif variant and _lower_name(_at(toks, 0)) == lower and len(toks) > 1 and toks[1].raw_text == "=":
            at = toks[0]
            what = "an assignment cannot replace it"
        if at is not None:
            push(
                "arrayTemporarilyLocked",
                f"'{at.raw_text}' is locked by {lock}, so {what}. This will raise Run-time error '10': "
                "This array is fixed or temporarily locked.",
                Span(node.span.start + at.start, node.span.start + at.end),
            )


def _check_element_pass(
    span: Span,
    dynamic: AbstractSet[str],
    procedures: Mapping[str, ProcedureNode | None],
    source: str,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """`Zap a(0), a`: an element passed ByRef beside the array to a procedure that erases or ReDims it."""
    call = extract_call(source, span)
    callee = procedures.get(call.name.lower()) if call is not None and not call.qualifier else None
    if call is None or callee is None or any(is_named_slot(slot) for slot in call.slots):
        return
    slots = [[tok for tok in slot if tok.kind is not TokenKind.COMMENT] for slot in call.slots]
    for k, slot in enumerate(slots):
        param = callee.params[k] if k < len(callee.params) else None
        lower = _lower_name(_at(slot, 0))
        if param is None or not param.is_array or len(slot) != 1 or not lower or lower not in dynamic:
            continue
        # Another argument is an element of the same array, passed ByRef.
        element = next(
            (
                j
                for j, other in enumerate(slots)
                if j != k
                and _lower_name(_at(other, 0)) == lower
                and len(other) > 1
                and other[1].raw_text == "("
                and other[-1].raw_text == ")"
                and j < len(callee.params)
                and not callee.params[j].by_val
                and not callee.params[j].is_array
            ),
            -1,
        )
        if element < 0:
            continue
        name = param.name.lower()

        def unlocks_with(toks: Sequence[VbaToken]) -> bool:
            head = token_text(_at(toks, 0))
            target = (
                _at(toks, 1)
                if head == "erase"
                else _at(toks, 2 if token_text(_at(toks, 1)) == "preserve" else 1)
                if head == "redim"
                else None
            )
            return _lower_name(target) == name

        if any(unlocks_with(toks) for _node, toks in _every_time(source, callee.body, activity)):
            at = slots[element][0]
            slot_spans = call.slot_spans
            offset = slot_spans[element].start if slot_spans is not None and element < len(slot_spans) else None
            push(
                "arrayTemporarilyLocked",
                f"'{call.name}' takes '{''.join(tok.raw_text for tok in slots[element])}' ByRef, which locks "
                f"'{slot[0].raw_text}', and then resizes or erases that array through '{param.name}'. This will "
                "raise Run-time error '10': This array is fixed or temporarily locked.",
                Span(offset, offset + (slots[element][-1].end - at.start)) if offset is not None else call.name_span,
            )
