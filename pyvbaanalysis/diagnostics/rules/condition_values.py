"""Rule family: a Collection, an array or a Variant holding an array read as a
condition or a Boolean operand (XLIDE issue #424, each measured in Excel 16.0).

A Collection's default member Item needs an index. As the condition of an If, a
loop or IIf it raises 450 when it runs; as a Select Case subject or an operand of
Not, And or Or it does not compile ("Argument not optional"). An array declared
as one is a compile "Type mismatch" as a condition or a Select Case subject, and
IIf raises 13 on it; `Not a` runs (it is the `Not Not a` idiom), and And or Or on
it is non-scalar-binary-operand's. A Variant holding an array raises 13 in every
one of these places.

The forms are read on each statement, a block If's own line and each ElseIf
line, a loop's or a Select's opening line and a Do's Loop line.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/conditionValues.ts.
"""

from __future__ import annotations

from collections.abc import Callable, Iterator, Mapping

from ...conditional import ConditionalActivityTracker
from ...parser.nodes import (
    BodyNode,
    IfBlockNode,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    Span,
    StatementNode,
    is_leaf_statement,
)
from ...symbols.symbol_model import ModuleSymbols, VbaSymbolKind
from ...types.type_inference import procedure_symbol_for
from ...types.type_names import normalize_type
from ..block_headers import block_header_statements
from ..condition_operands import ConditionForm, condition_operands
from ..context import PushFn
from ..held_objects import HeldObjects, held_objects_at
from ..walker import active_module_members, is_inactive_node, statement_tokens_after_leading_label
from .arrays import FixedArrayBound, known_array_shapes_at, module_option_base

_WHERE: dict[ConditionForm, str] = {
    "condition": "the condition",
    "select": "Select Case",
    "iif": "IIf",
    "not": "'Not'",
    "logical": "the Boolean operator",
}


def check_condition_values(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    option_base = module_option_base(mod, activity)
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        proc_symbol = procedure_symbol_for(symbols, member)
        locals_ = [
            child
            for child in ((proc_symbol.children if proc_symbol is not None else None) or [])
            if child.kind is VbaSymbolKind.LOCAL_VARIABLE
        ]
        collections: dict[str, bool] = {
            child.name.lower(): child.is_auto_instantiated is True
            for child in locals_
            if not child.is_array and normalize_type(child.as_type) == "collection"
        }
        arrays = {child.name.lower() for child in locals_ if child.is_array}
        variants = {
            child.name.lower()
            for child in locals_
            if not child.is_array and (normalize_type(child.as_type) or "variant") == "variant"
        }
        if not collections and not arrays and not variants:
            continue
        _check_procedure(source, member, symbols, activity, push, option_base, collections, arrays, variants)


def _check_procedure(
    source: str,
    member: ProcedureNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    option_base: int,
    collections: Mapping[str, bool],
    arrays: set[str],
    variants: set[str],
) -> None:
    shapes_at: list[Callable[[LeafStatementNode], Mapping[str, FixedArrayBound]]] = []
    # A Collection set here holds one, as an `As New` one always does (issue
    # #415, measured in Excel 16.0).
    held_at: list[Callable[[BodyNode], HeldObjects]] = []

    def holds_one(stmt: LeafStatementNode, lower: str) -> bool:
        if collections.get(lower) is True:
            return True
        if not held_at:
            held_at.append(held_objects_at(source, member, symbols, activity))
        held = held_at[0](stmt).classes.get(lower)
        return held is not None and held.lower() == "collection"

    def check(stmt: LeafStatementNode) -> None:
        toks = statement_tokens_after_leading_label(source, stmt.span)
        for operand in condition_operands(toks):
            form = operand.form
            tok = toks[operand.index]
            lower = tok.raw_text.lower()
            at = Span(stmt.span.start + tok.start, stmt.span.start + tok.end)
            where = _WHERE[form]
            if lower in collections:
                if form in ("select", "not", "logical"):
                    push(
                        "collectionOperand",
                        f"'{tok.raw_text}' is a Collection: its default member Item needs an index, so {where} "
                        "has no value to work on. This is a VBE compile error: Argument not optional.",
                        at,
                    )
                elif holds_one(stmt, lower):
                    push(
                        "objectDefaultValue",
                        f"'{tok.raw_text}' is a Collection: its default member Item needs an index, so {where} "
                        "has no value to read. This will raise Run-time error '450': Wrong number of arguments "
                        "or invalid property assignment.",
                        at,
                    )
            elif lower in arrays:
                if form in ("condition", "select"):
                    push(
                        "nonScalarBinaryOperand",
                        f"'{tok.raw_text}' is declared as an array, which {where} cannot read as one value. "
                        "This will fail to compile with 'Type mismatch'.",
                        at,
                    )
                elif form == "iif":
                    push(
                        "variantValueMisuse",
                        f"'{tok.raw_text}' is an array, which IIf cannot read as its condition. This will raise "
                        "Run-time error '13': Type mismatch.",
                        at,
                    )
            elif lower in variants:
                if not shapes_at:
                    shapes_at.append(known_array_shapes_at(source, symbols, member, activity, option_base))
                shape = shapes_at[0](stmt).get(lower)
                if shape:
                    push(
                        "variantValueMisuse",
                        f"'{tok.raw_text}' holds an array from {shape.origin} here, which {where} cannot read "
                        "as one value. This will raise Run-time error '13': Type mismatch.",
                        at,
                    )

    # Upstream recurses per block: a block's header line, its body, then the line
    # that closes it. The explicit stack keeps that order (a pending closing line
    # waits under the body's iterator).
    stack: list[Iterator[BodyNode] | StatementNode] = [iter(member.body)]
    while stack:
        top = stack[-1]
        if isinstance(top, StatementNode):
            stack.pop()
            check(top)
            continue
        node = next(top, None)
        if node is None:
            stack.pop()
            continue
        if is_inactive_node(activity, node):
            continue
        if is_leaf_statement(node):
            check(node)
            continue
        body = getattr(node, "body", None)
        if not isinstance(body, list):
            continue
        before, after = block_header_statements(source, node)
        opening = node.branches[0].header_span if isinstance(node, IfBlockNode) and node.branches else None
        header = before
        if header is None and opening is not None:
            header = StatementNode(span=opening, raw=source[opening.start : opening.end])
        if header is not None:
            check(header)
        if after is not None:
            stack.append(after)
        stack.append(iter(body))
