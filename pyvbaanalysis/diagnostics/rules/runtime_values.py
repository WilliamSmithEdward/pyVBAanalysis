"""Rule family: deterministic runtime argument / conversion values.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/runtimeValues.ts. Some
runtime-library arguments have deterministic value bounds even when the argument
type is valid (e.g. Left(s, -1) raises Run-time error 5), and selected
conversions fail for provably invalid literals (CDate("abc") raises error 13).
"""

from __future__ import annotations

import re
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass

from ...conditional import ConditionalActivityTracker
from ...constants.integer_constant_expression import (
    IntegerConstantLookup,
    evaluate_integer_constant_expression,
    parse_vba_integer_literal,
    resolve_raw_integer_constants,
)
from ...lexer.token_helpers import match_paren_from
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import LeafStatementNode, ModuleNode, ProcedureNode, Span
from ...symbols.symbol_model import ModuleSymbols, VbaProcedureSignature, VbaSymbol
from ...types.type_inference import procedure_symbol_for, type_environment_for
from ..call_extraction import (
    CallableTypeSignature,
    empty_arg_split,
    named_argument_slot,
    split_arg_slots,
    string_literal_value,
    unwrap_outer_parens,
)
from ..callable_signatures import (
    SourceNameScope,
    callable_type_signatures_for,
    runtime_callable_source_shadowed,
    scoped_integer_constant_lookup,
    source_name_scope_for,
)
from ..const_expr import collect_body_literal_integer_constants, collect_module_literal_integer_constants
from ..context import PushFn, statement_tokens
from ..walker import ProcedureStatementVisitor, token_name, token_text


@dataclass(frozen=True, slots=True)
class _RuntimeArgumentValueSpec:
    canonical_name: str
    parameter_name: str
    argument_index: int
    minimum: int | None = None
    maximum: int | None = None
    minimum_slot_count: int | None = None
    allow_named: bool = True


_DECLARATION_HEADS = frozenset(
    {
        "dim", "static", "const", "private", "public", "friend", "declare",
        "sub", "function", "property", "type", "enum",
    }
)
_SUFFIX_FUNCTIONS = frozenset({"Left", "Right", "String", "Space", "Mid"})


def check_runtime_argument_values(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None,
    project_integer_constants: Mapping[str, str | None] | None,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> ProcedureStatementVisitor:
    module_signatures = callable_type_signatures_for(symbols, project_procedures)
    project_constants = resolve_raw_integer_constants(project_integer_constants or {}, {})
    module_constants = collect_module_literal_integer_constants(mod, activity, project_constants)

    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        env = type_environment_for(symbols, member)
        source_names = source_name_scope_for(symbols, member, project_visible_symbols)
        procedure_constants = dict(module_constants)
        collect_body_literal_integer_constants(member.body, procedure_constants, activity)
        proc_sym = procedure_symbol_for(symbols, member)
        constants = scoped_integer_constant_lookup(
            procedure_constants, symbols, proc_sym, project_visible_symbols
        )

        def visitor(stmt: LeafStatementNode) -> None:
            for display_name, parameter_name, value, span in _runtime_argument_value_hits(
                source, stmt.span, module_signatures, env, constants, source_names
            ):
                push(
                    "runtimeArgumentValue",
                    f"Argument '{parameter_name}' of '{display_name}' is {value}; this will "
                    "raise Run-time error '5': Invalid procedure call or argument.",
                    span,
                )

        return visitor

    return factory


def _runtime_argument_value_hits(
    source: str,
    span: Span,
    module_signatures: Mapping[str, CallableTypeSignature],
    env: Mapping[str, str],
    constants: IntegerConstantLookup,
    source_names: SourceNameScope,
) -> list[tuple[str, str, int, Span]]:
    toks = statement_tokens(source, span)
    if _is_declaration_like_statement(toks):
        return []
    hits: list[tuple[str, str, int, Span]] = []
    for i in range(len(toks) - 1):
        call = _runtime_argument_value_call_at(toks, i, span, module_signatures, env, source_names)
        if call is None:
            continue
        display_name, specs, slots = call
        for spec in specs:
            slot = _runtime_argument_value_slot(slots, spec)
            literal = (
                _integer_argument_outside_bounds(source, slot, span.start, spec, constants)
                if slot is not None
                else None
            )
            if literal is None:
                continue
            value, hit_span = literal
            hits.append((display_name, spec.parameter_name, value, hit_span))
    return hits


def _runtime_argument_value_call_at(
    toks: Sequence[VbaToken],
    index: int,
    span: Span,
    module_signatures: Mapping[str, CallableTypeSignature],
    env: Mapping[str, str],
    source_names: SourceNameScope,
) -> tuple[str, Sequence[_RuntimeArgumentValueSpec], list[list[VbaToken]]] | None:
    name = token_name(toks[index])
    if not name:
        return None
    qualifier = (
        token_name(toks[index - 2])
        if index >= 2 and toks[index - 1].raw_text == "."
        else None
    )
    if index > 0 and toks[index - 1].raw_text == "." and not qualifier:
        return None
    if qualifier and qualifier.lower() != "vba":
        return None

    paren_index = index + 1
    suffix = ""
    if paren_index < len(toks) and toks[paren_index].raw_text == "$":
        suffix = toks[paren_index].raw_text
        paren_index += 1
    if paren_index >= len(toks) or toks[paren_index].raw_text != "(":
        return None

    specs = _runtime_argument_value_specs(name)
    if not specs:
        return None
    if suffix and specs[0].canonical_name not in _SUFFIX_FUNCTIONS:
        return None
    lower = specs[0].canonical_name.lower()
    if not qualifier and (
        lower in module_signatures
        or lower in env
        or runtime_callable_source_shadowed(name, source_names)
    ):
        return None

    close = match_paren_from(toks, paren_index)
    if close < 0:
        return None
    inner = list(toks[paren_index + 1 : close])
    split = empty_arg_split() if not inner else split_arg_slots(inner, span.start)
    return (f"{specs[0].canonical_name}{suffix}", specs, split.slots)


def _runtime_argument_value_specs(name: str) -> list[_RuntimeArgumentValueSpec]:
    lower = name.lower()
    if lower == "left":
        return [_RuntimeArgumentValueSpec("Left", "Length", 1, minimum=0)]
    if lower == "right":
        return [_RuntimeArgumentValueSpec("Right", "Length", 1, minimum=0)]
    if lower == "string":
        return [_RuntimeArgumentValueSpec("String", "Number", 0, minimum=0)]
    if lower == "space":
        return [_RuntimeArgumentValueSpec("Space", "Number", 0, minimum=0)]
    if lower == "mid":
        return [
            _RuntimeArgumentValueSpec("Mid", "Start", 1, minimum=1),
            _RuntimeArgumentValueSpec("Mid", "Length", 2, minimum=0),
        ]
    if lower == "replace":
        return [
            _RuntimeArgumentValueSpec("Replace", "Start", 3, minimum=1),
            _RuntimeArgumentValueSpec("Replace", "Count", 4, minimum=-1),
        ]
    if lower == "instr":
        return [
            _RuntimeArgumentValueSpec(
                "InStr", "Start", 0, minimum=1, minimum_slot_count=3, allow_named=False
            )
        ]
    if lower == "chr":
        return [_RuntimeArgumentValueSpec("Chr", "CharCode", 0, minimum=0, maximum=255)]
    if lower == "chrw":
        return [_RuntimeArgumentValueSpec("ChrW", "CharCode", 0, maximum=65535)]
    return []


def _runtime_argument_value_slot(
    slots: Sequence[list[VbaToken]], spec: _RuntimeArgumentValueSpec
) -> list[VbaToken] | None:
    if spec.minimum_slot_count is not None and len(slots) < spec.minimum_slot_count:
        return None
    positional_index = 0
    for slot in slots:
        named = named_argument_slot(slot)
        if named is not None:
            if not spec.allow_named:
                continue
            if named[0].lower() == spec.parameter_name.lower():
                return named[1]
            continue
        if positional_index == spec.argument_index:
            return slot
        positional_index += 1
    return None


def _integer_argument_outside_bounds(
    source: str,
    slot: Sequence[VbaToken],
    slice_start: int,
    spec: _RuntimeArgumentValueSpec,
    constants: IntegerConstantLookup,
) -> tuple[int, Span] | None:
    toks = unwrap_outer_parens(
        [t for t in slot if t.kind is not TokenKind.COMMENT and t.kind is not TokenKind.NEWLINE]
    )
    if not toks:
        return None
    sign = 1
    literal: VbaToken | None = toks[0]
    start = toks[0].start
    literal_value: int | None = None
    signed_literal = len(toks) == 2 and toks[0].raw_text in ("-", "+")
    if signed_literal:
        sign = -1 if toks[0].raw_text == "-" else 1
        literal = toks[1]
        start = toks[0].start
    if (
        literal is not None
        and literal.kind is TokenKind.INTEGER_LITERAL
        and (len(toks) == 1 or signed_literal)
    ):
        raw_value = parse_vba_integer_literal(literal.raw_text)
        if raw_value is not None:
            literal_value = sign * raw_value
    if literal_value is not None and literal is not None:
        if _integer_argument_value_in_bounds(literal_value, spec):
            return None
        return (literal_value, Span(slice_start + start, slice_start + literal.end))

    expression_value = evaluate_integer_constant_expression(
        source[slice_start + toks[0].start : slice_start + toks[-1].end], constants
    )
    if expression_value is None or _integer_argument_value_in_bounds(expression_value, spec):
        return None
    return (expression_value, Span(slice_start + toks[0].start, slice_start + toks[-1].end))


def _integer_argument_value_in_bounds(value: int, spec: _RuntimeArgumentValueSpec) -> bool:
    if spec.minimum is not None and value < spec.minimum:
        return False
    return not (spec.maximum is not None and value > spec.maximum)


def _is_declaration_like_statement(toks: Sequence[VbaToken]) -> bool:
    return bool(toks) and token_text(toks[0]) in _DECLARATION_HEADS


# -- checkRuntimeConversionValues ------------------------------------------

_MONTH_NAME = re.compile(
    r"\b(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?"
    r"|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)\b",
    re.IGNORECASE,
)
_HAS_DIGIT = re.compile(r"[0-9]")
_NON_ASCII = re.compile(r"[^\x00-\x7F]")
_ALPHA_SPACE = re.compile(r"^[A-Za-z\s]+$")


def check_runtime_conversion_values(
    source: str,
    symbols: ModuleSymbols,
    project_visible_symbols: Sequence[VbaSymbol] | None,
    push: PushFn,
) -> ProcedureStatementVisitor:
    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        source_names = source_name_scope_for(symbols, member, project_visible_symbols)

        def visitor(stmt: LeafStatementNode) -> None:
            for display_name, name, span in _runtime_conversion_value_hits(
                source, stmt.span, source_names
            ):
                push(
                    "runtimeConversionValue",
                    f"{display_name} cannot convert {name} to Date. This will raise "
                    "Run-time error '13': Type mismatch.",
                    span,
                )

        return visitor

    return factory


def _runtime_conversion_value_hits(
    source: str, span: Span, source_names: SourceNameScope
) -> list[tuple[str, str, Span]]:
    toks = statement_tokens(source, span)
    if _is_declaration_like_statement(toks):
        return []
    hits: list[tuple[str, str, Span]] = []
    for i in range(len(toks) - 2):
        name = token_name(toks[i])
        if not name or name.lower() != "cdate":
            continue
        if toks[i + 1].raw_text != "(" or not _is_bare_or_vba_qualified_intrinsic_call(toks, i):
            continue
        qualified = i >= 1 and toks[i - 1].raw_text == "."
        if not qualified and runtime_callable_source_shadowed(name, source_names):
            continue
        close = match_paren_from(toks, i + 1)
        if close < 0:
            continue
        split = split_arg_slots(list(toks[i + 2 : close]), span.start)
        first_slot = split.slots[0] if split.slots else []
        if len(first_slot) != 1 or first_slot[0].kind is not TokenKind.STRING_LITERAL:
            continue
        value = string_literal_value(first_slot[0].raw_text)
        if not _is_definitely_invalid_date_string(value):
            continue
        hit_span = split.spans[0] if split.spans else Span(
            span.start + first_slot[0].start, span.start + first_slot[0].end
        )
        hits.append((f"VBA.{name}" if qualified else name, first_slot[0].raw_text, hit_span))
    return hits


def _is_definitely_invalid_date_string(value: str) -> bool:
    trimmed = value.strip()
    if not trimmed:
        return True
    if _HAS_DIGIT.search(trimmed) or _NON_ASCII.search(trimmed):
        return False
    if _ALPHA_SPACE.match(trimmed) is None:
        return False
    return _MONTH_NAME.search(trimmed) is None


def _is_bare_or_vba_qualified_intrinsic_call(toks: Sequence[VbaToken], name_index: int) -> bool:
    if name_index < 1 or toks[name_index - 1].raw_text != ".":
        return True
    qualifier = token_name(toks[name_index - 2]) if name_index >= 2 else None
    return (
        qualifier is not None
        and qualifier.lower() == "vba"
        and (name_index < 3 or toks[name_index - 3].raw_text != ".")
    )
