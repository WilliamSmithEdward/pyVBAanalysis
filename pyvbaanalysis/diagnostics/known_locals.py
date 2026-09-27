"""Locals whose value the procedure's text fixes.

Ported from knownLocalLiteralValues and plainLiteralText in
xlide_vscode/src/analyzer/diagnostics/typeInference.ts (XLIDE issues #118 and
#119). A local nothing ever assigns holds its default, 0 for a number and "" for a
String, and one whose every assignment is the same literal holds that literal. The
overflow, runtime-value, variant-value and expression rules read the map to prove
a value the declared type alone leaves open.
"""

from __future__ import annotations

import math
import re
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass, field

from ..conditional import ConditionalActivityTracker, inactive_node_skip
from ..constants.integer_constant_expression import IntegerConstantLookup, parse_vba_integer_literal
from ..lexer.token_kinds import TokenKind, VbaToken
from ..parser.nodes import ForBlockNode, ProcedureNode, Span, is_leaf_statement, iter_body_nodes
from ..symbols.symbol_model import ModuleSymbols, SymbolVisibility, VbaSymbolKind
from ..types.type_inference import procedure_symbol_for
from ..types.type_names import is_numeric_type, normalize_type
from .call_extraction import string_literal_value, unwrap_outer_parens
from .context import statement_tokens
from .walker import (
    bare_assignment_target,
    first_executable_token_index,
    statement_and_branch_spans,
    token_name,
    token_text,
)


@dataclass(frozen=True, slots=True)
class KnownLocalValue:
    """A local whose value the procedure's text fixes: its default, or one literal."""

    kind: str  # "number" | "string"
    value: int | float | str
    # "default" when nothing ever assigns it, "literal" when every assignment is
    # the same literal.
    origin: str
    # A `Mid(x, ...) = ` statement rewrites characters of the value without
    # changing its length, so the length is still known and the characters are not.
    content_mutated: bool = False


class ConstantOrKnownLocalLookup:
    """An IntegerConstantLookup over a procedure's integer constants, then the
    locals whose value the text fixes to a whole number: the `lookup` upstream's
    division and runtime-value rules build."""

    __slots__ = ("_constants", "_known")

    def __init__(self, constants: IntegerConstantLookup, known: Mapping[str, KnownLocalValue]) -> None:
        self._constants = constants
        self._known = known

    def get(self, name: str, /) -> int | None:
        constant = self._constants.get(name)
        if constant is not None:
            return constant
        local = self._known.get(name.lower())
        if local is None or local.kind != "number":
            return None
        value = local.value
        if isinstance(value, int):
            return value
        return int(value) if isinstance(value, float) and value.is_integer() else None


@dataclass(slots=True)
class _Candidate:
    # "number" | "string"; None for a Variant no literal has given a kind yet.
    kind: str | None
    # Numbers are held as numbers, which compare the way upstream's String() forms
    # of them do: 1 and 1.0 are one literal.
    literals: set[int | float | str] = field(default_factory=set)
    mutated: bool = False
    content_mutated: bool = False


_MUTATING_HEADS = frozenset({"set", "redim", "input", "get", "line", "erase", "lset", "rset"})
_FLOAT_SUFFIX_RE = re.compile(r"[!#@]$")
# Integers past this lose precision as a JavaScript number, so they stay floats.
_EXACT_INTEGER_LIMIT = 2**53


def known_local_literal_values(
    source: str,
    proc: ProcedureNode,
    symbols: ModuleSymbols,
    activity: ConditionalActivityTracker | None,
) -> dict[str, KnownLocalValue]:
    """The locals of a procedure whose value is plain from the text, keyed by
    lowercased name.

    Anything that could change a local another way - passing it whole to a call
    (ByRef), a For counter, `Input #`/`Get #`/`Line Input #`, `Mid(x, ...) =`,
    ReDim, `Set` - drops it from the map, as does any assignment whose value is
    not a plain literal. Variant and object locals are left out: Empty and Nothing
    are not the values these rules ask about.
    """
    proc_sym = procedure_symbol_for(symbols, proc)
    # A Variant (or untyped) local has no kind until a literal gives it one;
    # literals of two kinds, or none, leave it unknown (XLIDE issue #121: `v = 5`
    # then `v.Foo`).
    candidates: dict[str, _Candidate] = {}
    for child in (proc_sym.children if proc_sym is not None else None) or []:
        if (
            child.kind is not VbaSymbolKind.LOCAL_VARIABLE
            or child.is_array
            or child.visibility is SymbolVisibility.STATIC
        ):
            continue
        type_ = normalize_type(child.as_type)
        kind: str | None
        if type_ is None or type_ == "variant":
            kind = None
        elif is_numeric_type(type_):
            kind = "number"
        elif type_ == "string":
            kind = "string"
        else:
            continue
        if child.fixed_length is not None:
            continue
        candidates[child.name.lower()] = _Candidate(kind)
    if not candidates:
        return {}

    def mutate(lower: str | None) -> None:
        entry = candidates.get(lower) if lower else None
        if entry is not None:
            entry.mutated = True

    for node in iter_body_nodes(proc.body, inactive_node_skip(activity)):
        if isinstance(node, ForBlockNode):
            mutate(node.control_variable.lower() if node.control_variable else None)
        if isinstance(getattr(node, "body", None), list) or not is_leaf_statement(node):
            continue
        for span in statement_and_branch_spans(node):
            _scan_statement(source, span, candidates, mutate)

    out: dict[str, KnownLocalValue] = {}
    for lower, entry in candidates.items():
        if entry.mutated or entry.kind is None:
            continue  # a Variant nothing assigned is Empty, not a known literal
        if len(entry.literals) == 0:
            out[lower] = KnownLocalValue(
                entry.kind, 0 if entry.kind == "number" else "", "default", entry.content_mutated
            )
        elif len(entry.literals) == 1:
            (value,) = entry.literals
            out[lower] = KnownLocalValue(entry.kind, value, "literal", entry.content_mutated)
    return out


def _scan_statement(
    source: str,
    span: Span,
    candidates: dict[str, _Candidate],
    mutate: Callable[[str | None], None],
) -> None:
    toks = statement_tokens(source, span)
    first = first_executable_token_index(toks)
    head = token_text(_at(toks, first))
    bare = bare_assignment_target(source, span)
    if bare is not None:
        entry = candidates.get(bare[0].lower())
        if entry is not None:
            # Upstream reads the value from two tokens past the statement's first,
            # so `Let x = 5` never counts as a plain literal.
            value = [tok for tok in toks[first + 2 :] if tok.kind is not TokenKind.COMMENT]
            kind = entry.kind
            if kind is None:
                unwrapped = unwrap_outer_parens(value)
                kind = "string" if unwrapped and unwrapped[0].kind is TokenKind.STRING_LITERAL else "number"
            literal = _plain_literal(value, kind)
            if literal is None or (entry.kind is not None and entry.kind != kind):
                entry.mutated = True
            else:
                entry.kind = kind
                entry.literals.add(literal)
        return
    if head in _MUTATING_HEADS:
        for tok in toks:
            mutate(_lower_name(tok))
        return
    if head in ("mid", "mid$"):
        opener = _at(toks, first + 1)
        if opener is not None and opener.raw_text == "(":
            # `Mid(x, start, len) = value` rewrites characters of x and keeps its
            # length; anything else named in it is read.
            target = candidates.get(_lower_name(_at(toks, first + 2)) or "")
            if target is not None:
                target.content_mutated = True
            return
    # A whole name passed to any call may be ByRef: `Take d`, `Take(d)`,
    # `Call Take(d)`, `x = Take(d)`. Only a name standing alone in an argument
    # slot counts; `Take(d + 1)` copies.
    for i, tok in enumerate(toks):
        name = _lower_name(tok)
        if not name or name not in candidates:
            continue
        prev = _at(toks, i - 1)
        nxt = _at(toks, i + 1)
        opens_slot = (
            prev is None
            or prev.raw_text in ("(", ",")
            or prev.kind is TokenKind.IDENTIFIER
            or prev.kind is TokenKind.KEYWORD
        )
        closes_slot = (
            nxt is None
            or nxt.raw_text in (")", ",", ":")
            or nxt.kind is TokenKind.COMMENT
        )
        if (
            opens_slot
            and closes_slot
            and not (prev is not None and prev.kind is TokenKind.OPERATOR)
            and not (nxt is not None and nxt.kind is TokenKind.OPERATOR)
        ):
            mutate(name)


def _plain_literal(value: Sequence[VbaToken], kind: str) -> int | float | str | None:
    """Port of plainLiteralText: the literal a plain `x = literal` assigns, or None
    for any other value."""
    toks = unwrap_outer_parens(value)
    if kind == "string":
        if len(toks) == 1 and toks[0].kind is TokenKind.STRING_LITERAL:
            return string_literal_value(toks[0].raw_text)
        return None
    sign = 1
    rest = toks
    if rest and rest[0].raw_text in ("-", "+"):
        sign = -1 if rest[0].raw_text == "-" else 1
        rest = rest[1:]
    if len(rest) != 1:
        return None
    if rest[0].kind is TokenKind.INTEGER_LITERAL:
        parsed = parse_vba_integer_literal(rest[0].raw_text)
        return None if parsed is None else sign * parsed
    if rest[0].kind is TokenKind.FLOAT_LITERAL:
        # Number() in upstream: a D exponent (`1D3`) is not a number there either.
        try:
            parsed_float = float(_FLOAT_SUFFIX_RE.sub("", rest[0].raw_text))
        except ValueError:
            return None
        if not math.isfinite(parsed_float):
            return None
        return _as_js_number(sign * parsed_float)
    return None


def _as_js_number(value: float) -> int | float:
    """An integral float as an int, the way a JavaScript number prints: no `.0`,
    and equal as a literal to the same value written without a fraction."""
    if value.is_integer() and abs(value) < _EXACT_INTEGER_LIMIT:
        return int(value)
    return value


def _lower_name(tok: VbaToken | None) -> str | None:
    name = token_name(tok)
    return name.lower() if name else None


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    return toks[i] if 0 <= i < len(toks) else None
