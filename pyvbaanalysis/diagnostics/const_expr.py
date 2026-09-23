"""Integer constant collection for the diagnostics engine.

Ported from xlide_vscode/src/analyzer/diagnostics/constExpr.ts. The evaluator
itself lives in constants/integer_constant_expression.py (shared with the
structural analyzer); this module owns the diagnostics-side collection of raw
module-level and procedure-body integer constants and the fixed-length-string
size resolution that builds on them.

This collects the literal-integer constants the active rules need, resolves
external (VBA runtime / host) constants by name, and folds span-based integer
expressions via the shared evaluator.
"""

from __future__ import annotations

import re
from collections.abc import Mapping, Sequence

from ..conditional import ConditionalActivityTracker
from ..constants.integer_constant_expression import (
    IntegerConstantLookup,
    enum_member_raw_expression,
    evaluate_integer_constant_expression,
    parse_vba_integer_literal,
    resolve_raw_integer_constants,
    safe_integer,
)
from ..host.host_model import HostObjectModel, resolve_host_constant
from ..lexer.token_kinds import VbaToken
from ..parser.nodes import BodyNode, EnumNode, ModuleNode, Span, VariableGroupNode
from ..runtime.vba_runtime import resolve_runtime_constant
from .walker import active_module_members, for_each_variable_group


def collect_module_literal_integer_constants(
    mod: ModuleNode,
    activity: ConditionalActivityTracker | None,
    base: Mapping[str, int | None] | None = None,
) -> dict[str, int | None]:
    """Resolve module-level Const and Enum integer constants to their values.

    Only active Const groups and Enum members participate. Each value is an
    integer-literal expression or a reference to another such constant; anything
    that is not a deterministic integer resolves to None. A duplicate name is
    poisoned to None so an ambiguous constant never resolves.
    """
    raw_constants: dict[str, str | None] = {}
    seen: set[str] = set()
    for member in active_module_members(mod, activity):
        if isinstance(member, VariableGroupNode) and member.is_const:
            _add_raw_integer_constants(member, raw_constants, seen)
        elif isinstance(member, EnumNode):
            _add_raw_enum_integer_constants(member, raw_constants, seen)
    base_map: Mapping[str, int | None] = {} if base is None else base
    resolved: dict[str, int | None] = dict(base_map)
    for name, value in resolve_raw_integer_constants(raw_constants, base_map).items():
        resolved[name] = value
    return resolved


def collect_body_literal_integer_constants(
    body: Sequence[BodyNode],
    constants: dict[str, int | None],
    activity: ConditionalActivityTracker | None,
) -> None:
    """Fold a procedure body's local Const integer constants into ``constants``."""
    raw_constants: dict[str, str | None] = {}
    seen: set[str] = set()

    def collect(group: VariableGroupNode) -> None:
        if group.is_const:
            _add_raw_integer_constants(group, raw_constants, seen)

    for_each_variable_group(body, collect, activity)
    for name, value in resolve_raw_integer_constants(raw_constants, constants).items():
        constants[name] = value


def resolve_fixed_length_string_size(raw: str, constants: IntegerConstantLookup) -> int | None:
    """Resolve a fixed-length String size expression to an integer, or None."""
    return evaluate_integer_constant_expression(raw, constants)


def fold_integer_expression_tokens(
    source: str,
    span: Span,
    toks: Sequence[VbaToken],
    start: int,
    end_exclusive: int,
    constants: IntegerConstantLookup,
) -> int | None:
    """Evaluate the integer value of a token sub-range, or None if not constant."""
    if start >= end_exclusive:
        return None
    raw = source[span.start + toks[start].start : span.start + toks[end_exclusive - 1].end]
    return evaluate_integer_constant_expression(raw, constants)


def _add_raw_integer_constants(
    group: VariableGroupNode,
    raw_constants: dict[str, str | None],
    seen: set[str],
) -> None:
    for decl in group.declarations:
        name = _normalize_declared_constant_name(decl.name)
        if name is None:
            continue
        key = name.lower()
        if key in seen:
            raw_constants[key] = None
            continue
        seen.add(key)
        raw_constants[key] = decl.default_raw


def _add_raw_enum_integer_constants(
    en: EnumNode,
    raw_constants: dict[str, str | None],
    seen: set[str],
) -> None:
    previous_name: str | None = None
    for member in en.members:
        name = _normalize_declared_constant_name(member.name)
        if name is None:
            continue
        key = name.lower()
        if key in seen:
            raw_constants[key] = None
            previous_name = name
            continue
        seen.add(key)
        raw_constants[key] = enum_member_raw_expression(member.value_raw, previous_name)
        previous_name = name


def _normalize_declared_constant_name(raw: str) -> str | None:
    text = raw.strip()
    return text if text else None


# Qualifiers that name a host's own constant library. Membership is resolved
# against the CURRENT host's model, so `Word.wdRed` answers in a Word module and
# misses in an Excel one: Excel-specific names never leak into another host
# (XLIDE issue #24).
_HOST_QUALIFIERS = frozenset({"excel", "word", "powerpoint", "access", "office"})
_SIGNED_RE = re.compile(r"^([+-])(.+)$", re.DOTALL)


def external_integer_constant_value(name: str, model: HostObjectModel | None = None) -> int | None:
    """The integer value of a VBA runtime or host constant named `name` (lowercased,
    optionally `vba.` or host qualified), or None when it names none or the runtime
    and host disagree."""
    dot = name.find(".")
    if dot >= 0:
        qualifier = name[:dot]
        member = name[dot + 1 :]
        if qualifier == "vba":
            runtime = resolve_runtime_constant(member)
            return numeric_external_constant_value(runtime.get("value") if runtime else None)
        if qualifier in _HOST_QUALIFIERS:
            host = resolve_host_constant(member, model)
            return numeric_external_constant_value(host.get("value") if host else None)
        return None
    runtime = resolve_runtime_constant(name)
    host = resolve_host_constant(name, model)
    candidates = {
        value
        for value in (
            numeric_external_constant_value(runtime.get("value") if runtime else None),
            numeric_external_constant_value(host.get("value") if host else None),
        )
        if value is not None
    }
    return next(iter(candidates)) if len(candidates) == 1 else None


def numeric_external_constant_value(value: str | int | None) -> int | None:
    """An external constant's value as a safe integer, or None when it is not one."""
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return safe_integer(value)
    if not isinstance(value, str):
        return None
    text = value.strip()
    signed = _SIGNED_RE.match(text)
    if signed is not None:
        parsed = parse_vba_integer_literal(signed.group(2))
        if parsed is None:
            return None
        return safe_integer(-parsed if signed.group(1) == "-" else parsed)
    return parse_vba_integer_literal(text)
