"""Rule: members an Implements statement requires (XLIDE issue #125).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/implementsMembers.ts.

A class that says `Implements IFoo` must supply every member of IFoo as
`IFoo_Member`, with the interface's parameter list. Measured in Excel 16.0 (build
20326, 2026-09-25):

- implements-member-missing: `Implements IFoo` in a class with no `IFoo_Name` for
  IFoo's `Name` -> "Object module needs to implement 'Name' for interface 'IFoo'".
- implements-member-signature: `Private Function IFoo_Name(ByVal extra As Long) As
  String` against `Public Function Name() As String` -> "Procedure declaration
  does not match description of event or procedure having the same name".

The interface is read from the project index, so only a class module of this
project is judged; an interface from a type library is not.
"""

from __future__ import annotations

import re
from collections.abc import Sequence
from dataclasses import dataclass

from ...conditional import ConditionalActivityTracker
from ...js_compat import JS_WHITESPACE, js_trim
from ...lexer.token_kinds import VbaToken
from ...parser.nodes import ModuleNode, StatementNode
from ...symbols.symbol_model import (
    ModuleSymbolKind,
    ModuleSymbols,
    VbaProjectClassMember,
    VbaProjectClassMembers,
    VbaSymbol,
    VbaSymbolKind,
    is_procedure_kind,
    procedure_params_from_symbol,
)
from ...types.type_names import is_known_scalar_type, normalize_type
from ..context import PushFn, is_object_module_kind
from ..walker import (
    absolute_span,
    active_module_members,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)

# The signature label is read with upstream's regular expressions, written here
# with JavaScript's meaning. JavaScript's `\s` takes the ECMAScript WhiteSpace and
# LineTerminator characters (JS_WHITESPACE), where Python's also takes
# U+001C-U+001F and U+0085 and leaves out U+FEFF; its `.` (no s flag) stops at
# every line terminator, where Python's stops only at "\n"; its `$` (no m flag) is
# the end of the input, where Python's also matches before a final "\n"; and its i
# flag (no u flag) folds case within ASCII only.
_S = "[" + JS_WHITESPACE + "]"
_DOT = "[^\n\r" + chr(0x2028) + chr(0x2029) + "]"
_FLAGS = re.IGNORECASE | re.ASCII

# /^(?:(?:Optional|ByVal|ByRef|ParamArray)\s+)+/i
_PASSING_PREFIX_RE = re.compile(rf"^(?:(?:Optional|ByVal|ByRef|ParamArray){_S}+)+", _FLAGS)
# /\s*=.*$/
_DEFAULT_VALUE_RE = re.compile(rf"{_S}*={_DOT}*\Z")
# /\sAs\s+(.+)$/i
_AS_CLAUSE_RE = re.compile(rf"{_S}As{_S}+({_DOT}+)\Z", _FLAGS)
# /\sAs\s.*$/i
_AS_TAIL_RE = re.compile(rf"{_S}As{_S}{_DOT}*\Z", _FLAGS)
# /\(\s*\)/
_EMPTY_PARENS_RE = re.compile(rf"\({_S}*\)")
# /\)\s*As\s+(.+)$/i
_RETURNS_RE = re.compile(rf"\){_S}*As{_S}+({_DOT}+)\Z", _FLAGS)


@dataclass(frozen=True, slots=True)
class _ParsedParam:
    type_: str
    is_array: bool


@dataclass(frozen=True, slots=True)
class _ParsedSignature:
    params: list[_ParsedParam]
    returns: str


def check_implements_members(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    module_kind: ModuleSymbolKind,
    project_class_members: Sequence[VbaProjectClassMembers] | None,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """Report each interface member an `Implements` class leaves out or declares
    with another signature."""
    if (
        not is_object_module_kind(module_kind)
        or project_class_members is None
        or len(project_class_members) == 0
    ):
        return
    procedures: dict[str, list[VbaSymbol]] = {}
    for symbol in symbols.root.children or []:
        if is_procedure_kind(symbol.kind):
            procedures.setdefault(symbol.name.lower(), []).append(symbol)
    for member in active_module_members(mod, activity):
        if not isinstance(member, StatementNode):
            continue
        toks = statement_tokens_after_leading_label(source, member.span)
        if token_text(_first_token(toks)) != "implements":
            continue
        # `Implements Lib.IFoo` names the interface last.
        name_token = next((tok for tok in reversed(toks) if token_name(tok) is not None), None)
        interface_name = token_name(name_token) if name_token is not None else None
        if not interface_name or name_token is None:
            continue
        contract = _contract_named(project_class_members, interface_name)
        if contract is None:
            continue
        for required in contract.members:
            # A Friend member is no part of the interface (issue #291, measured).
            if required.kind == "event" or required.visibility == "Friend":
                continue
            implementations = procedures.get(f"{contract.name}_{required.name}".lower(), [])
            variable = (
                _variable_implementation_problem(required, implementations)
                if required.kind == "property" and required.procedure_params is None and len(implementations) > 0
                else None
            )
            if variable is not None:
                if variable.missing:
                    push(
                        "implementsMemberMissing",
                        f"Object module needs to implement '{required.name}' for interface "
                        f"'{contract.name}': {variable.message}.",
                        absolute_span(member.span, name_token),
                    )
                else:
                    push(
                        "implementsMemberSignature",
                        f"'{variable.at.name}' does not match '{contract.name}.{required.name}': "
                        f"{variable.message}. The procedure declaration must match the interface "
                        "member it implements.",
                        variable.at.name_span,
                    )
                continue
            if len(implementations) == 0:
                push(
                    "implementsMemberMissing",
                    f"Object module needs to implement '{required.name}' for interface "
                    f"'{contract.name}': add {_expected_procedure_label(contract.name, required)}.",
                    absolute_span(member.span, name_token),
                )
                continue
            # A readable and writable property (a Public variable of the interface,
            # or a Get with a Let or Set) needs both accessors; Excel refuses the
            # project with the Get alone (XLIDE issue #144, measured).
            if required.kind == "property" and required.writable and required.returns:
                has_get = any(
                    impl.kind is VbaSymbolKind.PROPERTY_GET for impl in implementations
                )
                has_setter = any(
                    impl.kind is VbaSymbolKind.PROPERTY_LET or impl.kind is VbaSymbolKind.PROPERTY_SET
                    for impl in implementations
                )
                if not has_get or not has_setter:
                    push(
                        "implementsMemberMissing",
                        f"Object module needs to implement '{required.name}' for interface "
                        f"'{contract.name}': add a Property {'Let or Set' if has_get else 'Get'} "
                        f"'{contract.name}_{required.name}' beside the Property "
                        f"{'Get' if has_get else 'Let'}.",
                        absolute_span(member.span, name_token),
                    )
            for implementation in implementations:
                problem = _signature_mismatch(required, implementation)
                if problem is None:
                    problem = _passing_mismatch(required, implementation)
                if problem:
                    push(
                        "implementsMemberSignature",
                        f"'{implementation.name}' does not match '{contract.name}.{required.name}': "
                        f"{problem}. The procedure declaration must match the interface member it "
                        "implements.",
                        implementation.name_span,
                    )


def _contract_named(
    project_class_members: Sequence[VbaProjectClassMembers], interface_name: str
) -> VbaProjectClassMembers | None:
    lower = interface_name.lower()
    for candidate in project_class_members:
        if candidate.kind == "class" and candidate.exhaustive is True and candidate.name.lower() == lower:
            return candidate
    return None


def _passing_mismatch(required: VbaProjectClassMember, implementation: VbaSymbol) -> str | None:
    """What else the VBE matches between an interface procedure and its
    implementation (issue #291, measured in Excel 16.0): a Function is not
    implemented by a Sub, and each parameter keeps its passing, ByVal or ByRef (a
    plain one is ByRef), its Optional, and its default."""
    # Upstream's truthiness tests: an empty list or mapping is still present.
    declared = required.procedure_params
    if declared is None:
        return None
    if declared.get("function") is not None and implementation.kind is VbaSymbolKind.SUB:
        return "the interface member is a Function, and a Sub returns nothing"
    expected = declared.get(implementation.kind.value)
    if expected is None:
        return None
    actual = procedure_params_from_symbol(implementation, include_passing=True)
    for i in range(min(len(expected), len(actual))):
        want = expected[i]
        got = actual[i]
        if want.param_array or got.param_array:
            continue
        if bool(want.by_val) != bool(got.by_val):
            return (
                f"parameter {i + 1} is {'ByVal' if got.by_val else 'ByRef'} here and "
                f"{'ByVal' if want.by_val else 'ByRef'} on the interface"
            )
        if want.optional != got.optional:
            return (
                f"parameter {i + 1} is {'Optional' if got.optional else 'required'} here and "
                f"{'Optional' if want.optional else 'required'} on the interface"
            )
        if want.optional and js_trim(want.default_raw or "").lower() != js_trim(got.default_raw or "").lower():
            got_default = js_trim(got.default_raw) if got.default_raw is not None else ""
            want_default = js_trim(want.default_raw) if want.default_raw is not None else ""
            return (
                f"parameter {i + 1} defaults to {got_default or 'nothing'} here and "
                f"{want_default or 'nothing'} on the interface"
            )
    return None


@dataclass(frozen=True, slots=True)
class _VariableProblem:
    missing: bool
    message: str
    at: VbaSymbol


def _variable_implementation_problem(
    required: VbaProjectClassMember, implementations: Sequence[VbaSymbol]
) -> _VariableProblem | None:
    """A Public variable of the interface, implemented by Property procedures
    (issue #291, measured in Excel 16.0). A value type needs a Get and a Let whose
    value is ByVal and of the variable's type; an object type a Get and a Set whose
    value is ByVal; a Variant a Get, a Let and a Set, neither value ByVal."""
    type_ = _normalized_or_variant(
        required.write_type if required.write_type is not None else required.returns
    )
    is_object = type_ == "object" or (type_ != "variant" and not is_known_scalar_type(type_))

    def by_kind(kind: VbaSymbolKind) -> VbaSymbol | None:
        return next((impl for impl in implementations if impl.kind is kind), None)

    get = by_kind(VbaSymbolKind.PROPERTY_GET)
    letter = by_kind(VbaSymbolKind.PROPERTY_LET)
    setter = by_kind(VbaSymbolKind.PROPERTY_SET)
    name = implementations[0].name

    def needs(what: str) -> _VariableProblem:
        return _VariableProblem(True, f"add a Property {what} '{name}'", implementations[0])

    if get is None:
        return needs("Get")
    lacking = (
        (letter is None or setter is None)
        if type_ == "variant"
        else (setter is None if is_object else letter is None)
    )
    if lacking:
        if type_ == "variant":
            return needs("Set" if letter is not None else "Let")
        return needs("Set" if is_object else "Let")
    write_type_text = required.write_type if required.write_type is not None else "undefined"
    for procedure in (letter, setter):
        if procedure is None:
            continue
        params = procedure_params_from_symbol(procedure, include_passing=True)
        value = params[-1] if params else None
        if value is None:
            continue
        by_val = bool(value.by_val)
        if by_val if type_ == "variant" else not by_val:
            shown = (
                "Variant"
                if type_ == "variant"
                else _capitalize(required.write_type if required.write_type is not None else type_)
            )
            return _VariableProblem(
                False,
                f"its value is {'ByVal' if by_val else 'ByRef'}, and a Public {shown} of the "
                f"interface takes it {'ByRef' if by_val else 'ByVal'}",
                procedure,
            )
        if not is_object and type_ != "variant" and _normalized_or_variant(value.type_) != type_:
            return _VariableProblem(
                False,
                f"its value is {value.type_ if value.type_ is not None else 'Variant'}, and the "
                f"interface's variable is {write_type_text}",
                procedure,
            )
    if not is_object and type_ != "variant" and _normalized_or_variant(get.as_type) != type_:
        return _VariableProblem(
            False,
            f"it returns {get.as_type if get.as_type is not None else 'Variant'}, and the "
            f"interface's variable is {write_type_text}",
            get,
        )
    return None


def _capitalize(type_name: str) -> str:
    return type_name[:1].upper() + type_name[1:]


def _expected_procedure_label(interface_name: str, member: VbaProjectClassMember) -> str:
    name = f"{interface_name}_{member.name}"
    if member.kind == "property":
        return f"a Property {'Let' if member.writable and not member.returns else 'Get'} '{name}'"
    return f"a {'Function' if member.returns else 'Sub or Function'} '{name}'"


def _parse_signature(signature: str | None) -> _ParsedSignature | None:
    """The parameter list and return type an interface member's signature label states."""
    if not signature:
        return None
    open_index = signature.find("(")
    if open_index < 0:
        return None
    # Parentheses and commas inside a string default are text: `Optional sep As
    # String = ", "` is one parameter and `= ")"` does not end the list (XLIDE
    # issue #144).
    depth = 0
    close = -1
    in_string = False
    parts: list[str] = []
    part_start = open_index + 1
    for i in range(open_index, len(signature)):
        ch = signature[i]
        if ch == '"':
            in_string = not in_string
            continue
        if in_string:
            continue
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
            if depth == 0:
                close = i
                break
        elif ch == "," and depth == 1:
            parts.append(signature[part_start:i])
            part_start = i + 1
    if close < 0:
        return None
    parts.append(signature[part_start:close])
    inner = js_trim(signature[open_index + 1 : close])
    params = [] if len(inner) == 0 else [_parse_param(part) for part in parts]
    returns_match = _RETURNS_RE.search(signature[close:])
    return _ParsedSignature(
        params,
        _normalized_or_variant(js_trim(returns_match.group(1)) if returns_match else None),
    )


def _parse_param(part: str) -> _ParsedParam:
    # Project signature labels wrap optional parameters in brackets. Those delimiters
    # are presentation, not part of the type (otherwise Date becomes "Date]").
    label = js_trim(part)
    parameter = label[1:-1] if label.startswith("[") and label.endswith("]") else label
    text = _DEFAULT_VALUE_RE.sub("", _PASSING_PREFIX_RE.sub("", parameter, count=1), count=1)
    as_match = _AS_CLAUSE_RE.search(text)
    return _ParsedParam(
        type_=_normalized_or_variant(js_trim(as_match.group(1)) if as_match else None),
        is_array=_EMPTY_PARENS_RE.search(_AS_TAIL_RE.sub("", text, count=1)) is not None,
    )


def _signature_mismatch(required: VbaProjectClassMember, implementation: VbaSymbol) -> str | None:
    expected = _parse_signature(required.signature)
    if expected is None:
        return None
    params = [
        child for child in implementation.children or [] if child.kind is VbaSymbolKind.PARAMETER
    ]
    # A setter's last parameter is the value the interface property holds.
    compared = (
        params[:-1]
        if implementation.kind is VbaSymbolKind.PROPERTY_LET
        or implementation.kind is VbaSymbolKind.PROPERTY_SET
        else params
    )
    if len(compared) != len(expected.params):
        count = len(expected.params)
        return (
            f"the interface member takes {count} parameter{'' if count == 1 else 's'}, "
            f"this procedure {len(compared)}"
        )
    for i, param in enumerate(compared):
        actual_type = _normalized_or_variant(param.as_type)
        if actual_type != expected.params[i].type_ or bool(param.is_array) != expected.params[i].is_array:
            declared = param.as_type if param.as_type is not None else "Variant"
            return f"parameter {i + 1} is {declared} here and {expected.params[i].type_} on the interface"
    if implementation.kind is VbaSymbolKind.FUNCTION or implementation.kind is VbaSymbolKind.PROPERTY_GET:
        actual_return = _normalized_or_variant(implementation.as_type)
        if required.returns is not None and actual_return != expected.returns:
            declared = implementation.as_type if implementation.as_type is not None else "Variant"
            return f"it returns {declared} where the interface member returns {required.returns}"
    return None


def _normalized_or_variant(type_name: str | None) -> str:
    """`normalizeType(type) ?? 'variant'`: an empty normalized name stays empty."""
    normalized = normalize_type(type_name)
    return normalized if normalized is not None else "variant"


def _first_token(toks: Sequence[VbaToken]) -> VbaToken | None:
    """`toks[0]`, undefined (None) for an empty statement."""
    return toks[0] if toks else None
