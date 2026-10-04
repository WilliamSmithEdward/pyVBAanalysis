"""Rule: AddressOf where the VBE refuses it (XLIDE issue #299, each measured in
64-bit Excel 16.0 as a compile error).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/addressOfUse.ts.

AddressOf compiles only as a whole argument of a project procedure or of a
method such as Collection.Add, and only on a Sub, Function or Property of a
standard module. Elsewhere:

 - outside an argument, `p = AddressOf Cb`, in parentheses of its own,
   `Take((AddressOf Cb))`, in Debug.Print, or as the argument of Len, CLng,
   CLngPtr, CStr or Abs: "Syntax error"; of ObjPtr: "Type mismatch". VarPtr,
   StrPtr, Hex, IsEmpty and TypeName take it. With an operator after it,
   `Take(AddressOf Cb + 1)`: "Argument not optional";
 - on a name nothing declares, or a class's member, `AddressOf K.Run`:
   "Variable not defined" (under Option Explicit); on a variable: "Expected
   Sub, Function, or Property"; on a VBA function such as Len: "Syntax error";
   on a Declare: "Invalid use of AddressOf operator";
 - into a ByVal Long, Integer or Byte parameter in 64-bit VBA: "Type mismatch",
   since AddressOf gives a LongPtr (#298).
"""

from __future__ import annotations

import re
from collections.abc import Mapping, Sequence
from collections.abc import Set as AbstractSet
from dataclasses import dataclass

from ...conditional import ConditionalActivityTracker
from ...conditional.conditional_compilation import (
    ConditionalCompilationEnvironment,
    compiler_constants_with_defaults,
)
from ...lexer.token_kinds import VbaToken
from ...parser.nodes import LeafStatementNode, ModuleNode, ProcedureNode, Span
from ...runtime.vba_runtime import resolve_runtime_function
from ...symbols.symbol_model import ModuleSymbols, VbaProcedureSignature, VbaProjectClassMembers
from ...types.type_inference import procedure_symbol_for
from ...types.type_names import normalize_type
from ..call_extraction import CallableTypeSignature
from ..callable_signatures import callable_type_signatures_for
from ..context import PushFn, statement_tokens
from ..walker import (
    active_module_members,
    for_each_statement,
    statement_and_branch_spans,
    token_name,
    token_text,
)
from .malformed_lines import VALUE_WORD_ERRORS

_NARROW: frozenset[str] = frozenset({"long", "integer", "byte"})

# The VBA functions measured to refuse an AddressOf argument, and the error each gives.
_ADDRESSOF_REFUSED_BY: dict[str, str] = {
    "len": "Syntax error",
    "clng": "Syntax error",
    "clngptr": "Syntax error",
    "cstr": "Syntax error",
    "abs": "Syntax error",
    "objptr": "Type mismatch",
}

_ADDRESSOF_RE = re.compile(r"\bAddressOf\b", re.IGNORECASE | re.ASCII)
# JavaScript's multiline `^` also follows a lone CR and U+2028/U+2029.
_OPTION_EXPLICIT_RE = re.compile(
    "(?:\\A|(?<=[\\n\\r\u2028\u2029]))[ \\t]*Option[ \\t]+Explicit\\b", re.IGNORECASE | re.ASCII
)


def _at(toks: Sequence[VbaToken], index: int) -> VbaToken | None:
    return toks[index] if 0 <= index < len(toks) else None


def _raw(toks: Sequence[VbaToken], index: int) -> str | None:
    tok = _at(toks, index)
    return tok.raw_text if tok is not None else None


def check_address_of_use(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None,
    project_class_members: Sequence[VbaProjectClassMembers] | None,
    conditional_compilation: ConditionalCompilationEnvironment | None,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    if _ADDRESSOF_RE.search(source) is None:
        return
    win64 = compiler_constants_with_defaults(conditional_compilation).get("win64")
    if isinstance(win64, bool):
        is64 = win64 is True
    elif isinstance(win64, (int, float)):
        is64 = win64 != 0
    else:
        is64 = False
    explicit = _OPTION_EXPLICIT_RE.search(source) is not None
    signatures = callable_type_signatures_for(symbols, project_procedures)
    module_members: dict[str, str] = {
        child.name.lower(): str(child.kind.value) for child in symbols.root.children or []
    }
    standard_modules: dict[str, frozenset[str]] = {
        project_type.name.lower(): frozenset(m.name.lower() for m in project_type.members)
        for project_type in project_class_members or []
        if project_type.kind == "standardModule"
    }
    project_procedure_names = frozenset(name for names in standard_modules.values() for name in names)
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        proc_sym = procedure_symbol_for(symbols, member)
        locals_ = frozenset(
            [
                *(param.name.lower() for param in member.params),
                *(child.name.lower() for child in (proc_sym.children if proc_sym is not None else None) or []),
            ]
        )
        scope = _Scope(
            member=member.name.lower(),
            locals=locals_,
            module_members=module_members,
            standard_modules=standard_modules,
            project_procedure_names=project_procedure_names,
            signatures=signatures,
            is64=is64,
            explicit=explicit,
        )

        def each_statement(stmt: LeafStatementNode, scope: _Scope = scope) -> None:
            for span in statement_and_branch_spans(stmt):
                toks = statement_tokens(source, span)
                for i, tok in enumerate(toks):
                    if token_text(tok) != "addressof":
                        continue
                    problem = _address_of_problem(toks, i, scope)
                    if problem is None:
                        continue
                    why, error = problem
                    if _at(toks, i + 1) is not None and token_name(toks[i + 1]):
                        end = i + 3 if _raw(toks, i + 2) == "." and _at(toks, i + 3) is not None else i + 1
                    else:
                        end = i
                    push(
                        "addressOfMisuse",
                        f"{why} This is a VBE compile error: {error}.",
                        Span(span.start + tok.start, span.start + toks[end].end),
                    )

        for_each_statement(member.body, each_statement, activity)


@dataclass(frozen=True, slots=True)
class _Scope:
    member: str
    locals: AbstractSet[str]
    module_members: Mapping[str, str]
    standard_modules: Mapping[str, AbstractSet[str]]
    project_procedure_names: AbstractSet[str]
    signatures: Mapping[str, CallableTypeSignature]
    is64: bool
    explicit: bool


def _address_of_problem(toks: Sequence[VbaToken], at: int, scope: _Scope) -> tuple[str, str] | None:
    """(why, error) for the AddressOf at toks[at], or None where the VBE takes it."""
    first = _at(toks, at + 1)
    name = token_name(first)
    if not name:
        return None
    qualified = _raw(toks, at + 2) == "." and token_name(_at(toks, at + 3)) is not None
    last = at + 3 if qualified else at + 1
    text = "".join(tok.raw_text for tok in toks[at + 1 : last + 1])
    # Where it stands: the open parenthesis or the call statement it is an argument of.
    depth = 0
    open_index = -1
    for k in range(at - 1, -1, -1):
        raw = toks[k].raw_text
        if raw == ")":
            depth += 1
        elif raw == "(":
            if depth == 0:
                open_index = k
                break
            depth -= 1
    before = _at(toks, at - 1)
    after = _at(toks, last + 1)
    # A call statement's argument: `Take AddressOf Cb`, `c.Add AddressOf Cb`.
    callee_end = next(
        (
            k
            for k, tok in enumerate(toks)
            if k > 0 and (tok.raw_text != "." if k % 2 == 1 else token_name(tok) is None)
        ),
        -1,
    )
    chain_end = len(toks) if callee_end < 0 else callee_end
    statement_argument = (
        open_index < 0
        and at >= 1
        and token_name(toks[0]) is not None
        and not any(k < at and tok.raw_text == "=" for k, tok in enumerate(toks))
        and (at == chain_end or (before is not None and before.raw_text == "," and at > chain_end))
    )
    # `Debug.Print AddressOf Cb`: Print takes no AddressOf, whatever it names.
    if (
        open_index < 0
        and token_text(toks[0]) == "debug"
        and _raw(toks, 1) == "."
        and token_text(_at(toks, 2)) == "print"
    ):
        return (f"Debug.Print takes no 'AddressOf {text}'.", "Syntax error")
    starts_slot = (before is not None and before.raw_text in ("(", ",")) or statement_argument
    if not starts_slot or (open_index < 0 and not statement_argument):
        return (
            f"AddressOf can stand only as an argument, and 'AddressOf {text}' is not one.",
            "Syntax error",
        )
    if after is not None and after.raw_text != "," and after.raw_text != ")":
        return (
            f"'AddressOf {text}' must be the whole argument, and an operator follows it.",
            "Argument not optional",
        )
    callee = _at(toks, open_index - 1) if open_index >= 0 else toks[0]
    callee_raw_name = token_name(callee)
    callee_name = callee_raw_name.lower() if callee_raw_name else None
    if open_index >= 0 and (not callee_name or _raw(toks, open_index - 1) == "("):
        return (f"'AddressOf {text}' cannot stand in parentheses of its own.", "Syntax error")
    member = _raw(toks, open_index - 2) == "." if open_index >= 0 else _raw(toks, 1) == "."
    # Some VBA functions take it and some do not: VarPtr, StrPtr, Hex, IsEmpty and
    # TypeName compile; these do not (measured in Excel 16.0).
    refused = (
        _ADDRESSOF_REFUSED_BY.get(callee_name)
        if callee_name and not member and callee_name not in scope.signatures
        else None
    )
    if refused and callee is not None:
        return (f"{callee.raw_text} takes no 'AddressOf {text}'.", refused)
    # What it names.
    if qualified:
        procedures = scope.standard_modules.get(name.lower())
        procedure = (token_name(toks[at + 3]) or "").lower()
        if procedures is None or procedure not in procedures:
            return (
                (
                    f"'{text}' names no procedure of a standard module, and AddressOf takes only those.",
                    "Variable not defined",
                )
                if scope.explicit
                else None
            )
    else:
        lower = name.lower()
        kind = scope.module_members.get(lower)
        if lower in scope.locals or kind == "moduleVariable" or kind == "constant":
            return (
                f"'{name}' is a variable, and AddressOf takes a Sub, Function or Property.",
                "Expected Sub, Function, or Property",
            )
        if kind == "declare":
            return (
                f"'{name}' is a Declare, whose address AddressOf cannot take.",
                "Invalid use of AddressOf operator",
            )
        is_procedure = (
            kind == "sub"
            or kind == "function"
            or (kind is not None and kind.startswith("property"))
            or lower in scope.project_procedure_names
        )
        if not is_procedure:
            # Len, Array and the other reserved words are reserved-keyword-in-expression's.
            if lower in VALUE_WORD_ERRORS:
                return None
            if resolve_runtime_function(lower) is not None:
                return (
                    f"'{name}' is a VBA function, and AddressOf takes only the project's procedures.",
                    "Invalid use of AddressOf operator",
                )
            return (
                (f"'{name}' names no procedure of the project.", "Variable not defined")
                if scope.explicit
                else None
            )
    # AddressOf gives a LongPtr, which a ByVal Long takes no more than a LongLong (#298).
    signature = scope.signatures.get(callee_name) if callee_name and not member else None
    if signature is not None and scope.is64:
        position = 0
        for k in range(open_index + 1 if open_index >= 0 else 1, at):
            if toks[k].raw_text == ",":
                position += 1
        param = signature.params[position] if position < len(signature.params) else None
        param_type = normalize_type(param.type_) if param is not None else None
        if param is not None and not param.by_ref and param_type and param_type in _NARROW:
            return (
                f"'AddressOf {text}' is a LongPtr in 64-bit VBA, and the ByVal {param.type_} "
                f"'{param.name}' of '{signature.name}' takes no LongPtr.",
                "Type mismatch",
            )
    return None
