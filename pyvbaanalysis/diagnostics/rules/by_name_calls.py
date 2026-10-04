"""Rule family: a call made by a name in a string (XLIDE issue #243).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/byNameCalls.ts. Measured
in Excel 16.0 (build 20326, 2026-09-30); each compiles and raises every time.

 - `CallByName(c, "NoSuch", VbMethod)` with c a project class that has no
   Public member of that name, a Private one included -> 438. On a Collection,
   any name but Add, Count, Item and Remove -> 438.
 - `CallByName c, "Hello", VbLet, 1` with Hello a Function, or VbGet on a Sub
   or Function -> 450: the call type does not fit the member.
 - In Excel, `Application.Run "NoSuch"` with no Sub or Function of that name in
   a standard or document module of the project -> 1004. Private ones count; a
   class module's members do not.

XLIDE issue #408, measured in Excel 16.0 on 2026-10-02: the bare `Run` is
Application.Run; too many arguments for the procedure named raise 450 and too
few 449, through Run and CallByName alike; VbLet on a Property Get alone raises
451; and a Collection's four members are methods, so any call type but VbMethod
raises 438.

Names compare without case. A name built at run time is not judged.
"""

from __future__ import annotations

import math
import re
from collections.abc import Mapping, Sequence
from collections.abc import Set as AbstractSet

from ...completion.member_access import MemberCompletionContext
from ...conditional import ConditionalActivityTracker
from ...js_compat import js_trim
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import LeafStatementNode, ModuleNode, ProcedureNode, Span
from ...symbols.symbol_model import ModuleSymbols
from ...types.type_inference import type_environment_for
from ...types.type_names import normalize_type
from ..call_extraction import string_literal_value
from ..context import PushFn, statement_tokens
from ..walker import (
    active_module_members,
    for_each_statement,
    match_paren_from,
    statement_and_branch_spans,
    token_name,
    token_text,
)

_COLLECTION_MEMBERS: frozenset[str] = frozenset({"add", "count", "item", "remove"})

# VbCallType, by its name and its value.
_CALL_TYPES: dict[str, str] = {
    "vbmethod": "method", "vbget": "get", "vblet": "let", "vbset": "set",
    "1": "method", "2": "get", "4": "let", "8": "set",
}


def _raw(toks: Sequence[VbaToken], index: int) -> str | None:
    return toks[index].raw_text if 0 <= index < len(toks) else None


def _text(toks: Sequence[VbaToken], index: int) -> str:
    return token_text(toks[index]) if 0 <= index < len(toks) else ""


def check_by_name_calls(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    member_ctx: MemberCompletionContext,
    runnable: AbstractSet[str] | None,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    module_names = {child.name.lower() for child in symbols.root.children or []}
    # Word's Run also reaches the global templates, Normal.dotm among them, which
    # the project cannot see; only Excel's is judged.
    host_name = member_ctx.model.get("hostName") if member_ctx.model is not None else None
    excel = (host_name if host_name is not None else "Excel") == "Excel"
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        env = type_environment_for(symbols, member)

        def each_statement(stmt: LeafStatementNode, env: Mapping[str, str] = env) -> None:
            for span in statement_and_branch_spans(stmt):
                toks = statement_tokens(source, span)
                for i in range(len(toks)):
                    word = token_text(toks[i])
                    if word == "callbyname" and _raw(toks, i - 1) != "." and "callbyname" not in module_names:
                        _check_call_by_name(span, toks, i, env, member_ctx, push)
                    elif (
                        word == "run"
                        and runnable is not None
                        and excel
                        and (
                            (
                                _raw(toks, i - 1) == "."
                                and _text(toks, i - 2) == "application"
                                and _raw(toks, i - 3) != "."
                            )
                            # Excel's global Run is Application.Run (XLIDE issue #408).
                            or (_raw(toks, i - 1) != "." and "run" not in module_names and "run" not in env)
                        )
                    ):
                        _check_application_run(span, toks, i, runnable, member_ctx, push)

        for_each_statement(member.body, each_statement, activity)


def _arguments_after(toks: Sequence[VbaToken], name: int) -> list[list[VbaToken]]:
    """The arguments after `toks[name]`: in parentheses, or to the end of a call statement."""
    paren = _raw(toks, name + 1) == "("
    close = match_paren_from(toks, name + 1) if paren else len(toks)
    out: list[list[VbaToken]] = []
    current: list[VbaToken] = []
    depth = 0
    for i in range(name + (2 if paren else 1), close):
        raw = toks[i].raw_text
        depth += 1 if raw == "(" else -1 if raw == ")" else 0
        if raw == "," and depth == 0:
            out.append(current)
            current = []
            continue
        current.append(toks[i])
    out.append(current)
    return out


def _span_of(base: Span, arg: Sequence[VbaToken]) -> Span:
    return Span(base.start + arg[0].start, base.start + arg[-1].end)


def _check_call_by_name(
    span: Span,
    toks: Sequence[VbaToken],
    at: int,
    env: Mapping[str, str],
    member_ctx: MemberCompletionContext,
    push: PushFn,
) -> None:
    args = _arguments_after(toks, at)
    obj = args[0] if len(args) > 0 else None
    proc_name = args[1] if len(args) > 1 else None
    call_type = args[2] if len(args) > 2 else None
    if (
        not obj
        or len(obj) != 1
        or proc_name is None
        or len(proc_name) != 1
        or proc_name[0].kind is not TokenKind.STRING_LITERAL
        or call_type is None
        or len(call_type) != 1
    ):
        return
    object_name = token_name(obj[0])
    declared = env.get(object_name.lower() if object_name else "")
    name = string_literal_value(proc_name[0].raw_text)
    lower = name.lower()
    kind = _CALL_TYPES.get(token_text(call_type[0]))
    if not declared or not kind:
        return
    if normalize_type(declared) == "collection":
        if lower not in _COLLECTION_MEMBERS:
            push(
                "runtimeMemberNotFound",
                f"CallByName asks Collection '{obj[0].raw_text}' for '{name}', which it does not have. "
                "This will raise Run-time error '438': Object doesn't support this property or method.",
                _span_of(span, proc_name),
            )
        elif kind != "method":
            # Add, Count, Item and Remove are all methods (XLIDE issue #408, measured).
            push(
                "runtimeMemberNotFound",
                f"CallByName asks Collection '{obj[0].raw_text}' for '{name}' with "
                f"{call_type[0].raw_text}, but {name} is a method, which only VbMethod reaches. "
                "This will raise Run-time error '438': Object doesn't support this property or method.",
                _span_of(span, call_type),
            )
        return
    surface = next(
        (
            candidate
            for candidate in member_ctx.project_class_members or []
            if candidate.kind == "class" and candidate.name.lower() == declared.lower()
        ),
        None,
    )
    if surface is None:
        return
    found = next((candidate for candidate in surface.members if candidate.name.lower() == lower), None)
    if found is None:
        push(
            "runtimeMemberNotFound",
            f"CallByName asks '{obj[0].raw_text}', a {surface.name}, for '{name}', which is no Public "
            f"member of {surface.name}. This will raise Run-time error '438': Object doesn't support "
            "this property or method.",
            _span_of(span, proc_name),
        )
        return
    if found.kind == "method" and kind != "method":
        how = (
            call_type[0].raw_text
            if token_text(call_type[0]).startswith("vb")
            else f"call type {call_type[0].raw_text}"
        )
        push(
            "runtimeMemberNotFound",
            f"CallByName reaches {surface.name}.{found.name}, a "
            f"{'Function' if found.returns else 'Sub'}, with {how}, which only a property takes. "
            "This will raise Run-time error '450': Wrong number of arguments or invalid property "
            "assignment.",
            _span_of(span, call_type),
        )
        return
    # `CallByName c, "P", VbLet, 5` with P a Property Get alone (XLIDE issue #408).
    if (
        found.kind == "property"
        and kind == "let"
        and not found.let_accessor
        and not found.set_accessor
        and found.writable is not True
        and found.signature is not None
    ):
        push(
            "runtimeMemberNotFound",
            f"CallByName assigns {surface.name}.{found.name}, which has a Property Get and no "
            "Property Let. This will raise Run-time error '451': Property let procedure not defined "
            "and property get procedure did not return an object.",
            _span_of(span, call_type),
        )
        return
    if found.kind == "method" and found.signature:
        problem = _argument_count_problem(found.signature, len(args) - 3, f"{surface.name}.{found.name}")
        if problem:
            push("runtimeMemberNotFound", f"CallByName {problem}", _span_of(span, proc_name))


_PARAM_ARRAY_HEAD_RE = re.compile(r"^paramarray\b", re.IGNORECASE | re.ASCII)
_PARAM_ARRAY_RE = re.compile(r"paramarray", re.IGNORECASE | re.ASCII)


def _parameter_counts(signature: str) -> tuple[int, float] | None:
    """The parameters a member signature lists: how many a call must pass, and may."""
    open_index = signature.find("(")
    if open_index < 0:
        return None
    depth = 0
    close = -1
    for i in range(open_index, len(signature)):
        depth += 1 if signature[i] == "(" else -1 if signature[i] == ")" else 0
        if depth == 0:
            close = i
            break
    if close < 0:
        return None
    listed = js_trim(signature[open_index + 1 : close])
    params = [] if listed == "" else [js_trim(param) for param in listed.split(",")]
    required = len(
        [param for param in params if not param.startswith("[") and _PARAM_ARRAY_HEAD_RE.match(param) is None]
    )
    maximum: float = math.inf if any(_PARAM_ARRAY_RE.search(param) for param in params) else len(params)
    return required, maximum


def _argument_count_problem(signature: str, given: int, shown: str) -> str | None:
    """Too many arguments for a procedure called by name raise 450, too few 449
    (XLIDE issue #408, measured in Excel 16.0 through Application.Run and
    CallByName). The message reads after "Application.Run " or "CallByName "."""
    counts = _parameter_counts(signature)
    if counts is None:
        return None
    required, maximum = counts
    plural = "" if given == 1 else "s"
    if given > maximum:
        return (
            f"passes {given} argument{plural} to {shown}, which takes {int(maximum)}. This will "
            "raise Run-time error '450': Wrong number of arguments or invalid property assignment."
        )
    if given < required:
        return (
            f"passes {given} argument{plural} to {shown}, which needs {required}. This will raise "
            "Run-time error '449': Argument not optional."
        )
    return None


_A1_RE = re.compile(r"^([A-Za-z]{1,3})([0-9]+)\Z")
_R1C1_RE = re.compile(r"^(?:R([0-9]+)C([0-9]+)|R([0-9]+)|C([0-9]+))\Z", re.IGNORECASE | re.ASCII)


def _reads_as_cell_address(name: str) -> bool:
    """Whether a name is a cell address in A1 (`Pub2`, `XFD1`) or R1C1 (`R1C1`,
    `R2`, `C3`) style."""
    a1 = _A1_RE.match(name)
    if a1 is not None:
        column = 0
        for c in a1.group(1).upper():
            column = column * 26 + ord(c) - 64
        row = int(a1.group(2))
        if column <= 16384 and 1 <= row <= 1048576:
            return True
    r1c1 = _R1C1_RE.match(name)
    if r1c1 is None:
        return False
    row_text = r1c1.group(1) if r1c1.group(1) is not None else r1c1.group(3)
    column_text = r1c1.group(2) if r1c1.group(2) is not None else r1c1.group(4)
    row = int(row_text) if row_text is not None else 1
    column = int(column_text) if column_text is not None else 1
    return 1 <= row <= 1048576 and 1 <= column <= 16384


_MACRO_NAME_RE = re.compile(r"^[A-Za-z][A-Za-z0-9_]*(?:\.[A-Za-z][A-Za-z0-9_]*)?\Z")


def _check_application_run(
    span: Span,
    toks: Sequence[VbaToken],
    at: int,
    runnable: AbstractSet[str],
    member_ctx: MemberCompletionContext,
    push: PushFn,
) -> None:
    args = _arguments_after(toks, at)
    macro = args[0] if args else None
    if macro is None or len(macro) != 1 or macro[0].kind is not TokenKind.STRING_LITERAL:
        return
    name = string_literal_value(macro[0].raw_text)
    # Another workbook's macro, `Book1.xlsm!Macro`, or a quoted name, is not this
    # project's to judge.
    if _MACRO_NAME_RE.match(name) is None:
        return
    if name.lower() in runnable:
        # The arguments against the one Public procedure of that name.
        module_name: str | None
        if "." in name:
            module_name, procedure = name.lower().split(".")
        else:
            module_name, procedure = None, name.lower()
        candidates = [
            project_member
            for project_type in member_ctx.project_class_members or []
            if project_type.kind == "standardModule"
            and (module_name is None or project_type.name.lower() == module_name)
            for project_member in project_type.members
            if project_member.kind == "method" and project_member.name.lower() == procedure
        ]
        # A bare name that reads as a cell address, `Pub2` or `R1C1`, is taken as
        # the address (XLIDE issue #468, measured in Excel 16.0).
        if module_name is None and _reads_as_cell_address(name):
            owner = candidates[0].module_name if len(candidates) == 1 else None
            hint = f' Name it with its module: "{owner}.{name}".' if owner else ""
            push(
                "runtimeMemberNotFound",
                f"Application.Run reads '{name}' as a cell address, not as the procedure of that "
                f"name. This will raise Run-time error '1004': Cannot run the macro '{name}'.{hint}",
                _span_of(span, macro),
            )
            return
        signature = candidates[0].signature if len(candidates) == 1 else None
        problem = _argument_count_problem(signature, len(args) - 1, name) if signature else None
        if problem:
            push("runtimeMemberNotFound", f"Application.Run {problem}", _span_of(span, macro))
        return
    push(
        "runtimeMemberNotFound",
        f"Application.Run names '{name}', and no standard or document module of the project has a "
        f"Sub or Function of that name. This will raise Run-time error '1004': Cannot run the macro "
        f"'{name}'.",
        _span_of(span, macro),
    )
