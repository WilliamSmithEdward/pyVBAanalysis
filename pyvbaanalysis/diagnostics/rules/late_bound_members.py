"""Rule: a member the VBE binds at run time, on an object whose class the code
makes plain and whose member list is complete (XLIDE issue #121).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/lateBoundMembers.ts.
Measured in Excel 16.0 (build 20326, 2026-09-26); each compiles and raises 438,
"Object doesn't support this property or method", every time it runs.

 - `Application.Zzq`: Application is extensible, so the VBE compiles any
   name on it (worksheet functions such as Application.Match are ordinary
   VBA there), and a name that is neither an Application member nor a
   WorksheetFunction raises when it runs. Excel only: its model lists every
   member, hidden ones included.
 - `Dim o As Object: Set o = New Collection: o.Foo`: a late-bound variable
   holding a class with a known member list. Collection has Add, Count,
   Item and Remove; a project class module has its public members.

XLIDE issue #224 (measured in Excel 16.0): the class also reaches the variable
from one declared as it, `Set o = c`, where c raises 91 instead while it is
Nothing. A Private member is not on the list (438). A property with a Get and no
Let raises 451 when assigned, and one with a Let and no Get 450 when read.
"""

from __future__ import annotations

from ...symbols.symbol_model import VbaSymbolKind, SymbolVisibility

import dataclasses
import re
from collections.abc import Callable, Mapping, Sequence, Iterable
from collections.abc import Set as AbstractSet
from dataclasses import dataclass, field
from typing import Any, Literal

from ...completion.member_access import (
    MemberCompletionContext,
    project_type_at,
    resolve_receiver_type_at,
)
from ...conditional import ConditionalActivityTracker
from ...flow.procedure_labels import jump_target_label_declaration
from ...host.host_model import (
    HostObjectModel,
    get_excel_object_model,
    get_host_members,
    get_host_type,
)
from ...js_compat import JS_WHITESPACE, js_number, js_number_to_string, js_trim
from ...lexer.token_helpers import match_paren_from, split_top_level_token_groups
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import (
    BodyNode,
    ForBlockNode,
    LeafStatementNode,
    ModuleNode,
    ProcedureNode,
    Span,
    StatementNode,
    is_leaf_statement,
    iter_body_nodes,
)
from ...symbols.symbol_model import ModuleSymbols, VbaProjectClassMembers
from ...types.type_inference import procedure_symbol_for, type_environment_for
from ...types.type_names import is_known_scalar_type, normalize_type
from ..argument_inference import infer_expression_type
from ..call_extraction import CallableTypeSignature, InferredArgumentType, string_literal_value
from ..callable_signatures import (
    SourceNameScope,
    build_module_type_signatures,
    source_name_scope_for,
)
from ..context import PushFn, statement_tokens
from ..dataflow import BlockEnteringState, walk_entering_blocks
from ..held_objects import HELD_VALUE, held_objects_at
from ..walker import (
    active_module_members,
    for_each_statement,
    raw_expression_tokens,
    set_assignment_target,
    statement_and_branch_spans,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)
from .late_bound_objects import vbscript_pattern_error
from .shared import body_may_leave_loop, names_in
from .type_of_is import object_assignment_incompatibility_reason

_COLLECTION_MEMBERS: frozenset[str] = frozenset({"add", "count", "item", "remove"})

_JS_SPACE = f"[{re.escape(JS_WHITESPACE)}]"


def _at(toks: Sequence[VbaToken], index: int) -> VbaToken | None:
    """toks[index], or None past either end (JavaScript's `toks[i]?.`)."""
    return toks[index] if 0 <= index < len(toks) else None


def _raw(toks: Sequence[VbaToken], index: int) -> str | None:
    tok = _at(toks, index)
    return tok.raw_text if tok is not None else None


def _kind(toks: Sequence[VbaToken], index: int) -> TokenKind | None:
    tok = _at(toks, index)
    return tok.kind if tok is not None else None


@dataclass(frozen=True, slots=True)
class _KnownParam:
    name: str
    optional: bool
    param_array: bool


@dataclass(frozen=True, slots=True)
class _KnownClass:
    """Names a late-bound local is known to hold: the class display name and its members."""

    display: str
    members: AbstractSet[str]
    # Properties with a Get and no Let or Set: assigning one raises 451.
    read_only: AbstractSet[str] | None = None
    # Properties with a Let and no Get: reading one raises 450.
    write_only: AbstractSet[str] | None = None
    # Properties with only a Set: reading one raises 450 too (XLIDE issue #414).
    set_only: AbstractSet[str] | None = None
    # Properties with a Get and a Set and no Let: a Let of one raises 438 (issue #685).
    no_let: AbstractSet[str] | None = None
    # Subs: assigning one raises 450, reading one with arguments 451 (issue #414).
    subs: AbstractSet[str] | None = None
    # Fields of a scalar type, by lowercased name: a member of one raises 424 (issue #414).
    scalar_fields: Mapping[str, str] | None = None
    # Set from a variable that may still be Nothing: 91 before 438.
    may_be_nothing: bool = False
    # The parameters of each method, by lowercased name, where they are known (issue #485).
    params: Mapping[str, Sequence[_KnownParam]] | None = field(default=None)
    # A RegExp's pattern, where the code set it to a literal (issue #477).
    pattern: str | None = None


# VBScript's RegExp, as CreateObject("VBScript.RegExp") gives it (XLIDE issue #477).
_REGEXP_CLASS = _KnownClass(
    display="RegExp",
    members=frozenset({"pattern", "global", "ignorecase", "multiline", "test", "execute", "replace"}),
)

# ProgIDs everyday macros create, lowercased. A ProgID one letter away from one of
# these, with the same parts, is taken as a misspelling of it, and the bare last
# part ("Dictionary") is no ProgID at all (XLIDE issue #477, measured in Excel 16.0:
# each raises 429).
_KNOWN_PROGIDS: tuple[str, ...] = (
    "scripting.dictionary", "scripting.filesystemobject", "vbscript.regexp", "wscript.shell",
    "wscript.network", "shell.application",
    "adodb.connection", "adodb.recordset", "adodb.command", "adodb.stream",
    "msxml2.domdocument", "msxml2.domdocument.3.0", "msxml2.domdocument.4.0",
    "msxml2.domdocument.5.0", "msxml2.domdocument.6.0",
    "msxml2.xmlhttp", "msxml2.xmlhttp.3.0", "msxml2.xmlhttp.6.0", "msxml2.serverxmlhttp",
    "msxml2.serverxmlhttp.6.0", "winhttp.winhttprequest.5.1",
    "excel.application", "word.application", "powerpoint.application", "outlook.application",
    "access.application",
)


def _prog_id_problem(prog_id: str) -> str | None:
    """Why a ProgID literal names nothing, or None when it may be registered."""
    lower = js_trim(prog_id).lower()
    if lower == "":
        return "an empty ProgID names no class"
    if lower in _KNOWN_PROGIDS:
        return None
    bare = next(
        (
            known
            for known in _KNOWN_PROGIDS
            if "." not in lower and len(known.split(".")) > 1 and known.split(".")[1] == lower
        ),
        None,
    )
    if bare is not None:
        return f'"{prog_id}" lacks its library: the ProgID is "{bare}"'
    near = next(
        (
            known
            for known in _KNOWN_PROGIDS
            if len(known.split(".")) == len(lower.split(".")) and _one_letter_apart(known, lower)
        ),
        None,
    )
    return f'"{prog_id}" is one letter away from "{near}", and no class has that ProgID' if near else None


def _one_letter_apart(a: str, b: str) -> bool:
    """Whether two strings differ by one inserted, deleted or changed letter."""
    if abs(len(a) - len(b)) > 1 or a == b:
        return False
    i = 0
    while i < len(a) and i < len(b) and a[i] == b[i]:
        i += 1

    def letter(text: str, index: int) -> bool:
        return index >= len(text) or "a" <= text[index] <= "z"

    if len(a) == len(b):
        return letter(a, i) and letter(b, i) and a[i + 1 :] == b[i + 1 :]
    long, short = (a, b) if len(a) > len(b) else (b, a)
    return letter(long, i) and long[i + 1 :] == short[i:]


def _reg_exp_pattern_problem(pattern: str) -> tuple[str, str] | None:
    """What VBScript's RegExp refuses in a pattern, with the error it raises (XLIDE
    issue #477, measured in Excel 16.0): an unclosed group (5020), an unclosed
    class (5019), a quantifier with nothing to repeat (5018), and any other fault,
    a lookbehind or a named group among them (5017). (error, text)."""
    code = vbscript_pattern_error(pattern)
    if code == 5020:
        return (
            "'5020': Application-defined or object-defined error (VBScript: Expected ')' in "
            "regular expression)",
            "an unclosed group",
        )
    if code == 5019:
        return (
            "'5019': Application-defined or object-defined error (VBScript: Expected ']' in "
            "regular expression)",
            "an unclosed character class",
        )
    if code == 5018:
        return (
            "'5018': Application-defined or object-defined error (VBScript: Unexpected quantifier)",
            "a quantifier with nothing to repeat",
        )
    if code == 5017:
        return (
            "'5017': Application-defined or object-defined error (VBScript: Syntax error in "
            "regular expression)",
            "a form VBScript does not read, such as a lookbehind",
        )
    return None


# A Collection's methods as its type library declares them.
_COLLECTION_PARAMS: dict[str, tuple[_KnownParam, ...]] = {
    "add": (
        _KnownParam("Item", False, False),
        _KnownParam("Key", True, False),
        _KnownParam("Before", True, False),
        _KnownParam("After", True, False),
    ),
    "item": (_KnownParam("Index", False, False),),
    "remove": (_KnownParam("Index", False, False),),
    "count": (),
}


def _argument_refusal(
    toks: Sequence[VbaToken],
    at: int,
    params: Sequence[_KnownParam],
    member_name: str,
    is_property: bool,
) -> str | None:
    """What a call of a known member with these arguments raises, or None: a named
    argument it has no parameter for (448), more arguments than it takes (450), a
    required one missing (449), and an argument to a Count that takes none (451).
    The call is `o.M(...)`, or `o.M ...` as the statement. `toks[at]` is the
    receiver."""
    open_index = at + 3
    args: list[list[VbaToken]] | None = None
    if _raw(toks, open_index) == "(":
        close = match_paren_from(toks, open_index)
        if close < 0:
            return None
        args = [] if close == open_index + 1 else split_top_level_token_groups(toks, open_index + 1, ",", close)
    elif at == 0 and len(toks) > open_index and toks[open_index].raw_text not in ("=", "."):
        significant = [tok for tok in toks if tok.kind is not TokenKind.COMMENT]
        args = split_top_level_token_groups(significant, open_index, ",", len(significant))
    elif at == 0 and len(toks) == open_index:
        args = []
    elif at > 0 and _raw(toks, open_index) not in ("=", ".", "!"):
        # Read with no parentheses, `x = o.Idx`: no arguments (XLIDE issue #685).
        args = []
    if args is None:
        return None
    if is_property:
        return (
            f"its {member_name} takes no argument. This will raise Run-time error '451': "
            "Property let procedure not defined and property get procedure did not return an object"
            if len(args) > 0
            else None
        )
    named = [arg for arg in args if len(arg) >= 3 and arg[1].raw_text == ":="]
    for arg in named:
        if not any(param.name.lower() == arg[0].raw_text.lower() for param in params):
            return (
                f"its {member_name} has no parameter named '{arg[0].raw_text}'. This will raise "
                "Run-time error '448': Named argument not found"
            )
    positional = len(args) - len(named)
    if len(named) == 0 and not any(param.param_array for param in params) and positional > len(params):
        return (
            f"its {member_name} takes at most {len(params)} argument(s), and {positional} are "
            "passed. This will raise Run-time error '450': Wrong number of arguments or invalid "
            "property assignment"
        )
    given = {arg[0].raw_text.lower() for arg in named}
    missing = next(
        (
            param
            for k, param in enumerate(params)
            if not param.optional
            and not param.param_array
            and param.name.lower() not in given
            and (
                k >= positional
                or (
                    k < len(args)
                    and len([tok for tok in args[k] if tok.kind is not TokenKind.COMMENT]) == 0
                )
            )
        ),
        None,
    )
    if missing is None:
        return None
    return (
        f"its {member_name} needs '{missing.name}', which is not passed. This will raise "
        "Run-time error '449': Argument not optional"
    )


_OPTIONAL_HEAD_RE = re.compile(r"^optional\b", re.IGNORECASE | re.ASCII)
_PARAM_MODIFIER_RE = re.compile(r"^(optional|byval|byref|paramarray)\Z", re.IGNORECASE | re.ASCII)
_PARAM_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*")
_PARAM_ARRAY_RE = re.compile(r"\bparamarray\b", re.IGNORECASE | re.ASCII)
_JS_SPACES_RE = re.compile(f"{_JS_SPACE}+")


def _signature_params(signature: str) -> list[_KnownParam] | None:
    """The parameters a member signature lists: `M(ByVal a As Long, [ByVal b As Long])`."""
    from ...types.type_inference import runtime_signature_parameter_text, split_signature_top_level
    inner = runtime_signature_parameter_text(signature)
    if inner is None:
        return None
    listed = js_trim(inner)
    params: list[_KnownParam] = []
    for raw in [] if listed == "" else split_signature_top_level(listed):
        text = js_trim(raw)
        optional = text.startswith("[") or _OPTIONAL_HEAD_RE.match(text) is not None
        words = [
            word
            for word in _JS_SPACES_RE.split(text.replace("[", "").replace("]", ""))
            if _PARAM_MODIFIER_RE.match(word) is None
        ]
        name_match = _PARAM_NAME_RE.match(words[0] if words else "")
        if name_match is None:
            return None
        params.append(_KnownParam(name_match.group(0), optional, _PARAM_ARRAY_RE.search(text) is not None))
    return params


# The Scripting.FileSystemObject's members, as its type library lists them.
_FSO_CLASS = _KnownClass(
    display="FileSystemObject",
    members=frozenset({
        "drives", "buildpath", "copyfile", "copyfolder", "createfolder", "createtextfile",
        "deletefile", "deletefolder", "driveexists",
        "fileexists", "folderexists", "getabsolutepathname", "getbasename", "getdrive",
        "getdrivename", "getextensionname", "getfile",
        "getfilename", "getfileversion", "getfolder", "getparentfoldername", "getspecialfolder",
        "getstandardstream", "gettempname",
        "movefile", "movefolder", "opentextfile",
    }),
)


def _prog_id_class(prog_id: str) -> _KnownClass | None:
    """The class a CreateObject of a ProgID literal gives, where its members are known."""
    lower = js_trim(prog_id).lower()
    if lower == "vbscript.regexp":
        return _REGEXP_CLASS
    if lower == "scripting.filesystemobject":
        return _FSO_CLASS
    return None


_BOOLEAN_TEXT_RE = re.compile(f"^{_JS_SPACE}*(true|false){_JS_SPACE}*\\Z", re.IGNORECASE | re.ASCII)
_DIGIT_RE = re.compile(r"[0-9]")


def _check_prog_id_objects(
    base: int, toks: Sequence[VbaToken], held: Mapping[str, _KnownClass], push: PushFn
) -> None:
    """Faults a literal shows on these objects (XLIDE issue #477, measured in Excel
    16.0): CreateObject and GetObject of a ProgID no class has (429), a RegExp
    pattern VBScript refuses at Test, Execute or Replace, Null given to them (13),
    Global, IgnoreCase or MultiLine set to text that is no Boolean (13), and an
    IOMode OpenTextFile does not take (5)."""

    def at(tok: VbaToken) -> Span:
        return Span(base + tok.start, base + tok.end)

    i = 0
    while i + 2 < len(toks):
        word = token_text(toks[i])
        if word in ("createobject", "getobject") and toks[i + 1].raw_text == "(" and _raw(toks, i - 1) != ".":
            close = match_paren_from(toks, i + 1)
            args = split_top_level_token_groups(toks, i + 2, ",", close) if close > i + 1 else []
            arg_index = 0 if word == "createobject" else 1
            arg = args[arg_index] if arg_index < len(args) else None
            literal = arg[0] if arg is not None and len(arg) == 1 and arg[0].kind is TokenKind.STRING_LITERAL else None
            problem = _prog_id_problem(string_literal_value(literal.raw_text)) if literal is not None else None
            if literal is not None and problem:
                shown = "CreateObject" if word == "createobject" else "GetObject"
                push(
                    "runtimeArgumentValue",
                    f"{shown}: {problem}. This will raise Run-time error '429': ActiveX component "
                    "can't create object.",
                    at(literal),
                )
            i += 1
            continue
        name = token_name(toks[i])
        known = held.get(name.lower() if name else "")
        if known is None or _raw(toks, i - 1) == "." or toks[i + 1].raw_text != ".":
            i += 1
            continue
        member = token_text(toks[i + 2])
        if known.display == "RegExp":
            if member in ("test", "execute", "replace") and _raw(toks, i + 3) == "(":
                pattern_problem = _reg_exp_pattern_problem(known.pattern) if known.pattern is not None else None
                close = match_paren_from(toks, i + 3)
                first = split_top_level_token_groups(toks, i + 4, ",", close)[0] if close > i + 4 else None
                if pattern_problem is not None:
                    error, text = pattern_problem
                    push(
                        "runtimeArgumentValue",
                        f'The pattern "{known.pattern}" has {text}. This will raise Run-time error {error}.',
                        at(toks[i + 2]),
                    )
                elif first is not None and len(first) == 1 and token_text(first[0]) == "null":
                    push(
                        "runtimeArgumentValue",
                        f"RegExp.{toks[i + 2].raw_text} takes a String, and Null is none. This will "
                        "raise Run-time error '13': Type mismatch.",
                        at(first[0]),
                    )
            elif (
                member in ("global", "ignorecase", "multiline")
                and i == 0
                and _raw(toks, 3) == "="
                and _kind(toks, 4) is TokenKind.STRING_LITERAL
                and len(toks) == 5
            ):
                text = string_literal_value(toks[4].raw_text)
                if _BOOLEAN_TEXT_RE.search(text) is None and _DIGIT_RE.search(text) is None:
                    push(
                        "runtimeArgumentValue",
                        f'RegExp.{toks[2].raw_text} takes True or False, and "{text}" is neither. '
                        "This will raise Run-time error '13': Type mismatch.",
                        at(toks[4]),
                    )
        i += 1


def check_runtime_member_not_found(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    member_ctx: MemberCompletionContext,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    project_callables: Iterable[str] = (),
) -> None:
    """`Application.Zzq` in Excel, and `o.Foo` on an Object or Variant local that
    holds a New Collection or an exhaustive project class, raise 438 at run time."""
    model = member_ctx.model
    callables = {node.name.lower() for node in mod.members if isinstance(node, ProcedureNode)} | {name.lower() for name in project_callables}
    application_surface = _excel_application_surface(model)
    range_surface = _excel_range_surface(model) if application_surface is not None else None
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        env = type_environment_for(symbols, member)

        def each_statement(
            stmt: LeafStatementNode,
            env: Mapping[str, str] = env,
        ) -> None:
            for span in statement_and_branch_spans(stmt):
                toks = statement_tokens(source, span)
                _check_open_type_members(
                    source, span.start, toks, env, application_surface, range_surface, member_ctx, push
                )

        for_each_statement(member.body, each_statement, activity)
        _check_collection_items(source, member, symbols, env, member_ctx, activity, push)
        if re.search(r"\bon\s+error\b", source[member.span.start:member.span.end], re.IGNORECASE):
            continue
        proc_sym = procedure_symbol_for(symbols, member)
        auto_instanced = {
            child.name.lower()
            for child in (proc_sym.children if proc_sym is not None else None) or []
            if child.is_auto_instantiated
        }
        static_proc = re.match(r"\s*(?:(?:Public|Private|Friend)\s+)?Static\b", source[member.span.start:member.span.end], re.IGNORECASE) is not None
        locals_ = {child.name.lower() for child in (proc_sym.children if proc_sym else None) or [] if child.kind is VbaSymbolKind.LOCAL_VARIABLE and child.visibility is not SymbolVisibility.STATIC and not static_proc}
        _walk_held_classes(source, member, env, auto_instanced, application_surface, member_ctx, activity, push, locals_, callables)


def _walk_held_classes(
    source: str,
    member: ProcedureNode,
    env: Mapping[str, str],
    auto_instanced: AbstractSet[str],
    application_surface: AbstractSet[str] | None,
    member_ctx: MemberCompletionContext,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    locals_: AbstractSet[str],
    callables: AbstractSet[str],
) -> None:
    """The per-procedure walk that follows what class each late-bound local holds."""

    # Asked only for the target of a Set: walking the whole environment for every
    # procedure was 5% of a large module's pass (XLIDE issue #139).
    def is_late_bound(lower: str) -> bool:
        if lower not in env:
            return False
        normalized = normalize_type(env.get(lower))
        return normalized == "object" or normalized == "variant" or normalized is None

    held: dict[str, _KnownClass] = {}

    # Blocks are entered with the state they start with (XLIDE issue #237).
    def visit(node: BodyNode) -> None:
        if not is_leaf_statement(node):
            return  # a Dim inside the body declares, and runs nothing
        toks = statement_tokens_after_leading_label(source, node.span)
        if any(token_text(tok) in callables and not (_raw(toks, i - 1) == "." and token_text(_at(toks, i - 2)) in held) and not (token_text(tok) == member.name.lower() and _raw(toks, i + 1) == "=") for i, tok in enumerate(toks)):
            held.clear()
        if (
            jump_target_label_declaration(source, node.span) is not None
            or token_text(toks[0] if toks else None) == "gosub"
        ):
            held.clear()
        if isinstance(node, StatementNode) and node.single_line_if_branches:
            _forget_mentioned(toks, held)
            return
        _check_prog_id_objects(node.span.start, toks, held, push)
        _check_statement(source, node.span.start, toks, held, application_surface, member_ctx, push)
        # `re.Pattern = "(a"`: the pattern a later Test or Execute reads.
        target_name = token_name(toks[0] if toks else None)
        target = target_name.lower() if target_name else None
        reg_exp = held.get(target) if target else None
        if (
            target
            and reg_exp is not None
            and reg_exp.display == "RegExp"
            and _raw(toks, 1) == "."
            and token_text(_at(toks, 2)) == "pattern"
            and _raw(toks, 3) == "="
        ):
            value = [tok for tok in toks[4:] if tok.kind is not TokenKind.COMMENT]
            held[target] = dataclasses.replace(
                reg_exp,
                pattern=(
                    string_literal_value(value[0].raw_text)
                    if len(value) == 1 and value[0].kind is TokenKind.STRING_LITERAL
                    else None
                ),
            )
            return
        assigned = set_assignment_target(source, node.span)
        if assigned is not None and assigned[0].lower() in locals_ and is_late_bound(assigned[0].lower()):
            lower = assigned[0].lower()
            equals = next((k for k, tok in enumerate(toks) if tok.raw_text == "="), -1)
            value = list(toks[equals + 1 :])
            source1_name = token_name(value[0]) if len(value) == 1 else None
            source1 = source1_name.lower() if source1_name else None
            from_variable = (
                _known_class_named(env.get(source1), member_ctx)
                if source1 is not None and not is_late_bound(source1)
                else None
            )
            created = (
                _prog_id_class(string_literal_value(value[2].raw_text))
                if len(value) == 4
                and token_text(value[0]) == "createobject"
                and value[1].raw_text == "("
                and value[2].kind is TokenKind.STRING_LITERAL
                and value[3].raw_text == ")"
                else None
            )
            known: _KnownClass | None
            if created is not None:
                known = created
            elif len(value) == 2 and token_text(value[0]) == "new":
                known = _known_class_named(token_name(value[1]), member_ctx)
            elif from_variable is not None and source1 is not None:
                known = dataclasses.replace(from_variable, may_be_nothing=source1 not in auto_instanced)
            else:
                known = None
            if known is not None:
                held[lower] = known
            else:
                held.pop(lower, None)
            return
        _forget_other_uses(toks, held)

    def snapshot() -> dict[str, _KnownClass]:
        return dict(held)

    def restore(saved: dict[str, _KnownClass]) -> None:
        held.clear()
        held.update(saved)

    def forget(names: AbstractSet[str]) -> None:
        for lower in names:
            held.pop(lower, None)

    def touches(stmt: LeafStatementNode) -> set[str]:
        return names_in(source, stmt.span) | held.keys()

    walk_entering_blocks(
        source,
        member.body,
        lambda node: activity is not None and activity.is_inactive(node.span),
        visit,
        BlockEnteringState(snapshot=snapshot, restore=restore, forget=forget, touches=touches),
    )


def _check_collection_items(
    source: str,
    proc: ProcedureNode,
    symbols: ModuleSymbols,
    env: Mapping[str, str],
    member_ctx: MemberCompletionContext,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """The items a Collection local holds, by class (XLIDE issue #246, measured in
    Excel 16.0): `c.Add New Flat1` then `c(1).Radius()` asks a Flat1 for a member
    it lacks, 438; and `For Each x In c` with x a Round1 Sets each item into x,
    raising 13 at an item of another class."""
    held_at = held_objects_at(source, proc, symbols, activity)

    def each_statement(stmt: LeafStatementNode) -> None:
        items = held_at(stmt).items
        if len(items) == 0 or (isinstance(stmt, StatementNode) and stmt.single_line_if_branches):
            return
        toks = statement_tokens(source, stmt.span)
        for i in range(len(toks)):
            name = token_name(toks[i])
            lower = name.lower() if name else None
            held = items.get(lower) if lower else None
            if not held or _raw(toks, i - 1) == ".":
                continue
            # `c(1).Member` or `c.Item(1).Member`.
            if _raw(toks, i + 1) == "(":
                open_index = i + 1
            elif _raw(toks, i + 1) == "." and token_text(_at(toks, i + 2)) == "item" and _raw(toks, i + 3) == "(":
                open_index = i + 3
            else:
                open_index = -1
            close = match_paren_from(toks, open_index) if open_index >= 0 else -1
            if close < 0 or _raw(toks, close + 1) != "." or not token_name(_at(toks, close + 2)):
                continue
            index = (
                js_number(toks[open_index + 1].raw_text)
                if close == open_index + 2 and toks[open_index + 1].kind is TokenKind.INTEGER_LITERAL
                else None
            )
            class_name: str | None
            if index is not None and 1 <= index <= len(held):
                class_name = held[int(index) - 1]
            elif all(item.lower() == held[0].lower() for item in held):
                class_name = held[0]
            else:
                class_name = None
            known = _known_class_named(class_name, member_ctx)
            member_name = token_name(toks[close + 2]) or ""
            if known is None or member_name.lower() in known.members:
                continue
            shown = "".join(tok.raw_text for tok in toks[i : close + 1])
            push(
                "runtimeMemberNotFound",
                f"'{shown}' holds a {known.display} here, which has no member '{member_name}'. "
                "This will raise Run-time error '438': Object doesn't support this property or "
                "method.",
                Span(stmt.span.start + toks[i].start, stmt.span.start + toks[close].end),
            )

    for_each_statement(proc.body, each_statement, activity)
    signatures = build_module_type_signatures(symbols)
    source_names = source_name_scope_for(symbols, proc)
    for node in iter_body_nodes(proc.body, lambda n: activity is not None and activity.is_inactive(n.span)):
        if isinstance(node, ForBlockNode) and node.each:
            _check_for_each_items(source, node, held_at, env, signatures, source_names, member_ctx, push)


_LIBRARY_PREFIX_RE = re.compile(r"^\w+\.", re.ASCII)
_VOWEL_HEAD_RE = re.compile(r"^[AEIOU]")


def _check_for_each_items(
    source: str,
    loop: ForBlockNode,
    held_at: Callable[[BodyNode], Any],
    env: Mapping[str, str],
    signatures: Mapping[str, CallableTypeSignature],
    source_names: SourceNameScope,
    member_ctx: MemberCompletionContext,
    push: PushFn,
) -> None:
    source1 = js_trim(loop.source_expression).lower() if loop.source_expression else None
    control = loop.control_variable.lower() if loop.control_variable else None
    held: Sequence[str] | None = held_at(loop).items.get(source1) if source1 else None
    expected = env.get(control) if control else None
    span = loop.source_expression_span
    # `For Each c In Worksheets` with c As Range: each sheet is Set into c, and a
    # sheet is no Range (XLIDE issue #447, measured in Excel 16.0).
    element = (
        _host_element_type(source, span, env, signatures, source_names, member_ctx)
        if expected and span is not None and not held
        else None
    )
    if element and span is not None and expected:
        bare = _LIBRARY_PREFIX_RE.sub("", element, count=1)
        article = "an" if _VOWEL_HEAD_RE.match(bare) else "a"
        label = f"the items of '{js_trim(loop.source_expression or '')}', each {article} {bare}"
        model_ctx = dataclasses.replace(
            member_ctx, model=member_ctx.model if member_ctx.model is not None else get_excel_object_model()
        )
        if object_assignment_incompatibility_reason(
            expected, InferredArgumentType(element, label, span), model_ctx
        ):
            push(
                "assignmentObjectTypeMismatch",
                f"For Each Sets {label}, into '{loop.control_variable}', a {expected}. This will "
                "raise Run-time error '13': Type mismatch.",
                span,
            )
            return
    if not held or not expected or span is None:
        return
    # A body that may leave the loop may stop before any later item: only the
    # first is certainly Set (XLIDE issue #356, measured in Excel 16.0).
    reached = held[:1] if body_may_leave_loop(source, loop.body) else held
    # A number or string Set into an object variable is Object required, 424
    # (issue #447, measured in Excel 16.0); a Variant takes it.
    object_control = normalize_type(expected) != "variant" and not is_known_scalar_type(
        normalize_type(expected) or ""
    )
    position = next(
        (
            k
            for k, name in enumerate(reached)
            if (
                object_control
                if name == HELD_VALUE
                else object_assignment_incompatibility_reason(
                    expected, InferredArgumentType(name, name, span), member_ctx
                )
                is not None
            )
        ),
        -1,
    )
    shown = js_trim(loop.source_expression or "")
    if position >= 0 and held[position] == HELD_VALUE:
        push(
            "assignmentObjectTypeMismatch",
            f"For Each Sets each item of '{shown}' into '{loop.control_variable}', a {expected}, "
            f"and item {position + 1} is a number or string, no object. This will raise Run-time "
            "error '424': Object required.",
            span,
        )
        return
    if position >= 0:
        push(
            "assignmentObjectTypeMismatch",
            f"For Each Sets each item of '{shown}' into '{loop.control_variable}', a {expected}, "
            f"and item {position + 1} is a {held[position]}. This will raise Run-time error '13': "
            "Type mismatch.",
            span,
        )


def _host_element_type(
    source: str,
    span: Span,
    env: Mapping[str, str],
    signatures: Mapping[str, CallableTypeSignature],
    source_names: SourceNameScope,
    member_ctx: MemberCompletionContext,
) -> str | None:
    """The host type For Each hands out over a host collection: its Item's type (a
    Worksheet over Worksheets, a Workbook over Workbooks), and a Range over a
    Range. None where that is Object or Variant, as over Sheets, which holds
    charts too."""
    toks = [
        tok
        for tok in raw_expression_tokens(source[span.start : span.end])
        if tok.kind is not TokenKind.COMMENT
    ]
    # The analyzer's default host is Excel: with no model given, its globals still resolve.
    model = member_ctx.model if member_ctx.model is not None else get_excel_object_model()
    # The model's keys keep their case: Excel.Worksheets, not excel.worksheets.
    inferred = infer_expression_type(
        toks,
        0,
        env,
        signatures,
        source_names,
        source=source,
        member_ctx=dataclasses.replace(member_ctx, model=model),
    )
    collection = inferred.type_ if inferred is not None else None
    if not collection or "." not in collection:
        return None
    if normalize_type(collection) == "excel.range":
        return "Excel.Range"
    item = next((m for m in get_host_members(collection, model) if m["name"] == "Item"), None)
    item_type = item.get("returns") if item is not None else None
    return item_type if item_type and "." in item_type else None


def _known_class_named(name: str | None, member_ctx: MemberCompletionContext) -> _KnownClass | None:
    if not name:
        return None
    if name.lower() == "collection":
        return _KnownClass(display="Collection", members=_COLLECTION_MEMBERS, params=_COLLECTION_PARAMS)
    project_type: VbaProjectClassMembers | None = next(
        (
            candidate
            for candidate in member_ctx.project_class_members or []
            if candidate.kind == "class"
            and candidate.exhaustive is True
            and candidate.name.lower() == name.lower()
        ),
        None,
    )
    if project_type is None:
        return None
    members = project_type.members
    properties = [m for m in members if m.kind == "property" and m.signature is not None]
    params: dict[str, Sequence[_KnownParam]] = {}
    for m in members:
        # A Property Get's parameters too: `o.Idx` with Idx(ByVal i As Long) raises
        # 449 (XLIDE issue #685).
        listed = (
            _signature_params(m.signature)
            if (m.kind == "method" or (m.kind == "property" and not m.let_accessor and not m.set_accessor))
            and m.signature
            else None
        )
        if listed is not None:
            params[m.name.lower()] = listed
    return _KnownClass(
        params=params,
        display=project_type.name,
        members=frozenset(m.name.lower() for m in members),
        read_only=frozenset(
            m.name.lower() for m in properties if not m.let_accessor and not m.set_accessor
        ),
        write_only=frozenset(
            m.name.lower()
            for m in members
            if m.kind == "property" and m.let_accessor and m.signature is None
        ),
        set_only=frozenset(
            m.name.lower()
            for m in members
            if m.kind == "property" and m.set_accessor and not m.let_accessor and m.signature is None
        ),
        no_let=frozenset(
            m.name.lower()
            for m in members
            if m.kind == "property" and m.set_accessor and not m.let_accessor and m.signature is not None
        ),
        subs=frozenset(m.name.lower() for m in members if m.kind == "method" and m.sub),
        scalar_fields={
            m.name.lower(): m.returns
            for m in members
            if m.kind == "property"
            and m.signature is None
            and not m.let_accessor
            and not m.set_accessor
            and m.returns is not None
            and m.returns.lower() in _SCALAR_FIELD_TYPES
        },
    )


def _worksheet_function_names(model: HostObjectModel | None) -> AbstractSet[str]:
    """The members of Excel's WorksheetFunction, lowercased. The list is the type
    library's, the one Application's check already reads as complete: Ifs, Switch,
    VStack and TextSplit are absent from both, and raise 438 through either (XLIDE
    issue #442, measured in Excel 16.0 build 20430)."""
    return {member["name"].lower() for member in get_host_members("Excel.WorksheetFunction", model)}


def _excel_application_surface(model: HostObjectModel | None) -> AbstractSet[str] | None:
    """Excel's Application members plus the worksheet functions it also answers to."""
    if model is not None and model.get("hostName") is not None and model.get("hostName") != "Excel":
        return None
    application = get_host_type("Excel.Application", model)
    if application is None or application.get("exhaustive") is not True:
        return None
    names: set[str] = set()
    for member in get_host_members("Excel.Application", model):
        names.add(member["name"].lower())
    for member in get_host_members("Excel.WorksheetFunction", model):
        names.add(member["name"].lower())
    return names


def _excel_range_surface(model: HostObjectModel | None) -> AbstractSet[str]:
    """A Range's members, when the model knows all of them; empty otherwise."""
    range_type = get_host_type("Excel.Range", model)
    if range_type is None or range_type.get("exhaustive") is not True:
        return frozenset()
    return {member["name"].lower() for member in get_host_members("Excel.Range", model)}


def _check_statement(
    source: str,
    base: int,
    toks: Sequence[VbaToken],
    held: Mapping[str, _KnownClass],
    application_surface: AbstractSet[str] | None,
    member_ctx: MemberCompletionContext,
    push: PushFn,
) -> None:
    for i in range(len(toks) - 2):
        # `WorksheetFunction.Mid`, `Application.WorksheetFunction.Summ`: the VBE
        # compiles any name there too, and one that is no worksheet function raises
        # 438 (XLIDE issue #442, measured in Excel 16.0).
        if (
            application_surface is not None
            and token_text(toks[i]) == "worksheetfunction"
            and toks[i + 1].raw_text == "."
        ):
            function_name = token_name(toks[i + 2])
            functions = _worksheet_function_names(member_ctx.model) if function_name else None
            if (
                function_name
                and functions is not None
                and function_name.lower() not in functions
                and resolve_receiver_type_at(source, base + toks[i + 1].end, member_ctx)
                == "Excel.WorksheetFunction"
            ):
                push(
                    "runtimeMemberNotFound",
                    f"WorksheetFunction has no function '{function_name}'. The VBE compiles the "
                    "name; this will raise Run-time error '438': Object doesn't support this "
                    "property or method.",
                    Span(base + toks[i + 2].start, base + toks[i + 2].end),
                )
            continue
        if toks[i + 1].raw_text != "." or _raw(toks, i - 1) == ".":
            continue
        receiver = token_name(toks[i])
        member_name = token_name(toks[i + 2])
        if not receiver or not member_name:
            continue
        at = Span(base + toks[i + 2].start, base + toks[i + 2].end)
        known = held.get(receiver.lower())
        if known is not None:
            lower = member_name.lower()
            nothing = f", or '91' while '{receiver}' is Nothing" if known.may_be_nothing else ""
            if lower not in known.members:
                push(
                    "runtimeMemberNotFound",
                    f"'{receiver}' holds a {known.display} here, which has no member "
                    f"'{member_name}'. This will raise Run-time error '438': Object doesn't "
                    f"support this property or method{nothing}.",
                    at,
                )
                continue
            # A Sub assigned, `o.M = 5`, raises 450, and one read with arguments,
            # `x = o.M(1)`, 451 (XLIDE issue #414, measured in Excel 16.0).
            statement_head = token_text(toks[0])
            target = _raw(toks, i + 3) == "=" and (
                i == 0 or (i == 1 and statement_head in ("set", "let"))
            )
            if known.subs is not None and lower in known.subs:
                if target:
                    push(
                        "runtimeMemberNotFound",
                        f"'{receiver}' holds a {known.display} here, whose '{member_name}' is a "
                        "Sub, which takes no assignment. This will raise Run-time error '450': "
                        f"Wrong number of arguments or invalid property assignment{nothing}.",
                        at,
                    )
                    continue
                if i > 0 and _raw(toks, i + 3) == "(" and statement_head != "call":
                    push(
                        "runtimeMemberNotFound",
                        f"'{receiver}' holds a {known.display} here, whose '{member_name}' is a "
                        "Sub, which gives no value to read. This will raise Run-time error '451': "
                        "Property let procedure not defined and property get procedure did not "
                        f"return an object{nothing}.",
                        at,
                    )
                    continue
            # `o.S.Add 1` with S a String field (issue #414).
            scalar_type = known.scalar_fields.get(lower) if known.scalar_fields is not None else None
            if scalar_type and _raw(toks, i + 3) == "." and token_name(_at(toks, i + 4)):
                push(
                    "variantValueMisuse",
                    f"'{receiver}' holds a {known.display} here, whose '{member_name}' is a "
                    f"{scalar_type}, which has no members. This will raise Run-time error '424': "
                    f"Object required{nothing}.",
                    at,
                )
                continue
            # `o.O = New Collection` with no Set, and O a Get and a Set (issue #685,
            # measured in Excel 16.0).
            if target and statement_head != "set" and known.no_let is not None and lower in known.no_let:
                push(
                    "runtimeMemberNotFound",
                    f"'{receiver}' holds a {known.display} here, whose '{member_name}' has a "
                    "Property Get and a Property Set and no Property Let, so it takes no value "
                    "without Set. This will raise Run-time error '438': Object doesn't support "
                    f"this property or method{nothing}.",
                    at,
                )
                continue
            if (
                not target
                and known.set_only is not None
                and lower in known.set_only
                and not (i > 0 and _raw(toks, i - 1) == ".")
            ):
                push(
                    "runtimeMemberNotFound",
                    f"'{receiver}' holds a {known.display} here, whose '{member_name}' has a "
                    "Property Set and no Property Get, so it has no value to read. This will "
                    "raise Run-time error '450': Wrong number of arguments or invalid property "
                    f"assignment{nothing}.",
                    at,
                )
                continue
            # The arguments the member refuses (issue #485, measured in Excel 16.0).
            params = known.params.get(lower) if known.params is not None else None
            refusal = (
                _argument_refusal(
                    toks, i, params, member_name, known.display == "Collection" and lower == "count"
                )
                if params is not None
                else None
            )
            if refusal:
                push(
                    "runtimeMemberNotFound",
                    f"'{receiver}' holds a {known.display} here: {refusal}{nothing}.",
                    at,
                )
                continue
            # `o.RO = 5` as the statement, a Let into a Get-only property.
            assigned = i == 0 and _raw(toks, i + 3) == "="
            if assigned and known.read_only is not None and lower in known.read_only:
                push(
                    "runtimeMemberNotFound",
                    f"'{receiver}' holds a {known.display} here, whose '{member_name}' has a "
                    "Property Get and no Property Let. This will raise Run-time error '451': "
                    "Property let procedure not defined and property get procedure did not "
                    f"return an object{nothing}.",
                    at,
                )
            elif not assigned and known.write_only is not None and lower in known.write_only:
                push(
                    "runtimeMemberNotFound",
                    f"'{receiver}' holds a {known.display} here, whose '{member_name}' has a "
                    "Property Let and no Property Get, so it has no value to read. This will "
                    "raise Run-time error '450': Wrong number of arguments or invalid property "
                    f"assignment{nothing}.",
                    at,
                )
            continue
        if (
            application_surface is not None
            and receiver.lower() == "application"
            and member_name.lower() not in application_surface
        ):
            receiver_type = resolve_receiver_type_at(source, base + toks[i + 1].end, member_ctx)
            if receiver_type == "Excel.Application":
                push(
                    "runtimeMemberNotFound",
                    f"Application has no member '{member_name}', and it is not a worksheet "
                    "function either. The VBE compiles the name because Application is "
                    "extensible; this will raise Run-time error '438': Object doesn't "
                    "support this property or method.",
                    at,
                )


# The scalar types a field may be declared as, whose value has no members.
_SCALAR_FIELD_TYPES: frozenset[str] = frozenset(
    {"string", "long", "integer", "double", "single", "boolean", "date", "currency", "byte", "longlong"}
)

# Collection's members, its hidden enumerator included.
_COLLECTION_SURFACE: frozenset[str] = _COLLECTION_MEMBERS | {"_newenum"}

_MEMBER_NOT_SUPPORTED = "This will raise Run-time error '438': Object doesn't support this property or method."


def _check_open_type_members(
    source: str,
    base: int,
    toks: Sequence[VbaToken],
    env: Mapping[str, str],
    application_surface: AbstractSet[str] | None,
    range_names: AbstractSet[str] | None,
    member_ctx: MemberCompletionContext,
    push: PushFn,
) -> None:
    """A member no early-bound receiver of an open type has (XLIDE issue #305, each
    measured in Excel 16.0): the VBE compiles the name, since the interface is
    extensible, and the call raises 438. A local As Collection, a Range by any
    route (`Cells.Nope`, `Range("A1").Nope`, a Range variable), a variable As
    Application, and ActiveSheet when neither a Worksheet nor a Chart nor any
    document module of the project has the name."""
    model = member_ctx.model
    project_types = member_ctx.project_class_members or []
    sheet_names: AbstractSet[str] | None = None
    for i in range(1, len(toks) - 1):
        name = token_name(toks[i + 1]) if toks[i].raw_text == "." else None
        if (
            not name
            or (token_name(toks[i - 1]) is None and toks[i - 1].raw_text != ")")
            or toks[i - 1].kind is TokenKind.KEYWORD
        ):
            continue
        lower = name.lower()
        at = Span(base + toks[i + 1].start, base + toks[i + 1].end)
        # A plain name before the dot: `c.Nope`, `a.Nope`, `ActiveSheet.Nope`.
        receiver = (
            token_name(toks[i - 1])
            if toks[i - 1].raw_text != ")" and _raw(toks, i - 2) != "."
            else None
        )
        declared = normalize_type(env.get(receiver.lower())) if receiver else None
        if declared in ("collection", "vba.collection"):
            if lower not in _COLLECTION_SURFACE and not any(
                project_type.name.lower() == "collection" for project_type in project_types
            ):
                push(
                    "runtimeMemberNotFound",
                    f"'{receiver}' is a Collection, which has only Add, Count, Item and Remove. "
                    f"The VBE compiles '{name}'; {_MEMBER_NOT_SUPPORTED}",
                    at,
                )
            continue
        if application_surface is None:
            continue
        if declared in ("application", "excel.application") and lower not in application_surface:
            push(
                "runtimeMemberNotFound",
                f"'{receiver}' is an Application, which has no member '{name}', and it is not a "
                "worksheet function either. The VBE compiles the name because Application is "
                f"extensible; {_MEMBER_NOT_SUPPORTED}",
                at,
            )
            continue
        if receiver and token_text(toks[i - 1]) == "activesheet" and "activesheet" not in env:
            if sheet_names is None:
                sheet_names = _sheet_surface(model, project_types)
            if lower not in sheet_names:
                push(
                    "runtimeMemberNotFound",
                    f"ActiveSheet has no member '{name}': neither a Worksheet nor a Chart has "
                    "one, and no document module of the project declares it. "
                    f"{_MEMBER_NOT_SUPPORTED}",
                    at,
                )
            continue
        if (
            range_names
            and lower not in range_names
            and resolve_receiver_type_at(source, base + toks[i].end, member_ctx) == "Excel.Range"
        ):
            push(
                "runtimeMemberNotFound",
                f"A Range has no member '{name}'. The VBE compiles the name because Range is "
                f"extensible; {_MEMBER_NOT_SUPPORTED}",
                at,
            )


def _sheet_surface(
    model: HostObjectModel | None, project_types: Sequence[VbaProjectClassMembers]
) -> AbstractSet[str]:
    """What ActiveSheet may answer to: a Worksheet's members, a Chart's, and every
    document module's."""
    names: set[str] = set()
    for type_name in ("Excel.Worksheet", "Excel.Chart"):
        for member in get_host_members(type_name, model):
            names.add(member["name"].lower())
    for project_type in project_types:
        if project_type.kind == "document":
            for project_member in project_type.members:
                names.add(project_member.name.lower())
    return names


_MSFORMS_RETURNS_RE = re.compile(r"^MSForms\.", re.IGNORECASE | re.ASCII)


def _check_form_control_names(
    source: str,
    base: int,
    toks: Sequence[VbaToken],
    member_ctx: MemberCompletionContext,
    added: AbstractSet[str] | Literal["any"],
    push: PushFn,
) -> None:
    """`f.Controls("Nope")` on a form whose controls are known, with no control of
    that name (case-insensitive, those inside a Frame included), raises
    -2147024809, "Could not find the specified object" (XLIDE issue #226, measured
    in Excel 16.0). `Me.Controls(...)` inside the form does the same."""
    for i in range(1, len(toks) - 3):
        if (
            token_text(toks[i]) != "controls"
            or toks[i - 1].raw_text != "."
            or toks[i + 1].raw_text != "("
            or toks[i + 2].kind not in (TokenKind.STRING_LITERAL, TokenKind.INTEGER_LITERAL)
            or toks[i + 3].raw_text != ")"
        ):
            continue
        form: VbaProjectClassMembers | None = project_type_at(source, base + toks[i - 1].end, member_ctx)
        if form is None or form.kind != "userform" or form.exhaustive is not True or added == "any":
            continue
        controls = [m for m in form.members if _MSFORMS_RETURNS_RE.match(m.returns or "")]
        # `Me.Controls(99)`: Controls counts from 0 (XLIDE issue #315, measured in
        # Excel 16.0). A procedure that adds a control is not judged.
        if toks[i + 2].kind is TokenKind.INTEGER_LITERAL:
            index = js_number(toks[i + 2].raw_text)
            if len(added) == 0 and index >= len(controls):
                push(
                    "runtimeMemberNotFound",
                    f"The form {form.name} has {len(controls)} control"
                    f"{'' if len(controls) == 1 else 's'}, indexed 0 to {len(controls) - 1}; "
                    f"{js_number_to_string(index)} is none of them. This will raise Run-time "
                    "error '-2147024809': Invalid argument.",
                    Span(base + toks[i + 2].start, base + toks[i + 2].end),
                )
            continue
        name = string_literal_value(toks[i + 2].raw_text)
        # `Me.Controls.Add "Forms.TextBox.1", "Dyn"` names one the designer lacks.
        if name.lower() in added:
            continue
        if not any(control.name.lower() == name.lower() for control in controls):
            push(
                "runtimeMemberNotFound",
                f'The form {form.name} has no control named "{name}". This will raise Run-time '
                "error '-2147024809': Could not find the specified object.",
                Span(base + toks[i + 2].start, base + toks[i + 2].end),
            )


def _controls_added_in(
    source: str, body: Sequence[BodyNode], activity: ConditionalActivityTracker | None
) -> AbstractSet[str] | Literal["any"]:
    """The control names a procedure gives `Controls.Add` as a literal second
    argument, lowercased (XLIDE issue #315), or "any" when one is added under a name
    the code does not spell out."""
    names: set[str] = set()
    any_name = False

    def each_statement(stmt: LeafStatementNode) -> None:
        nonlocal any_name
        for span in statement_and_branch_spans(stmt):
            toks = statement_tokens(source, span)
            for i in range(2, len(toks)):
                if (
                    token_text(toks[i]) != "add"
                    or toks[i - 1].raw_text != "."
                    or token_text(toks[i - 2]) != "controls"
                ):
                    continue
                open_index = i + 1 if _raw(toks, i + 1) == "(" else -1
                close = match_paren_from(toks, open_index) if open_index > 0 else len(toks)
                args = split_top_level_token_groups(
                    toks, open_index + 1 if open_index > 0 else i + 1, ",", close
                )
                named = next(
                    (arg for arg in args if _raw(arg, 1) == ":=" and token_text(arg[0]) == "name"),
                    None,
                )
                arg: list[VbaToken] | None
                if named is not None:
                    arg = named[2:]
                elif len(args) > 1 and _raw(args[1], 1) == ":=":
                    arg = None
                else:
                    arg = args[1] if len(args) > 1 else None
                if arg is not None and len(arg) == 1 and arg[0].kind is TokenKind.STRING_LITERAL:
                    names.add(string_literal_value(arg[0].raw_text).lower())
                else:
                    any_name = True

    for_each_statement(body, each_statement, activity)
    return "any" if any_name else names


def _forget_other_uses(toks: Sequence[VbaToken], held: dict[str, _KnownClass]) -> None:
    """A tracked variable named in any position other than `name.Member` is no longer followed."""
    for i in range(len(toks)):
        name = token_name(toks[i])
        lower = name.lower() if name is not None else None
        if lower and lower in held and _raw(toks, i - 1) != "." and _raw(toks, i + 1) != ".":
            del held[lower]


def _forget_mentioned(toks: Sequence[VbaToken], held: dict[str, _KnownClass]) -> None:
    for tok in toks:
        name = token_name(tok)
        lower = name.lower() if name is not None else None
        if lower and lower in held:
            del held[lower]
