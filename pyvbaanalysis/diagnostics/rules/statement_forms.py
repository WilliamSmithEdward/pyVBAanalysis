"""Rule family: statement forms the VBE refuses while compiling (XLIDE issue #125).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/statementForms.ts.

Measured in Excel 16.0 (build 20326, 2026-09-25):

- collection-operand: `x = c + 1` with c As New Collection -> "Argument not
  optional". A Collection's default member Item takes an index, so the bare
  variable has no value for the operator.
- sub-used-as-value: `x = Foo` where Foo is a Sub -> "Expected Function or
  variable".
- rem-after-then: `If x Then Rem note` -> "Syntax error". Rem starts a comment
  only at the start of a statement.
- rem-after-statement (issue #231): `x = 1 Rem note`, `Next Rem note`
  -> "Syntax error"; `Dim m As Long Rem note` at module level -> "Expected:
  end of statement". In a one-line If's Then or Else list it is a comment.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Mapping, Sequence
from typing import Protocol, cast

from ...completion.member_access import MemberCompletionContext, project_class_member_at
from ...conditional import ConditionalActivityTracker
from ...flow.procedure_labels import statement_label_declarations, statement_label_references
from ...host.host_model import resolve_host_alias
from ...lexer.token_helpers import is_decimal_line_number
from ...lexer.token_kinds import TokenKind, VbaToken
from ...lexer.tokenize import tokenize_cached
from ...parser.nodes import LeafStatementNode, ModuleNode, ProcedureNode, Span, StatementNode
from ...symbols.symbol_model import (
    ModuleSymbols,
    VbaProcedureSignature,
    VbaSymbolKind,
)
from ...types.type_inference import (
    argumentless_host_default,
    object_holding_default,
    object_value_needs_index,
    procedure_symbol_for,
    type_environment_for,
)
from ...types.type_names import is_known_scalar_type, normalize_type
from ..callable_signatures import build_module_type_signatures
from ..context import PushFn
from ..walker import (
    active_module_members,
    bare_assignment_target,
    first_executable_token_index,
    for_each_statement,
    statement_and_branch_spans,
    statement_tokens,
    token_name,
    token_text,
)

_SCALAR_OPERATORS = frozenset({"=", "<", ">", "<=", ">=", "<>", "+", "-", "*", "/", "\\", "&", "^"})

# Excel's Sheets and Worksheets, whose default member is typed Object (issue #369).
_SHEETS_TYPES = frozenset({"excel.sheets", "excel.worksheets"})


class _HoldingDefault(Protocol):
    """objectHoldingDefault's `{ name, returns }`."""

    @property
    def name(self) -> str: ...

    @property
    def returns(self) -> str: ...


class _RequiredParam(Protocol):
    @property
    def optional(self) -> bool: ...

    @property
    def param_array(self) -> bool: ...


def _required(params: Sequence[_RequiredParam]) -> bool:
    return any(not p.optional and not p.param_array for p in params)


def _article(type_name: str) -> str:
    """`/^[aeiou]/i.test(typeName) ? 'an' : 'a'`."""
    return "an" if type_name[:1] != "" and type_name[0] in "aeiouAEIOU" else "a"


def check_statement_forms(
    source: str,
    mod: ModuleNode,
    symbols: ModuleSymbols,
    project_procedures: Mapping[str, Sequence[VbaProcedureSignature]] | None,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    member_ctx: MemberCompletionContext | None = None,
) -> None:
    """Report Rem placement, Collection operands, Subs read as values, and bare module names."""
    if member_ctx is None:
        member_ctx = MemberCompletionContext()
    ctx = member_ctx
    _check_rem_placement(source, mod, activity, push)
    module_lower = symbols.module_name.lower()
    root_children = symbols.root.children or []
    # The project's other standard modules, and the names this module declares.
    other_modules = {
        type_.name.lower()
        for type_ in (ctx.project_class_members or [])
        if type_.kind == "standardModule" and type_.name.lower() != module_lower
    }
    own_names = {symbol.name.lower() for symbol in root_children}
    # An Enum of the module, unless a Function, Property Get or Declare of
    # the module shares its name: that one is read, and runs (issue #639,
    # measured in Excel 16.0). A variable or a Sub of the name does not.
    own_values = {
        symbol.name.lower()
        for symbol in root_children
        if symbol.kind
        in (VbaSymbolKind.FUNCTION, VbaSymbolKind.PROPERTY_GET, VbaSymbolKind.DECLARE)
    }
    own_enums = {
        symbol.name.lower()
        for symbol in root_children
        if symbol.kind is VbaSymbolKind.ENUM and symbol.name.lower() not in own_values
    }
    # Subs of this module, and of the project's standard modules, by name;
    # a name that is also a Function or a module-level variable anywhere is
    # not judged.
    subs: set[str] = set()
    not_subs: set[str] = set()
    for symbol in root_children:
        lower = symbol.name.lower()
        # A Declare Sub returns nothing either (issue #254).
        if symbol.kind is VbaSymbolKind.SUB or (
            symbol.kind is VbaSymbolKind.DECLARE and symbol.declare_kind == "Sub"
        ):
            subs.add(lower)
        elif symbol.kind is not VbaSymbolKind.TYPE:
            # A Type of the name gives the Sub no value either (issue #639).
            not_subs.add(lower)
    if project_procedures is not None:
        for key, signatures in project_procedures.items():
            for signature in signatures:
                (subs if signature.kind is VbaSymbolKind.SUB else not_subs).add(key.lower())
    # Functions and Property Gets that need an argument, by name: this
    # module's, and another module's Public one when it is the only one of
    # its name. Read bare, `Main = F`, one is "Argument not optional" (issue
    # #645, measured in Excel 16.0).
    needs_argument: dict[str, str] = {}
    own_signatures = build_module_type_signatures(symbols)
    for lower, own_signature in own_signatures.items():
        if own_signature.valued and _required(own_signature.params):
            needs_argument[lower] = module_lower
    if project_procedures is not None:
        for key, signatures in project_procedures.items():
            if not signatures:
                continue
            only = signatures[0]
            key_lower = key.lower()
            if (
                len(signatures) == 1
                and key_lower not in own_signatures
                and key_lower not in own_names
                and only.kind is VbaSymbolKind.FUNCTION
                and (only.visibility.value if only.visibility is not None else None) != "Private"
                and only.module_name.lower() != module_lower
                and _required(only.params)
            ):
                needs_argument[key_lower] = only.module_name.lower()
    for member in active_module_members(mod, activity):
        if not isinstance(member, ProcedureNode):
            continue
        _check_procedure(
            source,
            member,
            symbols,
            ctx,
            activity,
            push,
            other_modules,
            own_names,
            own_enums,
            subs,
            not_subs,
            needs_argument,
        )


def _check_procedure(
    source: str,
    member: ProcedureNode,
    symbols: ModuleSymbols,
    member_ctx: MemberCompletionContext,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
    other_modules: set[str],
    own_names: set[str],
    own_enums: set[str],
    subs: set[str],
    not_subs: set[str],
    needs_argument: Mapping[str, str],
) -> None:
    env = type_environment_for(symbols, member)
    member_lower = member.name.lower()
    # An object whose default member needs an index: a Collection, or
    # Excel's Hyperlinks, Areas, Borders, Windows, Workbooks and Shapes
    # (issue #221, measured in Excel 16.0).
    needs_index: dict[str, bool] = {}

    def indexed(lower: str) -> bool:
        answer = needs_index.get(lower)
        if answer is None:
            type_ = env.get(lower)
            # A variable As Sheets or Worksheets too, though its default is
            # typed Object: `s = o` and `o & "x"` do not compile (issue #369).
            resolved = resolve_host_alias(type_, member_ctx.model) if type_ is not None else None
            sheets = (
                type_ is not None
                and (resolved.lower() if resolved is not None else "") in _SHEETS_TYPES
            )
            answer = type_ is not None and (
                sheets or bool(object_value_needs_index(type_, member_ctx))
            )
            needs_index[lower] = answer
        return answer

    def typed_value(lower: str) -> bool:
        type_ = normalize_type(member.return_type if lower == member_lower else env.get(lower))
        return type_ is not None and is_known_scalar_type(type_)

    locals_: set[str] = set()
    # Arrays, whose `x(1)` is an element: the procedure's, and the module's
    # that no local hides.
    arrays: set[str] = set()
    proc_sym = procedure_symbol_for(symbols, member)
    for child in (proc_sym.children if proc_sym is not None else None) or []:
        locals_.add(child.name.lower())
        if child.is_array:
            arrays.add(child.name.lower())
    for child in symbols.root.children or []:
        if child.is_array and child.name.lower() not in locals_:
            arrays.add(child.name.lower())

    def visit(stmt: LeafStatementNode) -> None:
        for span in statement_and_branch_spans(stmt):
            toks = statement_tokens(source, span)

            def at(i: int, span: Span = span, toks: list[VbaToken] = toks) -> Span:
                return Span(span.start + toks[i].start, span.start + toks[i].end)

            target = bare_assignment_target(source, span)
            target_name = target[0] if target is not None else None
            # A Set's `=` is the assignment too: `Set c = New Collection` is no
            # operand, and neither is `Set cols(1) = c`, whose target is
            # indexed (issue #140).
            first = first_executable_token_index(toks)
            # `Foo` or `Call Foo` from another module, where a module is
            # named Foo: the name means the module before its Sub (issue
            # #369, measured in Excel 16.0). `Foo.Foo` compiles.
            # A line label is its own namespace: `Foo:`, `GoTo Foo` and
            # `Resume Foo` compile beside a module Foo (issue #403).
            labels = {label.span.start for label in statement_label_declarations(source, span)}
            labels.update(label.span.start for label in statement_label_references(source, span))

            def is_label(i: int, span: Span = span, toks: list[VbaToken] = toks) -> bool:
                return 0 <= i < len(toks) and (span.start + toks[i].start) in labels

            callee = first + 1 if token_text(_at(toks, first)) == "call" else first
            callee_token = _at(toks, callee)
            callee_raw = (
                token_name(callee_token) if target is None and not is_label(callee) else None
            )
            callee_name = callee_raw.lower() if callee_raw is not None else None
            if (
                callee_name
                and callee_token is not None
                and _raw_at(toks, callee + 1) != "."
                and _raw_at(toks, callee + 1) != "="
                and callee_name in other_modules
                and callee_name not in locals_
                and callee_name not in own_names
            ):
                raw = callee_token.raw_text
                push(
                    "malformedStatement",
                    f"'{raw}' names a module of this project before any procedure in it, so it "
                    f"cannot be called bare from another module; write {raw}.{raw}. This is a VBE "
                    "compile error: Expected variable or procedure, not module.",
                    at(callee),
                )
            assigns = target is not None or token_text(_at(toks, first)) == "set"
            eq = _find_index(toks, lambda tok: tok.raw_text == "=") if assigns else -1
            # A one-line If is judged as its condition here; each branch is its
            # own span with its own assignment (issue #140: `If c Is Nothing
            # Then Set c = New Collection`).
            then = (
                _find_index(toks, lambda tok: token_text(tok) == "then")
                if token_text(_at(toks, first)) == "if"
                and isinstance(stmt, StatementNode)
                and stmt.single_line_if_branches is not None
                else -1
            )
            limit = then if then > 0 else len(toks)
            for i in range(limit):
                tok = toks[i]
                name = token_name(tok)
                # `x = c.DoIt()` with DoIt a Sub of c's class (issue #369).
                # After AddressOf it is addressof-misuse's (issue #299).
                if (
                    name
                    and target is not None
                    and i > eq
                    and _raw_at(toks, i - 1) == "."
                    and _raw_at(toks, i + 1) != "."
                    and token_text(_at(toks, i - 3)) != "addressof"
                ):
                    found = project_class_member_at(
                        source, span.start + toks[i - 1].end, name, member_ctx
                    )
                    if found is not None and found.sub:
                        push(
                            "subUsedAsValue",
                            f"'{name}' is a Sub of the class, which returns nothing, so it cannot be "
                            "used as a value. This is a VBE compile error: Expected Function or "
                            "variable.",
                            at(i),
                        )
                        continue
                # `Main = F`, `F + 1`, `CStr(F)` and `Module1.F` with F a Function
                # that needs an argument (issue #645).
                bare_lower = name.lower() if name is not None else None
                home = needs_argument.get(bare_lower) if bare_lower is not None else None
                qualifier_raw = (
                    token_name(_at(toks, i - 2)) if _raw_at(toks, i - 1) == "." else None
                )
                qualifier = qualifier_raw.lower() if qualifier_raw is not None else None
                if (
                    name
                    and bare_lower is not None
                    and home
                    and target is not None
                    and i > eq
                    and (_raw_at(toks, i + 1) or "") not in ("(", ".", "!", ":=")
                    and (
                        _raw_at(toks, i - 1) != "."
                        or (qualifier == home and _raw_at(toks, i - 3) != ".")
                    )
                    and token_text(_at(toks, i - 1)) != "addressof"
                    and not (qualifier and token_text(_at(toks, i - 3)) == "addressof")
                    and bare_lower not in locals_
                    and bare_lower not in env
                    and bare_lower != member_lower
                ):
                    push(
                        "argumentCount",
                        f"'{name}' needs an argument, and is read here with none. This is a VBE "
                        "compile error: Argument not optional.",
                        at(i),
                    )
                    continue
                if (
                    not name
                    or _raw_at(toks, i - 1) == "."
                    or _raw_at(toks, i + 1) == ":="
                    or i == eq - 1
                ):
                    continue
                # `Main = Foo()` reads the module Foo too (issue #369).
                name_lower = name.lower()
                if (
                    i != callee
                    and _raw_at(toks, i + 1) != "."
                    and name_lower in other_modules
                    and name_lower not in locals_
                    and name_lower not in own_names
                    and not is_label(i)
                ):
                    push(
                        "malformedStatement",
                        f"'{name}' names a module of this project before any procedure in it, so it "
                        f"cannot be used bare from another module; write {name}.{name}. This is a "
                        "VBE compile error: Expected variable or procedure, not module.",
                        at(i),
                    )
                    continue
                # `Main = E` reads an Enum type as a value (issue #436, measured
                # in Excel 16.0).
                # So does TypeName(E) (issue #639).
                type_name_argument = (
                    token_text(_at(toks, i - 2)) == "typename"
                    and _raw_at(toks, i - 1) == "("
                    and _raw_at(toks, i + 1) == ")"
                )
                if (
                    (
                        (target is not None and i == eq + 1 and len(toks) == eq + 2)
                        or type_name_argument
                    )
                    and name_lower in own_enums
                    and name_lower not in locals_
                ):
                    push(
                        "malformedStatement",
                        f"'{name}' names an Enum type, which has no value; name one of its members, "
                        f"as in {name}.Member. This is a VBE compile error: Expected variable or "
                        "procedure, not enum type.",
                        at(i),
                    )
                    continue
                # `AddressOf TimerProc` takes the procedure's address, not its value.
                if token_text(_at(toks, i - 1)) == "addressof":
                    continue
                lower = name.lower()
                # `CStr(x)`, `Len(x)`: the whole argument of either (issue #438,
                # measured in Word 16.0).
                value_call = (
                    next((fn for fn in ("cstr", "len") if fn == token_text(_at(toks, i - 2))), None)
                    if _raw_at(toks, i - 1) == "("
                    and _raw_at(toks, i + 1) == ")"
                    and _raw_at(toks, i - 3) != "."
                    else None
                )
                # `x(1)` where no default member on the way takes an argument:
                # a Word Document's Name, a Range's Text (issue #438).
                if (
                    _raw_at(toks, i + 1) == "("
                    and _raw_at(toks, i - 1) != "."
                    and lower in env
                    and lower not in arrays
                ):
                    through = argumentless_host_default(env.get(lower), member_ctx)
                    if through:
                        type_name = env[lower]
                        push(
                            "argumentCount",
                            f"'{name}' is {_article(type_name)} {type_name}, whose default member "
                            f"{through} takes no argument. This is a VBE compile error: Wrong number "
                            "of arguments or invalid property assignment.",
                            at(i),
                        )
                        continue
                if _raw_at(toks, i + 1) != "(" and _raw_at(toks, i + 1) != "." and indexed(lower):
                    type_name = env[lower]
                    # A Collection's is builtin-arguments' (issue #242).
                    if value_call and normalize_type(type_name) != "collection":
                        push(
                            "collectionOperand",
                            f"'{name}' is {_article(type_name)} {type_name}: its default member Item "
                            f"needs an index, so {toks[i - 2].raw_text} has no value to take. This is "
                            "a VBE compile error: Argument not optional.",
                            at(i),
                        )
                        continue
                    previous = None if i - 1 == eq else _at(toks, i - 1)
                    operator = _scalar_operator(_at(toks, i + 1), previous)
                    if operator is not None:
                        push(
                            "collectionOperand",
                            f"'{name}' is {_article(type_name)} {type_name}: its default member Item "
                            f"needs an index, so '{operator.raw_text}' has no value to work on. This "
                            "is a VBE compile error: Argument not optional.",
                            at(i),
                        )
                        continue
                    # `s = c` with s a String: the whole value of a Let into a
                    # typed value (issue #221).
                    if (
                        target_name is not None
                        and i == eq + 1
                        and len(toks) == eq + 2
                        and typed_value(target_name.lower())
                    ):
                        push(
                            "collectionOperand",
                            f"'{name}' is {_article(type_name)} {type_name}: its default member Item "
                            f"needs an index, so it has no value for '{target_name}' to take. This is "
                            "a VBE compile error: Argument not optional.",
                            at(i),
                        )
                        continue
                # `x + 1` and `s = x` on a Word Paragraph, whose default member
                # Range holds an object (issue #462, measured in Word 16.0).
                holding = (
                    cast(
                        "_HoldingDefault | None", object_holding_default(env.get(lower), member_ctx)
                    )
                    if _raw_at(toks, i + 1) != "(" and _raw_at(toks, i + 1) != "." and lower in env
                    else None
                )
                if holding:
                    type_name = env[lower]
                    previous = None if i - 1 == eq else _at(toks, i - 1)
                    operator = _scalar_operator(_at(toks, i + 1), previous)
                    into_typed = (
                        operator is None
                        and target_name is not None
                        and i == eq + 1
                        and len(toks) == eq + 2
                        and typed_value(target_name.lower())
                    )
                    if value_call:
                        error = (
                            "Variable required - can't assign to this expression"
                            if value_call == "len"
                            else "Type mismatch"
                        )
                        push(
                            "collectionOperand",
                            f"'{name}' is {_article(type_name)} {type_name}: its default member "
                            f"{holding.name} holds an object ({holding.returns}), so "
                            f"{toks[i - 2].raw_text} has no value to take. This is a VBE compile "
                            f"error: {error}.",
                            at(i),
                        )
                        continue
                    if operator is not None or into_typed:
                        what = (
                            f"'{operator.raw_text}' has no value to work on"
                            if operator is not None
                            else f"it has no value for '{target_name}' to take"
                        )
                        push(
                            "collectionOperand",
                            f"'{name}' is {_article(type_name)} {type_name}: its default member "
                            f"{holding.name} holds an object ({holding.returns}), so {what}. This is "
                            "a VBE compile error: Type mismatch.",
                            at(i),
                        )
                        continue
                if (
                    target is not None
                    and i > eq
                    and lower in subs
                    and lower not in not_subs
                    and lower not in locals_
                    and lower not in env
                ):
                    push(
                        "subUsedAsValue",
                        f"'{name}' is a Sub, which returns nothing, so it cannot be used as a value. "
                        "This is a VBE compile error: Expected Function or variable.",
                        at(i),
                    )

    for_each_statement(member.body, visit, activity)


def _scalar_operator(following: VbaToken | None, previous: VbaToken | None) -> VbaToken | None:
    """`[toks[i + 1], previous].find(...)`: the first that is a scalar operator or Mod."""
    for candidate in (following, previous):
        if candidate is not None and (
            (candidate.kind is TokenKind.OPERATOR and candidate.raw_text in _SCALAR_OPERATORS)
            or token_text(candidate) == "mod"
        ):
            return candidate
    return None


_REM_COMMENT = re.compile(r"^rem\b", re.IGNORECASE | re.ASCII)


def _check_rem_placement(
    source: str,
    mod: ModuleNode,
    activity: ConditionalActivityTracker | None,
    push: PushFn,
) -> None:
    """Rules: rem-after-then and rem-after-statement (issues #125 and #231,
    measured in Excel 16.0). A Rem comment stands at the start of a statement,
    after a line number or a label, after a block Else, or after a statement
    in a one-line If's Then or Else list; it may swallow that If's Else. Right
    after Then, and after any other statement, it is a compile error: "Syntax
    error" in a procedure, "Expected: end of statement" at module level and on
    a procedure's own line. The lexer makes it a comment wherever it stands,
    so its words are never read as code; this judges where it stands."""
    # Procedure bodies, from the end of the header line to the End line.
    bodies: list[Span] = []
    for member in mod.members:
        if isinstance(member, ProcedureNode):
            line_end = source.find("\n", member.span.start)
            bodies.append(Span(member.span.end if line_end < 0 else line_end, member.span.end))

    def in_body(offset: int) -> bool:
        return any(offset > body.start and offset <= body.end for body in bodies)

    segment: list[VbaToken] = []
    one_line_if = False

    def judge(ended_by_colon: bool) -> None:
        nonlocal segment, one_line_if
        toks = segment
        segment = []
        if len(toks) == 0 or toks[0].kind is TokenKind.DIRECTIVE:
            return
        last = toks[-1]
        rem = (
            last if last.kind is TokenKind.COMMENT and _REM_COMMENT.search(last.raw_text) else None
        )
        head = 1 if is_decimal_line_number(toks[0]) else 0
        opener = token_text(_at(toks, head))
        then = (
            _find_index(toks, lambda tok: token_text(tok) == "then")
            if opener == "if" or opener == "elseif"
            else -1
        )
        # The lexer leaves a Rem right after Then a word, so an If stays a
        # one-line If: `If x Then Rem note`, and `ElseIf x Then Rem note`.
        word = _at(toks, then + 1) if then > head else None
        if (
            word is not None
            and token_text(word) == "rem"
            and not (activity is not None and activity.is_inactive(Span(word.start, word.end)))
        ):
            push(
                "remAfterThen",
                "'Rem' cannot follow 'Then' on one line: a Rem comment starts only at the start "
                "of a statement. This is a VBE compile error: Syntax error.",
                _rem_word(word),
            )
        if not one_line_if and opener == "if":
            if then > head and then < len(toks) - 1:
                one_line_if = True
                return
            # `If x Then:` opens a one-line If too.
            one_line_if = then > head and ended_by_colon
        if rem is None or one_line_if or len(toks) - 1 == head:
            return
        if len(toks) - 1 == head + 1 and token_text(toks[head]) == "else":
            return
        if activity is not None and activity.is_inactive(Span(rem.start, rem.end)):
            return
        error = "Syntax error" if in_body(rem.start) else "Expected: end of statement"
        push(
            "remAfterStatement",
            "'Rem' starts a comment only at the start of a statement, after a line number, a "
            "label or Else, or in a one-line If. Put a colon before it, or use an apostrophe. "
            f"This is a VBE compile error: {error}.",
            _rem_word(rem),
        )

    for token in tokenize_cached(source):
        if token.kind is TokenKind.NEWLINE or token.kind is TokenKind.COLON:
            judge(token.kind is TokenKind.COLON)
            if token.kind is TokenKind.NEWLINE:
                one_line_if = False
            continue
        segment.append(token)
    judge(False)


def _rem_word(token: VbaToken) -> Span:
    """The word Rem itself, not the comment it starts."""
    return Span(token.start, token.start + 3)


def _find_index(toks: Sequence[VbaToken], predicate: Callable[[VbaToken], bool]) -> int:
    """Array.prototype.findIndex: the first token the predicate accepts, or -1."""
    for k, tok in enumerate(toks):
        if predicate(tok):
            return k
    return -1


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    """`toks[i]` as JavaScript reads it: undefined (None) outside the list."""
    return toks[i] if 0 <= i < len(toks) else None


def _raw_at(toks: Sequence[VbaToken], i: int) -> str | None:
    """`toks[i]?.rawText`."""
    tok = _at(toks, i)
    return tok.raw_text if tok is not None else None
