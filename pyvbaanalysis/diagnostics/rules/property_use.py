"""Rule: a property of a project class used in a way its procedures do not
allow (XLIDE issue #266).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/propertyUse.ts. Measured
in Excel 16.0 (build 20326, 2026-10-01), on a variable declared As the class,
each is the compile error "Invalid use of property":

- reading a property that has a Property Let or Set and no Property Get:
  `Main = c.P`, `Set o = c.P`;
- calling a property as a statement: `c.P` with P a Property Get, and
  `c.P 5` with P a Property Let.

A variable As Object or Variant is late bound: the same read compiles and
raises 450, which runtime-member-not-found reports.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Sequence
from dataclasses import dataclass

from ...completion.member_access import (
    MemberCompletionContext,
    MemberCompletionEntry,
    resolve_exact_member_completion,
)
from ...js_compat import js_trim
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import LeafStatementNode, ProcedureNode, Span, StatementNode
from ...types.type_names import is_known_scalar_type, normalize_type
from ..context import PushFn
from ..walker import (
    ProcedureStatementVisitor,
    match_paren_from,
    statement_and_branch_spans,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)

# Heads whose `=` compares; any other statement's first top-level `=` assigns.
_COMPARING_HEADS = frozenset({"if", "elseif", "do", "loop", "while", "select", "case", "for"})


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    """`toks[i]` as JavaScript reads it: undefined (None) outside the list."""
    return toks[i] if 0 <= i < len(toks) else None


def _raw_at(toks: Sequence[VbaToken], i: int) -> str | None:
    """`toks[i]?.rawText`."""
    tok = _at(toks, i)
    return tok.raw_text if tok is not None else None


def check_invalid_property_use(
    source: str,
    member_ctx: MemberCompletionContext,
    push: PushFn,
) -> ProcedureStatementVisitor:
    classes = {
        type_.name.lower()
        for type_ in (member_ctx.project_class_members or [])
        if type_.kind == "class"
    }
    if len(classes) == 0:

        def no_visitor(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
            return None

        return no_visitor

    def check(span: Span, condition_only: bool) -> None:
        all_ = statement_tokens_after_leading_label(source, span)
        # A single-line If is read up to Then: its branches are statements of
        # their own, checked as such.
        if condition_only:
            then = next((k for k, tok in enumerate(all_) if token_text(tok) == "then"), -1)
            toks = all_[: then + 1]
        else:
            toks = all_
        head = token_text(_at(toks, 0))
        first = 1 if head in ("call", "set", "let") else 0
        assign_at = -1 if head in _COMPARING_HEADS else _top_level_equals(toks)
        # A With member assigned, `.M = 1`, stands at the statement's start.
        with_start = _raw_at(toks, first) == "." and assign_at > first
        i = first if with_start else first + 1
        while i + 1 < len(toks):
            current = i
            i += 1
            _check_member(
                source,
                member_ctx,
                push,
                classes,
                span,
                toks,
                head,
                first,
                assign_at,
                with_start,
                current,
            )

    def visit_member(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        def visit(stmt: LeafStatementNode) -> None:
            for span in statement_and_branch_spans(stmt):
                check(
                    span,
                    span is stmt.span
                    and isinstance(stmt, StatementNode)
                    and stmt.single_line_if_branches is not None,
                )

        return visit

    return visit_member


def _check_member(
    source: str,
    member_ctx: MemberCompletionContext,
    push: PushFn,
    classes: set[str],
    span: Span,
    toks: Sequence[VbaToken],
    head: str,
    first: int,
    assign_at: int,
    with_start: bool,
    i: int,
) -> None:
    """One `.` of the statement, at `i`: the loop body of upstream's check."""
    if (
        toks[i].raw_text != "."
        or token_name(toks[i + 1]) is None
        or (i > first and token_name(_at(toks, i - 1)) is None)
        or (i == first and not with_start)
    ):
        return
    # `If .State("q") > 3`, `Case .Count`, `Debug.Print .Count`: a With
    # member after a keyword, which is no receiver (XLIDE issue #413). Me is.
    if i > first and toks[i - 1].kind is TokenKind.KEYWORD and token_text(toks[i - 1]) != "me":
        return
    name = toks[i + 1]
    at = Span(span.start + name.start, span.start + name.end)
    member = resolve_exact_member_completion(source, name.raw_text, at.end, member_ctx)
    # A receiver As Object or Variant resolves no member, so only an
    # early-bound one gets this far.
    if member is None or member.owner.lower() not in classes:
        return
    close = match_paren_from(toks, i + 2) if _raw_at(toks, i + 2) == "(" else i + 1
    after = _raw_at(toks, close + 1)

    # `a.b.M 1` or `.b.M 1`: names and dots from the statement's start.
    # `a.b.M 1`, `.b(2).M 1`: names, each maybe indexed, and dots from `from` to the member.
    def chain_from(from_: int) -> bool:
        k = from_
        while k < i and token_name(_at(toks, k)) is not None:
            k += 1
            if _raw_at(toks, k) == "(":
                shut = match_paren_from(toks, k)
                if shut < 0:
                    return False
                k = shut + 1
            if k == i:
                return True
            if _raw_at(toks, k) != ".":
                return False
            k += 1
        return False

    head_call = (
        assign_at < 0
        and head != "set"
        and head != "let"
        and (chain_from(first) or (_raw_at(toks, first) == "." and chain_from(first + 1)))
    )
    # `o.M .Left + 1`: in a call statement, a dot after a space starts
    # an argument, a With member, and takes no member of M.
    argument_dot = (
        head_call and after == "." and close >= 0 and toks[close + 1].start > toks[close].end
    )
    head_chain = all(
        (token_name(tok) is not None) if k % 2 == 0 else tok.raw_text == "."
        for k, tok in enumerate(toks[first:i])
    )
    misuse = (
        None
        if close < 0
        else _member_misuse(
            member,
            _MemberUse(
                indexed=close > i + 1,
                after=None if argument_dot else after,
                target=token_text(_at(toks, first)) != "."
                and head_chain
                and assign_at == close + 1,
                set_target=head == "set" and assign_at == close + 1,
                statement_call=head_call
                and (argument_dot or (after != "." and after != "!" and after != "(")),
                set_read=head == "set" and assign_at >= 0 and assign_at < i,
                with_target=i == first and assign_at == close + 1,
            ),
        )
    )
    if misuse is not None:
        rule, message = misuse
        text = "".join(tok.raw_text for tok in toks[max(first, i - 1) : i + 2])
        push(rule, f"'{text}' {message}", at)
        return
    if member.kind != "property" or i == first:
        return
    # `c.P.X` or `c.P(1).X`: a member of what it returns, which the
    # property itself does not decide.
    if close < 0 or after == "." or after == "!" or (close > i + 1 and after == "("):
        return
    # `c.P` or `a.b.P` from the statement's start: names and dots in turn.
    if head_chain and assign_at > i:
        return  # the assignment's target, `c.P = 5` or `c.P(1) = 5`, which the assignment rules judge
    chain_start = i - 1
    while (
        chain_start >= 2
        and toks[chain_start - 1].raw_text == "."
        and token_name(toks[chain_start - 2]) is not None
    ):
        chain_start -= 2
    label = "".join(tok.raw_text for tok in toks[chain_start : i + 2])
    if head_chain and assign_at < 0 and head != "set" and head != "let":
        push(
            "invalidPropertyUse",
            f"'{label}' is a property, and a statement cannot call one. This is a VBE compile "
            "error: Invalid use of property.",
            at,
        )
        return
    if member.signature is None and (member.let_accessor or member.set_accessor):
        setter = "a Property Let" if member.let_accessor else "a Property Set"
        push(
            "invalidPropertyUse",
            f"'{label}' has {setter} and no Property Get, so it has no value to read. This is a "
            "VBE compile error: Invalid use of property.",
            at,
        )


@dataclass(frozen=True, slots=True)
class _MemberUse:
    # Written with an argument list: `c.M(1)`.
    indexed: bool
    # The token after the member and its argument list.
    after: str | None
    # The statement assigns to it: `c.M = 5`, `c.M(1) = 2`.
    target: bool
    # The target of a Set: `Set c.M = x`.
    set_target: bool
    # Read whole by a Set: `Set o = c.M`.
    set_read: bool
    # Called as a statement: `c.M`, `Call c.M(1)`.
    statement_call: bool
    # A With member assigned: `.M = 1`.
    with_target: bool


def _member_misuse(member: MemberCompletionEntry, use: _MemberUse) -> tuple[str, str] | None:
    """What the VBE refuses in a use of a project class's member through a
    variable declared As the class (XLIDE issue #414, each measured in Excel 16.0):
    a Sub used as a value or assigned; a member of a scalar's value, or Is
    Nothing on it; an argument list on a Property Get that takes none; a
    Function or Get whose argument is left out; an element of a String Get
    assigned; a member of a property with no Get; and a
    Collection field beside an operator."""
    writes = use.target or use.with_target
    if member.kind == "method" and member.sub:
        # A plain read, `Main = c.M`, is the value rules' (XLIDE issue #369).
        if writes or use.set_target or use.set_read or use.after == ".":
            return (
                "subUsedAsValue",
                "is a Sub, which gives no value and takes no assignment. This is a VBE compile "
                "error: Expected Function or variable.",
            )
        return None
    from ..member_parameter_counts import member_parameter_counts
    total, required = member_parameter_counts(member.signature)
    type_ = normalize_type(member.returns if member.returns is not None else member.declared_type)
    scalar = type_ is not None and type_ != "variant" and is_known_scalar_type(type_)
    if member.signature is not None and member.known_value in ("scalar", "empty") and type_ in (None, "variant") and use.after == "." and (use.indexed or required == 0):
        return "variantValueMisuse", "returns a Variant that holds no object to take a member of. This will raise Run-time error '424': Object required."
    # `c.M = 9` with M a Function returning Long (XLIDE issue #423).
    if member.kind == "method" and scalar and type_ is not None and use.target and not use.indexed:
        return (
            "assignmentToProcedureName",
            f"is a Function returning {_capitalized(type_)}, and a call cannot be assigned to. "
            "This is a VBE compile error: Function call on left-hand side of assignment must "
            "return Variant or Object.",
        )
    # `c.M = 5` or `.M = 1` calls M and assigns to what it returns: a
    # Variant holding a value raises 424, and a Collection, whose Item needs
    # an index, does not compile (XLIDE issue #414, measured in Excel 16.0).
    # Known to return Empty or a value: one holding an object raises 438
    # instead, and is left alone.
    holds_value = member.known_value == "empty" or member.known_value == "scalar"
    if member.kind == "method" and writes and not use.indexed and required == 0:
        if (type_ is None or type_ == "variant") and holds_value:
            return (
                "variantValueMisuse",
                "is a Function, so the assignment calls it and assigns to the Variant it returns, "
                "which holds no object. This will raise Run-time error '424': Object required.",
            )
        if type_ == "collection":
            return (
                "argumentCount",
                "is a Function returning a Collection, so the assignment reaches its default "
                "member Item, which needs an index. This is a VBE compile error: Argument not "
                "optional.",
            )
    if (
        member.kind == "property"
        and member.signature is None
        and (member.let_accessor or member.set_accessor)
    ):
        # `c.M(1) = 2` with M a Property Set and no Let (XLIDE issue #414).
        if writes and not use.set_target and use.indexed and member.set_accessor and not member.let_accessor:
            return (
                "invalidPropertyUse",
                "has a Property Set and no Property Let, so a value cannot be assigned to it. "
                "This is a VBE compile error: Invalid use of property.",
            )
        if use.after == ".":
            setter = "a Property Let" if member.let_accessor else "a Property Set"
            return (
                "invalidPropertyUse",
                f"has {setter} and no Property Get, so it has no value to take a member of. "
                "This is a VBE compile error: Invalid use of property.",
            )
        # `c.M(1) = 2` with a Let that takes only the value (XLIDE issue #414,
        # measured in Excel 16.0).
        if use.target and use.indexed and member.let_accessor and member.let_param_count == 1:
            return (
                "invalidPropertyUse",
                "has a Property Let that takes no index and no Property Get, so an element of it "
                "cannot be assigned. This is a VBE compile error: Invalid use of property.",
            )
        return None
    # A plain read of a property is argument-count's (XLIDE issue #224).
    if (
        required > 0
        and not use.indexed
        and not writes
        and not use.set_target
        and not use.statement_call
        and (member.kind == "method" or use.set_read or use.after == ".")
    ):
        needs = "an argument" if required == 1 else f"{required} arguments"
        return (
            "argumentCount",
            f"needs {needs}, and is read here with none. This is a VBE compile error: Argument "
            "not optional.",
        )
    if scalar and type_ is not None and not use.indexed and use.after == ".":
        return (
            "scalarMemberAccess",
            f"holds {_article(type_)} {_capitalized(type_)}, which has no members. This is a VBE "
            "compile error: Invalid qualifier.",
        )
    if (
        scalar
        and type_ is not None
        and not use.indexed
        and use.after is not None
        and use.after.lower() == "is"
    ):
        return (
            "isOperatorNonObject",
            f"holds {_article(type_)} {_capitalized(type_)}, which Is cannot compare with an "
            "object. This is a VBE compile error: Type mismatch.",
        )
    get_only = (
        member.kind == "property"
        and member.signature is not None
        and not member.let_accessor
        and not member.set_accessor
        and not member.writable
    )
    if get_only and scalar and use.indexed and total == 0:
        if use.target and type_ == "string":
            return (
                "readonlyMemberAssignment",
                "has a Property Get and no Property Let, so an element of it cannot be assigned. "
                "This is a VBE compile error: Can't assign to read-only property.",
            )
        if not writes:
            return (
                "argumentCount",
                "is a Property Get that takes no argument, so it cannot be given one. This is a "
                "VBE compile error: Wrong number of arguments or invalid property assignment.",
            )
    if (
        member.kind == "property"
        and member.signature is None
        and type_ == "collection"
        and not use.indexed
        and use.after is not None
        and use.after in _SCALAR_OPERATOR_TEXT
    ):
        return (
            "collectionOperand",
            f"is a Collection: its default member Item needs an index, so '{use.after}' has no "
            "value to work on. This is a VBE compile error: Argument not optional.",
        )
    return None


_SCALAR_OPERATOR_TEXT = frozenset({"&", "+", "-", "*", "/", "\\", "^"})

_OPTIONAL_OR_PARAM_ARRAY_RE = re.compile(r"^(Optional|ParamArray)\b", re.IGNORECASE | re.ASCII)


def _parameter_counts(signature: str | None) -> tuple[int, int]:
    """How many parameters a source signature declares, and how many a call must
    pass: (total, required)."""
    open_ = signature.find("(") if signature is not None else -1
    close = signature.find(")", open_) if signature is not None and open_ >= 0 else -1
    inner = (
        js_trim(signature[open_ + 1 : close])
        if signature is not None and open_ >= 0 and close > open_
        else ""
    )
    if not inner:
        return 0, 0
    parts = [js_trim(part) for part in inner.split(",")]
    required = [
        part
        for part in parts
        if not _OPTIONAL_OR_PARAM_ARRAY_RE.search(part) and not part.startswith("[")
    ]
    return len(parts), len(required)


def _article(type_: str) -> str:
    return "an" if type_[:1] != "" and type_[0] in "aeiouAEIOU" else "a"


def _capitalized(type_: str) -> str:
    if type_ == "longlong":
        return "LongLong"
    if type_ == "longptr":
        return "LongPtr"
    return type_[:1].upper() + type_[1:]


def _top_level_equals(toks: Sequence[VbaToken]) -> int:
    depth = 0
    for i, tok in enumerate(toks):
        raw = tok.raw_text
        if raw == "(":
            depth += 1
        elif raw == ")":
            depth -= 1
        elif raw == "=" and depth == 0 and tok.kind is TokenKind.OPERATOR:
            return i
    return -1
