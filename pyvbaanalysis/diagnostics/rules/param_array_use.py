"""Rule: a ParamArray used where the VBE refuses it (XLIDE issue #445).

Ported from xlide_vscode/src/analyzer/diagnostics/rules/paramArrayUse.ts. Each
is a compile error, measured in Excel 16.0 (build 20430, 2026-10-02), inside
`Function F(ParamArray p() As Variant)`:

 - `ReDim p(3)` and `Erase p`: "Invalid ParamArray use".
 - `G(p)`, `G((p))`, `Call S(p)` and `S p` with G's or S's parameter ByRef, a
   Variant or an array: "Invalid ParamArray use". A ByVal parameter and another
   ParamArray take it. So do a named argument, `G(v:=p)`, `Module2.G(p)` and
   `k.G(p)` of a project class by their parameter, and any member of a
   late-bound object, `o.G(p)`, whose parameter the compiler cannot see (issue
   #685).
 - `Set p = Nothing`: "Can't assign to array".
 - `Function F(ParamArray p)`, declared with no parentheses: refused.

`p = Array(1)`, `p(0) = 9`, `UBound(p)` and `For Each` over p compile.
"""

from __future__ import annotations

import re
from collections.abc import Callable, Mapping, Sequence
from typing import Literal

from ...completion.member_access import MemberCompletionContext
from ...js_compat import JS_WHITESPACE
from ...lexer.token_helpers import match_paren_from, split_top_level_token_groups
from ...lexer.token_kinds import TokenKind, VbaToken
from ...parser.nodes import LeafStatementNode, ProcedureNode, Span
from ...symbols.symbol_model import ModuleSymbols, VbaProcedureParam, VbaProjectClassMember
from ...types.type_inference import type_environment_for
from ...types.type_names import normalize_type
from ..call_extraction import CallableTypeSignature
from ..context import PushFn
from ..walker import (
    ProcedureStatementVisitor,
    statement_and_branch_spans,
    statement_tokens_after_leading_label,
    token_name,
    token_text,
)

_NEW_PREFIX_RE = re.compile(f"^new[{JS_WHITESPACE}]+", re.IGNORECASE)


def _takes_by_ref(params: Sequence[VbaProcedureParam], index: int, named: str | None = None) -> bool | None:
    """Whether the parameter at `index`, or named `named`, takes its argument ByRef: None where there is none."""
    param: VbaProcedureParam | None
    if named is not None:
        param = next((p for p in params if p.name.lower() == named.lower()), None)
    else:
        param = params[index] if 0 <= index < len(params) else None
    return (not param.by_val and not param.param_array) if param is not None else None


def _at(toks: Sequence[VbaToken], i: int) -> VbaToken | None:
    return toks[i] if 0 <= i < len(toks) else None


def _raw_at(toks: Sequence[VbaToken], i: int) -> str | None:
    tok = _at(toks, i)
    return tok.raw_text if tok is not None else None


def _lower_name(tok: VbaToken | None) -> str | None:
    name = token_name(tok)
    return name.lower() if name else None


_Receiver = Sequence[VbaProcedureParam] | Literal["late"] | None


def check_param_array_use(
    source: str,
    signatures: Mapping[str, CallableTypeSignature],
    push: PushFn,
    symbols: ModuleSymbols | None = None,
    member_ctx: MemberCompletionContext | None = None,
) -> ProcedureStatementVisitor:
    def factory(member: ProcedureNode) -> Callable[[LeafStatementNode], None] | None:
        param = next((p for p in member.params if p.param_array), None)
        if param is None:
            return None
        lower = param.name.lower()
        if not param.is_array:
            push(
                "invalidParamArrayUse",
                f"ParamArray '{param.name}' is declared with no parentheses; a ParamArray is an array of "
                f"Variant, {param.name}(). This is a VBE compile error.",
                param.name_span if param.name_span is not None else param.span,
            )
        own = {p.name.lower() for p in member.params}

        def at(span: Span, tok: VbaToken) -> Span:
            return Span(span.start + tok.start, span.start + tok.end)

        # What `receiver.callee` takes: 'late' for an Object or Variant, the
        # member's parameters for a project class or another module, else unknown.
        env: Mapping[str, str] = type_environment_for(symbols, member) if symbols is not None else {}
        surfaces = (member_ctx.project_class_members if member_ctx is not None else None) or []

        def params_of(found: VbaProjectClassMember | None) -> Sequence[VbaProcedureParam] | None:
            by_kind = found.procedure_params if found is not None else None
            if not by_kind:
                return None
            for kind in ("function", "sub", "propertyGet"):
                params = by_kind.get(kind)
                if params is not None:
                    return params
            return None

        def receiver_signature(receiver: str, callee: str) -> _Receiver:
            if receiver in env:
                type_ = normalize_type(_NEW_PREFIX_RE.sub("", env[receiver]))
                if type_ is None or type_ in ("object", "variant"):
                    return "late"
                surface = next(
                    (s for s in surfaces if s.kind == "class" and s.name.lower() == type_),
                    None,
                )
            else:
                surface = next(
                    (s for s in surfaces if s.kind == "standardModule" and s.name.lower() == receiver),
                    None,
                )
            found = (
                next((m for m in surface.members if m.name.lower() == callee), None)
                if surface is not None
                else None
            )
            return params_of(found)

        def visitor(stmt: LeafStatementNode) -> None:
            for span in statement_and_branch_spans(stmt):
                toks = statement_tokens_after_leading_label(source, span)
                head = token_text(_at(toks, 0))

                def named(tok: VbaToken | None) -> bool:
                    return _lower_name(tok) == lower

                if head == "redim":
                    target = 2 if token_text(_at(toks, 1)) == "preserve" else 1
                    target_tok = _at(toks, target)
                    if target_tok is not None and named(target_tok):
                        push(
                            "invalidParamArrayUse",
                            f"ParamArray '{target_tok.raw_text}' holds what the call passes, and ReDim cannot "
                            "size it. This is a VBE compile error: Invalid ParamArray use.",
                            at(span, target_tok),
                        )
                    continue
                if head == "erase":
                    for group in split_top_level_token_groups(toks, 1, ",", len(toks)):
                        if len(group) == 1 and named(group[0]):
                            push(
                                "invalidParamArrayUse",
                                f"ParamArray '{group[0].raw_text}' holds what the call passes, and Erase cannot "
                                "clear it. This is a VBE compile error: Invalid ParamArray use.",
                                at(span, group[0]),
                            )
                    continue
                if head == "set" and named(_at(toks, 1)) and _raw_at(toks, 2) == "=":
                    push(
                        "arrayTargetAssignment",
                        f"ParamArray '{toks[1].raw_text}' is an array, which Set cannot assign to. "
                        "This is a VBE compile error: Can't assign to array.",
                        at(span, toks[1]),
                    )
                    continue
                # Passed whole to a ByRef parameter of a procedure the module knows,
                # by position or by name, or to a member of an object (issue #685,
                # measured in Excel 16.0): a late-bound one always, since the
                # compiler cannot tell it ByVal; a project class's or another
                # module's by its parameter.
                for i, tok in enumerate(toks):
                    _check_pass(span, toks, i, tok, signatures, own, named, receiver_signature, at, push)

        return visitor

    return factory


def _check_pass(
    span: Span,
    toks: Sequence[VbaToken],
    i: int,
    tok: VbaToken,
    signatures: Mapping[str, CallableTypeSignature],
    own: set[str],
    named: Callable[[VbaToken | None], bool],
    receiver_signature: Callable[[str, str], _Receiver],
    at: Callable[[Span, VbaToken], Span],
    push: PushFn,
) -> None:
    callee = _lower_name(tok)
    qualified = _raw_at(toks, i - 1) == "."
    receiver = _lower_name(_at(toks, i - 2)) if qualified and _raw_at(toks, i - 3) != "." else None
    signature = signatures.get(callee) if callee and callee not in own and not qualified else None
    via: _Receiver = receiver_signature(receiver, callee) if callee and receiver else None
    if signature is None and via is None:
        return
    parenthesized = _raw_at(toks, i + 1) == "("
    statement_call = (
        (i == 0 or (qualified and i == 2))
        and not parenthesized
        and len(toks) > i + 1
        and toks[i + 1].raw_text != "="
    )
    if not parenthesized and not statement_call:
        return
    close = match_paren_from(toks, i + 1) if parenthesized else len(toks)
    args = [] if close < 0 else split_top_level_token_groups(toks, i + 2 if parenthesized else i + 1, ",", close)
    for k, arg in enumerate(args):
        value = [t for t in arg if t.kind is not TokenKind.COMMENT]
        name_at = token_name(value[0]) if len(value) > 2 and value[1].raw_text == ":=" else None
        if name_at is not None:
            value = value[2:]
        while len(value) > 2 and value[0].raw_text == "(" and match_paren_from(value, 0) == len(value) - 1:
            value = value[1:-1]
        if len(value) != 1 or not named(value[0]):
            continue
        if signature is not None:
            if name_at is not None:
                target = next((p for p in signature.params if p.name.lower() == name_at.lower()), None)
            else:
                target = signature.params[k] if k < len(signature.params) else None
            if target is not None and not target.param_array and target.by_ref is not False:
                push(
                    "invalidParamArrayUse",
                    f"ParamArray '{value[0].raw_text}' is passed whole to '{target.name}' of '{signature.name}', "
                    "which takes it ByRef. This is a VBE compile error: Invalid ParamArray use.",
                    at(span, value[0]),
                )
            continue
        assert via is not None
        by_ref = True if via == "late" else _takes_by_ref(via, k, name_at)
        if by_ref:
            what = (
                f"a member of late-bound '{toks[i - 2].raw_text}', which may take it ByRef"
                if via == "late"
                else f"'{toks[i - 2].raw_text}.{toks[i].raw_text}', which takes it ByRef"
            )
            push(
                "invalidParamArrayUse",
                f"ParamArray '{value[0].raw_text}' is passed whole to {what}. "
                "This is a VBE compile error: Invalid ParamArray use.",
                at(span, value[0]),
            )
