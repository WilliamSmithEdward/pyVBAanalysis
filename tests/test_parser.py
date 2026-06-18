"""M2: parser never-throws + span integrity over the corpus, plus AST shape tests."""

from __future__ import annotations

import dataclasses

import pytest

from pyvbaanalysis.evidence import load_oracle_cases
from pyvbaanalysis.parser import parse_module
from pyvbaanalysis.parser.nodes import (
    Argument,
    AssignmentNode,
    AttributeNode,
    BinaryExpr,
    CallNode,
    ConditionalDirectiveKind,
    ConditionalDirectiveNode,
    DeclareNode,
    EnumNode,
    Expr,
    ForBlockNode,
    IdentifierExpr,
    IfBlockNode,
    IfBranchKind,
    IfBranchNode,
    IndexExpr,
    ModuleKind,
    NewExpr,
    Node,
    ProcedureNode,
    ProcKind,
    Span,
    StatementNode,
    TypeNode,
    VariableGroupNode,
)

_SOURCES: list[tuple[str, str]] = [
    (f"{c.id}::{m.name}", m.source) for c in load_oracle_cases() for m in c.modules
]


def _iter_spans(obj: object):
    """Yield every Span carried by an AST node reachable from obj."""
    if isinstance(obj, (Node, Expr, IfBranchNode, Argument)):
        span = getattr(obj, "span", None)
        if isinstance(span, Span):
            yield span
        for f in dataclasses.fields(obj):  # type: ignore[arg-type]
            yield from _iter_spans(getattr(obj, f.name))
    elif isinstance(obj, list):
        for item in obj:
            yield from _iter_spans(item)


@pytest.mark.parametrize("source", [s for _, s in _SOURCES], ids=[i for i, _ in _SOURCES])
def test_parse_never_throws_and_spans_in_bounds(source: str) -> None:
    module = parse_module(source)
    n = len(source)
    assert module.span == Span(0, n)
    for span in _iter_spans(module):
        assert 0 <= span.start <= span.end <= n, f"out-of-bounds span {span} for len {n}"


def test_parse_never_throws_on_malformed_input() -> None:
    # Truncated / scrambled fragments must still return a best-effort tree.
    for src in (
        "Sub",
        "If Then",
        "Function F(",
        "x = = )",
        "End Sub",
        "Type",
        ")(}{",
        "For = To",
        "#If",
        "Property Get",
    ):
        module = parse_module(src)
        assert module.span == Span(0, len(src))


def _only_proc(source: str) -> ProcedureNode:
    module = parse_module(source)
    procs = [m for m in module.members if isinstance(m, ProcedureNode)]
    assert len(procs) == 1
    return procs[0]


def test_assignment_value_form() -> None:
    proc = _only_proc("Sub S\n    x = 1 + 2\nEnd Sub")
    assign = proc.body[0]
    assert isinstance(assign, AssignmentNode)
    assert not assign.is_set and not assign.is_let
    assert isinstance(assign.lhs, IdentifierExpr) and assign.lhs.name == "x"
    assert isinstance(assign.rhs, BinaryExpr) and assign.rhs.operator == "+"


def test_set_assignment_with_new() -> None:
    proc = _only_proc("Sub S\n    Set o = New Collection\nEnd Sub")
    assign = proc.body[0]
    assert isinstance(assign, AssignmentNode)
    assert assign.is_set and not assign.is_let
    assert isinstance(assign.rhs, NewExpr) and assign.rhs.type_name == "Collection"


def test_explicit_call_statement() -> None:
    proc = _only_proc("Sub S\n    Call Foo(1, 2)\nEnd Sub")
    call = proc.body[0]
    assert isinstance(call, CallNode)
    assert call.has_call_keyword
    assert isinstance(call.callee, IdentifierExpr) and call.callee.name == "Foo"
    assert len(call.args) == 2


def test_implicit_call_with_parenless_args() -> None:
    proc = _only_proc("Sub S\n    Foo 1, 2\nEnd Sub")
    call = proc.body[0]
    assert isinstance(call, CallNode)
    assert not call.has_call_keyword
    assert len(call.args) == 2


def test_omitted_argument_slot() -> None:
    proc = _only_proc("Sub S\n    Foo(1, , 3)\nEnd Sub")
    call = proc.body[0]
    assert isinstance(call, CallNode)
    assert len(call.args) == 3
    assert call.args[1].value is None  # the omitted middle slot


def test_named_argument() -> None:
    proc = _only_proc("Sub S\n    Foo(Width:=10)\nEnd Sub")
    call = proc.body[0]
    assert isinstance(call, CallNode)
    assert call.args[0].name == "Width"


def test_mid_statement_stays_raw() -> None:
    # Mid(...) = ... is a dedicated statement form, not a [Let] lhs = rhs assignment.
    proc = _only_proc('Sub S\n    Mid(s, 1, 2) = "x"\nEnd Sub')
    assert isinstance(proc.body[0], StatementNode)


def test_if_block_branches() -> None:
    src = (
        "Sub S\n"
        "    If a Then\n"
        "        x = 1\n"
        "    ElseIf b Then\n"
        "        x = 2\n"
        "    Else\n"
        "        x = 3\n"
        "    End If\n"
        "End Sub"
    )
    block = _only_proc(src).body[0]
    assert isinstance(block, IfBlockNode)
    assert block.closed
    kinds = [br.branch_kind for br in block.branches]
    assert kinds == [IfBranchKind.IF, IfBranchKind.ELSE_IF, IfBranchKind.ELSE]
    assert block.branches[0].condition is not None
    assert block.branches[2].condition is None  # the else arm


def test_single_line_if_is_not_a_block() -> None:
    proc = _only_proc("Sub S\n    If a Then x = 1\nEnd Sub")
    assert not isinstance(proc.body[0], IfBlockNode)


def test_for_each_block() -> None:
    proc = _only_proc("Sub S\n    For Each item In coll\n        x = item\n    Next item\nEnd Sub")
    block = proc.body[0]
    assert isinstance(block, ForBlockNode)
    assert block.each
    assert block.control_variable == "item"
    assert block.source_expression == "coll"
    assert block.next_variable == "item"
    assert block.closed


def test_for_numeric_block() -> None:
    proc = _only_proc("Sub S\n    For i = 1 To 10\n        x = i\n    Next\nEnd Sub")
    block = proc.body[0]
    assert isinstance(block, ForBlockNode)
    assert not block.each
    assert block.control_variable == "i"


def test_property_accessors() -> None:
    src = (
        "Property Get Name() As String\n    Name = mName\nEnd Property\n"
        "Property Let Name(v As String)\n    mName = v\nEnd Property"
    )
    module = parse_module(src)
    procs = [m for m in module.members if isinstance(m, ProcedureNode)]
    assert [p.proc_kind for p in procs] == [ProcKind.PROPERTY_GET, ProcKind.PROPERTY_LET]
    assert procs[0].return_type == "String"


def test_declare_statement() -> None:
    src = 'Declare PtrSafe Function Foo Lib "kernel32" Alias "FooA" (ByVal x As Long) As Long'
    module = parse_module(src)
    decl = module.members[0]
    assert isinstance(decl, DeclareNode)
    assert decl.is_function and decl.ptr_safe
    assert decl.name == "Foo"
    assert decl.lib_name == "kernel32"
    assert decl.alias_name == "FooA"
    assert decl.return_type == "Long"
    assert [p.name for p in decl.params] == ["x"]
    assert decl.params[0].by_val


def test_type_block() -> None:
    src = "Public Type TPoint\n    X As Long\n    Y As Long\nEnd Type"
    module = parse_module(src)
    tnode = module.members[0]
    assert isinstance(tnode, TypeNode)
    assert tnode.name == "TPoint"
    assert tnode.closed
    assert [f.name for f in tnode.fields] == ["X", "Y"]
    assert tnode.visibility == "Public"


def test_enum_block_with_values() -> None:
    src = "Enum Color\n    Red = 1\n    Green\n    Blue = 4\nEnd Enum"
    module = parse_module(src)
    enode = module.members[0]
    assert isinstance(enode, EnumNode)
    assert [m.name for m in enode.members] == ["Red", "Green", "Blue"]
    assert enode.members[0].value_raw == "1"
    assert enode.members[1].value_raw is None


def test_class_module_kind_detection() -> None:
    src = 'Attribute VB_Name = "Class1"\nAttribute VB_Exposed = True\nPublic X As Long'
    module = parse_module(src)
    assert module.module_kind is ModuleKind.CLASS


def test_unclosed_procedure_diagnostic() -> None:
    module = parse_module("Sub S\n    x = 1")
    proc = module.members[0]
    assert isinstance(proc, ProcedureNode)
    assert not proc.closed
    assert any("missing End Sub" in d.message for d in module.diagnostics)


def test_raw_statement_fallback() -> None:
    # GoTo / labels are not structured into Assignment/Call - they stay raw.
    proc = _only_proc("Sub S\n    GoTo Done\nDone:\nEnd Sub")
    assert isinstance(proc.body[0], StatementNode)


def test_conditional_directive_if() -> None:
    module = parse_module("#If VBA7 Then\n#End If")
    directives = [m for m in module.members if isinstance(m, ConditionalDirectiveNode)]
    assert directives[0].directive_kind is ConditionalDirectiveKind.IF
    assert directives[0].condition_raw == "VBA7"
    assert directives[-1].directive_kind is ConditionalDirectiveKind.END_IF


def test_module_variable_with_visibility_only() -> None:
    module = parse_module("Public Foo As Long")
    group = module.members[0]
    assert isinstance(group, VariableGroupNode)
    assert group.modifier == "Public"
    assert group.declarations[0].name == "Foo"
    assert group.declarations[0].as_type == "Long"


def test_index_expression_in_assignment_rhs() -> None:
    proc = _only_proc("Sub S\n    x = arr(i, j)\nEnd Sub")
    assign = proc.body[0]
    assert isinstance(assign, AssignmentNode)
    assert isinstance(assign.rhs, IndexExpr)
    assert len(assign.rhs.args) == 2


def test_parse_cache_returns_same_object() -> None:
    src = "Sub S\n    x = 1\nEnd Sub"
    assert parse_module(src) is parse_module(src)


def test_attribute_value_parsing() -> None:
    module = parse_module('Attribute VB_Name = "Module1"')
    attr = module.members[0]
    assert isinstance(attr, AttributeNode)
    assert attr.name == "VB_Name"
    assert attr.value_raw == '"Module1"'
