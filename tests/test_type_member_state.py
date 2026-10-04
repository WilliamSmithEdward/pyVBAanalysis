"""typeFields.ts and typeMemberState.ts parity: the fields a module's Types
declare, the With subjects over them, and what a local of a module Type holds in
its members statement by statement (XLIDE issues #248, #253, #366, #417)."""

from __future__ import annotations

from pyvbaanalysis.diagnostics.type_fields import (
    field_chain,
    fixed_string_length,
    is_fixed_array_field,
    module_types,
    type_key,
    type_root_at,
    walk_with_subjects,
    with_subjects_in,
)
from pyvbaanalysis.diagnostics.type_member_state import (
    KnownNumber,
    is_array_bounds,
    is_known_number,
    type_member_states_at,
)
from pyvbaanalysis.diagnostics.walker import raw_expression_tokens
from pyvbaanalysis.parser.nodes import ModuleNode, ProcedureNode, iter_body_nodes
from pyvbaanalysis.parser.parse_module import parse_module
from pyvbaanalysis.symbols.build_module_symbols import BuildModuleSymbolsOptions, build_module_symbols
from pyvbaanalysis.symbols.symbol_model import ModuleSymbolKind, ModuleSymbols

_TYPES = "\n".join(
    [
        "Const N As Long = 4",
        "Type Inner",
        "    dyn() As Long",
        "    o As Collection",
        "    n As Long",
        "    s As String",
        "    v As Variant",
        "    fixed(3) As Long",
        "    byConst(1 To N) As Long",
        "    name As String * 3",
        "End Type",
        "Type Outer",
        "    kid As Inner",
        "    kids(2) As Inner",
        "End Type",
    ]
)


def _setup(body: str) -> tuple[str, ModuleNode, ModuleSymbols, ProcedureNode]:
    source = _TYPES + "\nSub P()\n" + body + "\nEnd Sub\n"
    mod = parse_module(source)
    symbols = build_module_symbols(
        "M", ModuleSymbolKind.STANDARD, source, BuildModuleSymbolsOptions(parsed_module=mod)
    )
    proc = next(member for member in mod.members if isinstance(member, ProcedureNode))
    return source, mod, symbols, proc


def test_type_key_drops_the_module_qualifier() -> None:
    assert type_key(" Module1.Outer ") == "outer"
    assert type_key("") is None
    assert type_key(None) is None


def test_module_types_read_fields_and_bounds() -> None:
    source, mod, _symbols, _proc = _setup("x = 1")
    types = module_types(source, mod, None)
    inner = types["inner"]
    assert inner["dyn"].is_array and not is_fixed_array_field(inner["dyn"])
    fixed = inner["fixed"]
    assert fixed.dims is not None and [(d.lower, d.upper, d.explicit_lower) for d in fixed.dims] == [(0, 3, True)]
    by_const = inner["byconst"]
    assert by_const.dims is not None and [(d.lower, d.upper) for d in by_const.dims] == [(1, 4)]
    assert inner["name"].fixed_length == "3"
    assert types["outer"]["kid"].type == "inner"
    assert module_types(source, mod, None) is types


def test_field_chains_and_fixed_string_lengths() -> None:
    source, mod, symbols, proc = _setup("Dim t As Outer\nx = 1")
    types = module_types(source, mod, None)
    toks = raw_expression_tokens("t.kids(1).dyn")
    root = type_root_at(toks, 0, symbols, proc, types, None)
    assert root is not None
    steps = field_chain(toks, root, types)
    assert [step.display for step in steps] == ["t.kids", "t.kids(1).dyn"]
    assert steps[0].path == "t.kids" and steps[1].path is None
    assert fixed_string_length(raw_expression_tokens("t.kid.name"), symbols, proc, types, {}) == 3


def test_with_subjects() -> None:
    body = "Dim t As Outer\nWith t\n.kid.n = 1\nWith .kid\n.n = 2\nEnd With\nEnd With\nx = 3"
    source, mod, symbols, proc = _setup(body)
    types = module_types(source, mod, None)
    subjects = with_subjects_in(source, proc, None, symbols, types)
    displays = sorted((source[start : source.index("\n", start)], subject.display) for start, subject in subjects.items())
    assert displays == [
        (".kid.n = 1", "t"),
        (".n = 2", "t.kid"),
        ("With .kid", "t"),
    ]


def test_with_subjects_walk_deep_nesting_without_recursion() -> None:
    depth = 1500
    body = "Dim t As Outer\n" + "If a Then\n" * depth + "With t\n.kid.n = 1\nEnd With\n" + "End If\n" * depth
    source, mod, symbols, proc = _setup(body)
    types = module_types(source, mod, None)
    seen: list[str] = []
    walk_with_subjects(
        source,
        proc.body,
        None,
        symbols,
        proc,
        types,
        None,
        lambda stmt, subject: seen.append(subject.display) if subject is not None else None,
    )
    # The `With t` header itself is outside any With; the statement in it is not.
    assert seen == ["t"]


def test_member_states_through_a_procedure() -> None:
    body = "\n".join(
        [
            "Dim t As Inner",
            "x = 1",
            "ReDim t.dyn(2)",
            "x = 2",
            "t.n = 5",
            "x = 3",
            "Erase t.dyn",
            "x = 4",
            "Set t.o = New Collection",
            "x = 5",
            "t.o.Add 1",
            "x = 6",
            "Fill t",
            "x = 7",
        ]
    )
    source, mod, symbols, proc = _setup(body)
    types = module_types(source, mod, None)
    at = type_member_states_at(source, symbols, proc, types, None, 0, lambda type_: type_ == "collection")
    seen = {
        source[node.span.start : node.span.end]: at(node, node.span.start)
        for node in iter_body_nodes(proc.body)
        if source[node.span.start : node.span.end].startswith("x = ")
    }
    first = seen["x = 1"]
    assert first["t.dyn"] == "unallocated"
    assert first["t.o"] == "nothing"
    assert first["t.n"] == KnownNumber(0)
    assert first["t.s"] == "emptyString"
    assert first["t.v"] == "empty"
    assert "t.fixed" not in first
    bounds = seen["x = 2"]["t.dyn"]
    assert is_array_bounds(bounds) and [(d.lower, d.upper) for d in bounds.dims] == [(0, 2)]
    number = seen["x = 3"]["t.n"]
    assert is_known_number(number) and number.number == 5
    assert seen["x = 4"]["t.dyn"] == "unallocated"
    assert seen["x = 5"]["t.o"] == "emptyCollection"
    assert "t.o" not in seen["x = 6"]
    assert seen["x = 7"] == {}


def test_member_states_entering_blocks() -> None:
    body = "\n".join(
        [
            "Dim t As Inner",
            "If a Then",
            "x = 1",
            "t.n = 2",
            "End If",
            "x = 3",
        ]
    )
    source, mod, symbols, proc = _setup(body)
    types = module_types(source, mod, None)
    at = type_member_states_at(source, symbols, proc, types, None, 0)
    seen = {
        source[node.span.start : node.span.end]: at(node, node.span.start)
        for node in iter_body_nodes(proc.body)
        if source[node.span.start : node.span.end].startswith("x = ")
    }
    assert seen["x = 1"]["t.n"] == KnownNumber(0)
    # The block names t, so every member of t is forgotten after it.
    assert seen["x = 3"] == {}
