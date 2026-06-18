"""M4b: cross-module ProjectIndex (projectIndex.ts parity)."""

from __future__ import annotations

import pytest

from pyvbaanalysis.evidence import load_oracle_cases
from pyvbaanalysis.symbols import (
    ModuleInput,
    ModuleSymbolKind,
    ProjectIndex,
)
from pyvbaanalysis.symbols.name_resolution import BareIdentifierContext, BareIdentifierResolutionScope

_MODULE_KIND = {
    "standard": ModuleSymbolKind.STANDARD,
    "class": ModuleSymbolKind.CLASS,
    "document": ModuleSymbolKind.DOCUMENT,
    "userform": ModuleSymbolKind.USERFORM,
}


def _index(*modules: tuple[str, ModuleSymbolKind, str]) -> ProjectIndex:
    idx = ProjectIndex()
    for name, kind, source in modules:
        idx.set_module(ModuleInput(name, kind, source))
    return idx


def test_visible_procedure_names_cross_module() -> None:
    idx = _index(
        ("ModA", ModuleSymbolKind.STANDARD, "Public Sub Exported()\nEnd Sub\nPrivate Sub Hidden()\nEnd Sub"),
        ("ModB", ModuleSymbolKind.STANDARD, "Sub Local()\nEnd Sub"),
        ("Cls", ModuleSymbolKind.CLASS, "Public Sub Method()\nEnd Sub"),
    )
    names = idx.visible_procedure_names("ModB")
    assert "exported" in names  # exported standard-module proc
    assert "local" in names  # same module (ModB's own)
    assert "hidden" not in names  # Private in another module
    assert "method" not in names  # class members need qualified binding


def test_exported_integer_constants_with_qualified_keys() -> None:
    idx = _index(
        ("ModA", ModuleSymbolKind.STANDARD, "Public Const MAX As Long = 10\nPrivate Const SECRET As Long = 1"),
        ("ModB", ModuleSymbolKind.STANDARD, "Sub S()\nEnd Sub"),
    )
    consts = idx.visible_external_integer_constant_expressions("ModB")
    assert consts.get("max") == "10"
    assert consts.get("moda.max") == "10"
    assert "secret" not in consts  # Private not exported


def test_duplicate_visible_constant_names_become_unknown() -> None:
    idx = _index(
        ("ModA", ModuleSymbolKind.STANDARD, "Public Const DUP As Long = 1"),
        ("ModC", ModuleSymbolKind.STANDARD, "Public Const DUP As Long = 2"),
        ("ModB", ModuleSymbolKind.STANDARD, "Sub S()\nEnd Sub"),
    )
    consts = idx.visible_external_integer_constant_expressions("ModB")
    assert consts.get("dup") is None  # ambiguous -> unknown
    # ...but the module-qualified keys remain distinct.
    assert consts.get("moda.dup") == "1"
    assert consts.get("modc.dup") == "2"


def test_visible_type_names() -> None:
    idx = _index(
        ("ModA", ModuleSymbolKind.STANDARD, "Public Type TPoint\n X As Long\nEnd Type\nPrivate Enum E\n A\nEnd Enum"),
        ("Person", ModuleSymbolKind.CLASS, "Public Name As String"),
        ("ModB", ModuleSymbolKind.STANDARD, "Sub S()\nEnd Sub"),
    )
    names_b = {(t.name, t.kind.value) for t in idx.visible_type_names("ModB")}
    assert ("Person", "class") in names_b  # class module is a type name
    assert ("TPoint", "userType") in names_b  # exported Type
    assert ("E", "enum") not in names_b  # Private Enum hidden cross-module
    # The owning module sees its own Private Enum.
    names_a = {t.name for t in idx.visible_type_names("ModA")}
    assert "E" in names_a


def test_resolve_type_definitions() -> None:
    idx = _index(("ModA", ModuleSymbolKind.STANDARD, "Public Type TPoint\n X As Long\nEnd Type"))
    defs = idx.resolve_type_definitions("ModA", "tpoint")
    assert len(defs) == 1 and defs[0].name == "TPoint"


def test_project_class_members_hide_private_and_collapse_properties() -> None:
    src = (
        "Public Name As String\n"
        "Private secret As Long\n"
        "Public Property Get Size() As Long\nEnd Property\n"
        "Public Property Let Size(ByVal v As Long)\nEnd Property"
    )
    idx = _index(("Person", ModuleSymbolKind.CLASS, src))
    person = next(c for c in idx.project_class_members() if c.name == "Person")
    by_name = {m.name: m for m in person.members}
    assert "secret" not in by_name  # Private hidden
    assert by_name["Name"].kind == "property"
    # Get + Let collapse to one member; Let makes it writable.
    assert by_name["Size"].writable is True
    assert person.exhaustive is True  # class modules are exhaustive


def test_default_member_attribute() -> None:
    src = (
        "Public Function Item(ByVal i As Long) As Variant\n"
        "Attribute Item.VB_UserMemId = 0\n"
        "End Function"
    )
    idx = _index(("Coll", ModuleSymbolKind.CLASS, src))
    coll = next(c for c in idx.project_class_members() if c.name == "Coll")
    item = next(m for m in coll.members if m.name == "Item")
    assert item.default_member is True


def test_project_standard_module_members_are_qualified_receivers() -> None:
    idx = _index(("Asserts", ModuleSymbolKind.STANDARD, "Public Sub AreEqual(a, b)\nEnd Sub"))
    surfaces = idx.project_standard_module_members("Other")
    asserts = next(s for s in surfaces if s.name == "Asserts")
    assert asserts.kind == "standardModule"
    assert any(m.name == "AreEqual" for m in asserts.members)


def test_project_member_surfaces_includes_user_types() -> None:
    idx = _index(("ModA", ModuleSymbolKind.STANDARD, "Public Type TPoint\n X As Long\n Y As Long\nEnd Type"))
    surfaces = idx.project_member_surfaces("ModA")
    tpoint = next(s for s in surfaces if s.name == "TPoint" and s.kind == "userType")
    assert {m.name for m in tpoint.members} == {"X", "Y"}
    assert all(m.writable for m in tpoint.members)


def test_resolve_bare_identifier_tiers() -> None:
    idx = _index(
        ("ModA", ModuleSymbolKind.STANDARD, "Public Sub Shared()\nEnd Sub"),
        ("ModB", ModuleSymbolKind.STANDARD, "Public Sub Caller(ByVal p As Long)\n Dim loc As Long\nEnd Sub"),
    )
    proc = next(c for c in (idx.get_module("ModB").root.children or []) if c.name == "Caller")  # type: ignore[union-attr]
    off = proc.full_span.start + 1
    assert idx.resolve_bare_identifier("ModB", "loc", off, BareIdentifierContext.EXPRESSION).scope is BareIdentifierResolutionScope.LOCAL
    assert idx.resolve_bare_identifier("ModB", "Caller", off, BareIdentifierContext.EXPRESSION).scope is BareIdentifierResolutionScope.MODULE
    assert idx.resolve_bare_identifier("ModB", "Shared", off, BareIdentifierContext.EXPRESSION).scope is BareIdentifierResolutionScope.PROJECT
    assert idx.resolve_bare_identifier("ModB", "Nope", off, BareIdentifierContext.EXPRESSION).scope is BareIdentifierResolutionScope.UNRESOLVED


def test_resolve_bare_identifier_ambiguous_project() -> None:
    idx = _index(
        ("ModA", ModuleSymbolKind.STANDARD, "Public Sub Dup()\nEnd Sub"),
        ("ModC", ModuleSymbolKind.STANDARD, "Public Sub Dup()\nEnd Sub"),
        ("ModB", ModuleSymbolKind.STANDARD, "Sub S()\nEnd Sub"),
    )
    res = idx.resolve_bare_identifier("ModB", "Dup", 0, BareIdentifierContext.EXPRESSION)
    assert res.scope is BareIdentifierResolutionScope.AMBIGUOUS
    assert len(res.definitions) == 2


def test_resolve_unindexed_module() -> None:
    idx = _index(("ModA", ModuleSymbolKind.STANDARD, "Sub S()\nEnd Sub"))
    res = idx.resolve_bare_identifier("Missing", "x", 0, BareIdentifierContext.EXPRESSION)
    assert res.scope is BareIdentifierResolutionScope.UNRESOLVED


def test_resolve_qualified_definition() -> None:
    idx = _index(("ModA", ModuleSymbolKind.STANDARD, "Public Sub Work()\nEnd Sub\nPrivate Sub Hide()\nEnd Sub"))
    assert len(idx.resolve_qualified_definition("ModA", "Work")) == 1
    assert idx.resolve_qualified_definition("ModA", "Hide") == []  # Private not exported


def test_duplicate_procedures() -> None:
    idx = _index(("ModA", ModuleSymbolKind.STANDARD, "Sub Foo()\nEnd Sub\nSub Foo()\nEnd Sub\nSub Bar()\nEnd Sub"))
    dupes = idx.duplicate_procedures("ModA")
    assert {d.name for d in dupes} == {"Foo"}
    assert len(dupes) == 2


def test_procedure_signatures_by_key() -> None:
    idx = _index(("ModA", ModuleSymbolKind.STANDARD, "Public Function Add(a As Long, b As Long) As Long\nEnd Function"))
    sigs = idx.procedure_signatures()
    assert "add" in sigs
    assert "moda.add" in sigs
    assert sigs["add"][0].signature == "Add(a As Long, b As Long) As Long"


def test_reference_scope_local_and_project() -> None:
    idx = _index(
        ("ModA", ModuleSymbolKind.STANDARD, "Public Sub Shared()\nEnd Sub"),
        ("ModB", ModuleSymbolKind.STANDARD, "Sub Caller()\n Dim x As Long\nEnd Sub"),
    )
    proc = next(c for c in (idx.get_module("ModB").root.children or []) if c.name == "Caller")  # type: ignore[union-attr]
    off = proc.full_span.start + 1
    local = idx.reference_scope("ModB", "x", off)
    assert local.kind == "local"
    assert local.search_modules == ["ModB"]
    project = idx.reference_scope("ModB", "Shared", off)
    assert project.kind == "project"
    assert set(project.search_modules) == {"ModA", "ModB"}


def test_remove_module_and_cache_invalidation() -> None:
    idx = _index(
        ("ModA", ModuleSymbolKind.STANDARD, "Public Sub Exported()\nEnd Sub"),
        ("ModB", ModuleSymbolKind.STANDARD, "Sub S()\nEnd Sub"),
    )
    assert "exported" in idx.visible_procedure_names("ModB")
    idx.remove_module("ModA")
    assert "exported" not in idx.visible_procedure_names("ModB")
    assert "ModA" not in idx.module_names()


def test_workspace_and_document_symbols() -> None:
    idx = _index(("ModA", ModuleSymbolKind.STANDARD, "Public Sub Alpha()\nEnd Sub\nPublic Sub Beta()\nEnd Sub"))
    assert idx.document_symbols("ModA") is not None
    assert idx.document_symbols("Missing") is None
    all_names = {s.name for s in idx.workspace_symbols()}
    assert {"Alpha", "Beta"} <= all_names
    assert {s.name for s in idx.workspace_symbols("alph")} == {"Alpha"}


_CASES = [(c.id, c) for c in load_oracle_cases()]


@pytest.mark.parametrize("case", [c for _, c in _CASES], ids=[i for i, _ in _CASES])
def test_project_index_never_throws_over_corpus(case: object) -> None:
    idx = ProjectIndex()
    modules = case.modules  # type: ignore[attr-defined]
    for m in modules:
        idx.set_module(ModuleInput(m.name, _MODULE_KIND.get(m.module_type, ModuleSymbolKind.STANDARD), m.source))
    first = modules[0].name
    # Exercise the diagnostic-facing query surfaces; none should raise.
    idx.visible_procedure_names(first)
    idx.visible_procedure_signatures(first)
    idx.visible_identifier_names(first)
    idx.visible_external_integer_constant_expressions(first)
    idx.visible_type_names(first)
    idx.procedure_signatures()
    idx.project_member_surfaces(first)
    for m in modules:
        idx.duplicate_procedures(m.name)
