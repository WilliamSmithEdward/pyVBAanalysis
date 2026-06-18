"""M4a: symbol graph + bare-identifier resolution (symbolModel / buildModuleSymbols
/ nameResolution parity)."""

from __future__ import annotations

import pytest

from pyvbaanalysis.evidence import load_oracle_cases
from pyvbaanalysis.parser.nodes import Span
from pyvbaanalysis.symbols import (
    BareIdentifierContext,
    BareIdentifierResolutionInput,
    BareIdentifierResolutionScope,
    ModuleSymbolKind,
    VbaSymbol,
    VbaSymbolKind,
    build_module_symbols,
    procedure_signature_from_symbol,
    resolve_bare_identifier_binding,
)

_MODULE_KIND = {
    "standard": ModuleSymbolKind.STANDARD,
    "class": ModuleSymbolKind.CLASS,
    "document": ModuleSymbolKind.DOCUMENT,
    "userform": ModuleSymbolKind.USERFORM,
}

_MODULES: list[tuple[str, str, str]] = [
    (f"{c.id}::{m.name}", m.module_type, m.source) for c in load_oracle_cases() for m in c.modules
]


@pytest.mark.parametrize(
    ("module_type", "source"),
    [(mt, s) for _, mt, s in _MODULES],
    ids=[i for i, _, _ in _MODULES],
)
def test_build_symbols_never_throws_and_spans_in_bounds(module_type: str, source: str) -> None:
    kind = _MODULE_KIND.get(module_type, ModuleSymbolKind.STANDARD)
    ms = build_module_symbols("M", kind, source)
    n = len(source)
    for symbol in [ms.root, *ms.all]:
        assert 0 <= symbol.name_span.start <= symbol.name_span.end <= n
        assert 0 <= symbol.full_span.start <= symbol.full_span.end <= n


def _build(source: str) -> object:  # ModuleSymbols
    return build_module_symbols("Module1", ModuleSymbolKind.STANDARD, source)


def _resolve(
    source: str,
    name: str,
    context: BareIdentifierContext = BareIdentifierContext.EXPRESSION,
    proc_name: str | None = None,
    offset: int | None = None,
    project: list[VbaSymbol] | None = None,
):  # BareIdentifierResolution
    ms = build_module_symbols("Module1", ModuleSymbolKind.STANDARD, source)
    proc = None
    if proc_name is not None:
        proc = next(c for c in (ms.root.children or []) if c.name == proc_name)
    return resolve_bare_identifier_binding(
        BareIdentifierResolutionInput(
            current_module=ms,
            name=name,
            context=context,
            enclosing_procedure=proc,
            offset=offset,
            project_visible_symbols=project or [],
        )
    )


def test_root_children_kinds_and_visibility() -> None:
    src = (
        "Public Const MAX As Long = 10\n"
        "Private mState As Long\n"
        "Public Sub DoWork(ByVal n As Long)\n"
        "    Dim total As Long\n"
        "End Sub"
    )
    ms = build_module_symbols("Module1", ModuleSymbolKind.STANDARD, src)
    kinds = {c.name: c.kind for c in (ms.root.children or [])}
    assert kinds["MAX"] is VbaSymbolKind.CONSTANT
    assert kinds["mState"] is VbaSymbolKind.MODULE_VARIABLE
    assert kinds["DoWork"] is VbaSymbolKind.SUB
    proc = next(c for c in (ms.root.children or []) if c.name == "DoWork")
    child_kinds = {c.name: c.kind for c in (proc.children or [])}
    assert child_kinds["n"] is VbaSymbolKind.PARAMETER
    assert child_kinds["total"] is VbaSymbolKind.LOCAL_VARIABLE


def test_enum_type_event_declare_symbols() -> None:
    src = (
        "Enum Color\n Red\n Green\nEnd Enum\n"
        "Type TPoint\n X As Long\nEnd Type\n"
        'Declare PtrSafe Sub Beep Lib "user32" ()\n'
        "Public Event Click(ByVal x As Long)"
    )
    ms = build_module_symbols("Module1", ModuleSymbolKind.STANDARD, src)
    by_name = {c.name: c for c in (ms.root.children or [])}
    assert by_name["Color"].kind is VbaSymbolKind.ENUM
    assert [m.name for m in (by_name["Color"].children or [])] == ["Red", "Green"]
    assert by_name["TPoint"].kind is VbaSymbolKind.TYPE
    assert [f.name for f in (by_name["TPoint"].children or [])] == ["X"]
    assert by_name["Beep"].kind is VbaSymbolKind.DECLARE
    assert by_name["Beep"].declare_kind == "Sub"
    assert by_name["Click"].kind is VbaSymbolKind.EVENT


def test_inactive_branch_symbols_are_filtered() -> None:
    src = "#If Win32 Then\nPublic A As Long\n#End If\nPublic B As Long"
    ms = build_module_symbols("Module1", ModuleSymbolKind.STANDARD, src)
    names = {c.name for c in (ms.root.children or [])}
    assert "A" not in names  # inactive branch removed (Win32 = false)
    assert "B" in names


def test_resolve_local_module_unresolved() -> None:
    src = (
        "Private mState As Long\n"
        "Public Sub DoWork(ByVal n As Long)\n"
        "    Dim total As Long\n"
        "End Sub"
    )
    assert _resolve(src, "total", proc_name="DoWork").scope is BareIdentifierResolutionScope.LOCAL
    assert _resolve(src, "n", proc_name="DoWork").scope is BareIdentifierResolutionScope.LOCAL
    assert _resolve(src, "mState", proc_name="DoWork").scope is BareIdentifierResolutionScope.MODULE
    assert _resolve(src, "Missing", proc_name="DoWork").scope is BareIdentifierResolutionScope.UNRESOLVED


def test_property_accessor_family_is_not_ambiguous() -> None:
    src = (
        "Property Get X() As Long\n X = 1\nEnd Property\n"
        "Property Let X(ByVal v As Long)\nEnd Property"
    )
    res = _resolve(src, "X")
    assert res.scope is BareIdentifierResolutionScope.MODULE  # collapsed accessor family
    assert len(res.definitions) == 2


def test_ambiguous_across_modules() -> None:
    span = Span(0, 0)
    a = VbaSymbol(name="Shared", kind=VbaSymbolKind.SUB, name_span=span, full_span=span, module_name="A")
    b = VbaSymbol(name="Shared", kind=VbaSymbolKind.SUB, name_span=span, full_span=span, module_name="B")
    res = _resolve("Public Z As Long", "Shared", project=[a, b])
    assert res.scope is BareIdentifierResolutionScope.AMBIGUOUS
    assert len(res.definitions) == 2


def test_project_scope_resolution() -> None:
    span = Span(0, 0)
    other = VbaSymbol(name="Helper", kind=VbaSymbolKind.SUB, name_span=span, full_span=span, module_name="Other")
    res = _resolve("Public Z As Long", "Helper", project=[other])
    assert res.scope is BareIdentifierResolutionScope.PROJECT


def test_type_name_context_only_matches_types_and_enums() -> None:
    src = "Type TPoint\n X As Long\nEnd Type\nSub Helper\nEnd Sub"
    assert _resolve(src, "TPoint", context=BareIdentifierContext.TYPE_NAME).scope is BareIdentifierResolutionScope.MODULE
    assert _resolve(src, "Helper", context=BareIdentifierContext.TYPE_NAME).scope is BareIdentifierResolutionScope.UNRESOLVED


def test_return_variable_offset_gating() -> None:
    src = "Public Function Calc(x As Long) As Long\n    Calc = x * 2\nEnd Function"
    # Past the header: the function name resolves to its local return variable.
    late = _resolve(src, "Calc", proc_name="Calc", offset=10_000)
    assert late.scope is BareIdentifierResolutionScope.LOCAL
    # At/inside the header: the return variable is gated out, so it falls to the
    # module-level function symbol.
    early = _resolve(src, "Calc", proc_name="Calc", offset=0)
    assert early.scope is BareIdentifierResolutionScope.MODULE


def test_procedure_signature_label() -> None:
    src = "Public Function Calc(ByVal x As Long, Optional y As Long = 0) As Long\nEnd Function"
    ms = build_module_symbols("Module1", ModuleSymbolKind.STANDARD, src)
    calc = next(c for c in (ms.root.children or []) if c.name == "Calc")
    sig = procedure_signature_from_symbol(calc)
    assert sig is not None
    # Non-external procedures omit ByVal/ByRef in the label (include_passing is
    # only enabled for external declares).
    assert sig.signature == "Calc(x As Long, [y As Long = 0]) As Long"


def test_external_declare_signature() -> None:
    src = 'Declare PtrSafe Function Foo Lib "kernel32" (ByVal x As Long) As Long'
    ms = build_module_symbols("Module1", ModuleSymbolKind.STANDARD, src)
    foo = next(c for c in (ms.root.children or []) if c.name == "Foo")
    sig = procedure_signature_from_symbol(foo)
    assert sig is not None
    assert sig.external is True
    assert sig.signature == 'Declare PtrSafe Function Foo Lib "kernel32" (ByVal x As Long) As Long'


def test_module_attributes_attached_to_root() -> None:
    src = 'Attribute VB_Name = "Module1"\nPublic X As Long'
    ms = build_module_symbols("Module1", ModuleSymbolKind.STANDARD, src)
    attr_names = {a.name for a in (ms.root.attributes or [])}
    assert "VB_Name" in attr_names
