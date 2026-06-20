"""Unit tests for the project-type-name resolver (completion/type_completion.py).

Covers the five classification outcomes the diagnostics rules depend on:
primitive, host, project (single match), ambiguous (project-type collision), and
unknown (None — the no-false-positive gate).
"""

from __future__ import annotations

from pyvbaanalysis.completion import resolve_type_name
from pyvbaanalysis.completion.type_completion import (
    project_type_candidates,
    type_completion_candidates,
)
from pyvbaanalysis.symbols.symbol_model import VbaProjectTypeKind, VbaProjectTypeName


def _project_type(name: str, kind: VbaProjectTypeKind, module: str) -> VbaProjectTypeName:
    return VbaProjectTypeName(name=name, kind=kind, module_name=module)


def test_primitive_resolves() -> None:
    resolved = resolve_type_name("Long")
    assert resolved is not None
    assert resolved.kind == "primitive"
    # Case-insensitive.
    lower = resolve_type_name("long")
    assert lower is not None and lower.kind == "primitive"


def test_host_type_resolves() -> None:
    resolved = resolve_type_name("Workbook")
    assert resolved is not None
    assert resolved.kind == "host"


def test_project_class_resolves_single() -> None:
    project_types = [_project_type("Person", VbaProjectTypeKind.CLASS, "Person")]
    resolved = resolve_type_name("Person", project_types)
    assert resolved is not None
    assert resolved.kind == "class"
    assert resolved.module_name == "Person"


def test_project_enum_resolves_single() -> None:
    project_types = [_project_type("Color", VbaProjectTypeKind.ENUM, "Module1")]
    resolved = resolve_type_name("Color", project_types)
    assert resolved is not None
    assert resolved.kind == "enum"


def test_ambiguous_project_type_collision() -> None:
    # Same name owned by two distinct project types -> the ambiguous marker.
    project_types = [
        _project_type("Shape", VbaProjectTypeKind.CLASS, "ModuleA"),
        _project_type("Shape", VbaProjectTypeKind.USER_TYPE, "ModuleB"),
    ]
    resolved = resolve_type_name("Shape", project_types)
    assert resolved is not None
    assert resolved.kind == "ambiguous"


def test_unknown_name_resolves_to_none() -> None:
    assert resolve_type_name("NoSuchTypeXyz") is None


def test_no_project_types_never_ambiguous() -> None:
    # With no project types, project_type_candidates yields nothing, so 'ambiguous'
    # is unreachable for a primitive/host name (the no-false-positive gate).
    assert not any(c.kind == "ambiguous" for c in type_completion_candidates(None))
    assert project_type_candidates([]) == []


def test_project_type_shadows_builtin() -> None:
    # A project type named like a primitive takes precedence (priority order).
    project_types = [_project_type("Long", VbaProjectTypeKind.CLASS, "Mod1")]
    resolved = resolve_type_name("Long", project_types)
    assert resolved is not None
    assert resolved.kind == "class"


def test_qualified_project_type_resolves() -> None:
    project_types = [_project_type("Color", VbaProjectTypeKind.ENUM, "Module1")]
    resolved = resolve_type_name("Module1.Color", project_types)
    assert resolved is not None
    assert resolved.kind == "enum"
    assert resolved.module_name == "Module1"


def test_qualified_unknown_module_resolves_to_none() -> None:
    project_types = [_project_type("Color", VbaProjectTypeKind.ENUM, "Module1")]
    assert resolve_type_name("Other.Color", project_types) is None
