"""M9 S1: VBA runtime constant/object negative-lookup tables (vbaRuntime.ts parity)."""

from __future__ import annotations

from pyvbaanalysis.host import application_member_names
from pyvbaanalysis.runtime import (
    resolve_runtime_constant,
    resolve_runtime_object,
    resolve_runtime_object_type,
)


def test_runtime_constants() -> None:
    assert resolve_runtime_constant("vbObjectError") is not None
    vb_obj_err = resolve_runtime_constant("vbObjectError")
    assert vb_obj_err is not None and vb_obj_err["value"] == -2147221504
    assert resolve_runtime_constant("VBBINARYCOMPARE") is not None  # case-insensitive
    # String constants exist without a literal value (vbCrLf, vbNullString).
    cr = resolve_runtime_constant("vbCrLf")
    assert cr is not None and cr["type"] == "String"
    assert resolve_runtime_constant("notAConstant") is None


def test_runtime_objects() -> None:
    err = resolve_runtime_object("Err")
    assert err is not None and err["type"] == "VBA.ErrObject"
    assert resolve_runtime_object("debug") is not None  # case-insensitive
    assert resolve_runtime_object("Application") is None  # host, not a runtime object
    by_type = resolve_runtime_object_type("VBA.ErrObject")
    assert by_type is not None and by_type["name"] == "Err"


def test_application_member_names() -> None:
    names = application_member_names()
    # Implicit Application members (Range, Cells, Calculate, ...) used as bare calls.
    assert "range" in names
    assert "calculate" in names
    assert "cells" in names
    assert len(names) > 100
