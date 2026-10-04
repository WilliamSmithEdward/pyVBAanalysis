"""M9 S1: VBA runtime constant/object negative-lookup tables (vbaRuntime.ts parity)."""

from __future__ import annotations

from pyvbaanalysis.host import application_member_names
from pyvbaanalysis.runtime import (
    resolve_runtime_constant,
    resolve_runtime_function,
    resolve_runtime_object,
    resolve_runtime_object_type,
)
from pyvbaanalysis.runtime.vba_runtime import VBA_RUNTIME_CONSTANTS, VBA_RUNTIME_OBJECTS


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
    # XLIDE 2f49b93 (issue #315): UserForms is a runtime object.
    user_forms = resolve_runtime_object("UserForms")
    assert user_forms is not None and user_forms["type"] == "VBA.UserForms"
    assert user_forms in VBA_RUNTIME_OBJECTS
    assert resolve_runtime_constant("vbCrLf") in VBA_RUNTIME_CONSTANTS


def test_runtime_function_params_2f49b93() -> None:
    # XLIDE 2f49b93: IRR's ValueArray is an array of Double (issue #218), and
    # Sqr raises 94 on Null while Oct passes it through (issue #332).
    irr = resolve_runtime_function("IRR")
    assert irr is not None and irr.params is not None
    assert (irr.params[0].name, irr.params[0].type_, irr.params[0].is_array) == (
        "ValueArray",
        "Double",
        True,
    )
    assert irr.params[1].is_array is False and irr.params[1].optional is True
    sqr = resolve_runtime_function("Sqr")
    assert sqr is not None and sqr.params is not None and sqr.params[0].null_raises
    oct_ = resolve_runtime_function("Oct")
    assert oct_ is not None and not any(p.null_raises for p in oct_.params or ())


def test_vba_library_names() -> None:
    from pyvbaanalysis.runtime.vba_library_names import (
        VBA_ERR_READ_ONLY,
        VBA_LIBRARY_CONTAINERS,
        VBA_LIBRARY_NAMES,
    )

    assert "left$" in VBA_LIBRARY_NAMES and "userforms" in VBA_LIBRARY_NAMES
    assert VBA_LIBRARY_CONTAINERS["vbtristate"] == frozenset({"vbfalse", "vbtrue", "vbusedefault"})
    assert VBA_ERR_READ_ONLY == frozenset({"lastdllerror"})


def test_application_member_names() -> None:
    names = application_member_names()
    # Implicit Application members (Range, Cells, Calculate, ...) used as bare calls.
    assert "range" in names
    assert "calculate" in names
    assert "cells" in names
    assert len(names) == 39
    assert "name" not in names
    assert "screenupdating" not in names
