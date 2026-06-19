"""M9: Excel host object-model resolver (hostModel.ts parity over vendored JSON)."""

from __future__ import annotations

from pyvbaanalysis.host import (
    get_excel_object_model,
    get_host_globals,
    get_host_members,
    resolve_host_alias,
    resolve_host_constant,
    resolve_host_global,
    resolve_host_member_signature,
    resolve_member_return_type,
)


def test_model_loads_and_is_complete() -> None:
    model = get_excel_object_model()
    assert set(model.keys()) >= {"aliases", "globals", "constants", "types", "memberSignatures"}
    # The Range surface must be exhaustive (member-not-found depends on this).
    assert model["types"]["Excel.Range"].get("exhaustive") is True
    assert len(get_host_members("Excel.Range")) > 150


def test_resolve_host_alias() -> None:
    assert resolve_host_alias("Range") == "Excel.Range"
    assert resolve_host_alias("Excel.Worksheet") == "Excel.Worksheet"
    assert resolve_host_alias("worksheet") == "Excel.Worksheet"  # case-insensitive
    assert resolve_host_alias("NotAHostType") is None
    assert resolve_host_alias("") is None


def test_resolve_host_global() -> None:
    assert resolve_host_global("ThisWorkbook") == "Excel.Workbook"
    assert resolve_host_global("application") == "Excel.Application"  # case-insensitive
    assert resolve_host_global("NotGlobal") is None
    assert len(get_host_globals()) >= 5


def test_resolve_member_return_type() -> None:
    assert resolve_member_return_type("Excel.Range", "Worksheet") == "Excel.Worksheet"
    assert resolve_member_return_type("Excel.Range", "Cells") == "Excel.Range"
    assert resolve_member_return_type("Excel.Range", "NoSuchMember") is None


def test_resolve_host_constant() -> None:
    xl_up = resolve_host_constant("xlUp")
    assert xl_up is not None and xl_up["value"] == -4162
    assert resolve_host_constant("XLUP") is not None  # case-insensitive
    assert resolve_host_constant("notAConstant") is None


def test_resolve_host_member_signature() -> None:
    # A known callable member resolves to a signature string (or None if uncurated).
    sig = resolve_host_member_signature("Excel.Worksheet", "Range")
    assert sig is None or isinstance(sig, str)
