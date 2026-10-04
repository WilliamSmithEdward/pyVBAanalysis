"""M9: Excel host object-model resolver (hostModel.ts parity over vendored JSON)."""

from __future__ import annotations

from pyvbaanalysis.host import (
    get_excel_object_model,
    get_host_globals,
    get_host_members,
    get_powerpoint_object_model,
    get_word_object_model,
    resolve_host_alias,
    resolve_host_constant,
    resolve_host_global,
    resolve_host_member_signature,
    resolve_member_return_type,
)
from pyvbaanalysis.host.host_model import (
    bare_type_name,
    get_host_enums,
    get_host_events,
    get_host_global_members,
    host_display_name,
    host_library_display_name,
    is_dispatch_only_host_type,
    is_host_member_name_anywhere,
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
    # XLIDE 2f49b93 (issue #197): Shapes.Range takes its own Index.
    assert resolve_host_member_signature("Excel.Shapes", "Range") == "Range(Index) As ShapeRange"


def test_dispatch_only_host_types() -> None:
    # XLIDE 2f49b93 (issue #198): only Excel's library has dispatch-only types.
    assert is_dispatch_only_host_type("Excel.Range") is True
    assert is_dispatch_only_host_type("excel.WORKBOOK") is False
    assert is_dispatch_only_host_type("Range") is False
    assert is_dispatch_only_host_type("Word.Range", get_word_object_model()) is False
    assert is_dispatch_only_host_type("Excel.Range", get_word_object_model()) is False


def test_host_model_helpers() -> None:
    assert bare_type_name("Excel.Range") == "Range"
    assert bare_type_name("Range") == "Range"
    assert host_display_name() == "Excel"
    assert host_display_name(get_word_object_model()) == "Word"
    assert host_library_display_name("Word.Application", get_excel_object_model()) == "Word"
    assert host_library_display_name("Range", get_word_object_model()) == "Word"
    assert host_library_display_name(None) == "Excel"
    powerpoint = get_powerpoint_object_model()
    assert [m["name"] for m in get_host_events("powerpoint.olecontrol", powerpoint)] == [
        "GotFocus",
        "LostFocus",
    ]
    assert get_host_events("Excel.NoSuchType") == []
    # Events count as member names anywhere, never as object-access members.
    assert is_host_member_name_anywhere("GOTFOCUS", powerpoint) is True
    assert is_host_member_name_anywhere("NoSuchMemberAnywhere") is False
    globals_ = [m["name"] for m in get_host_global_members()]
    assert "Union" in globals_ and not any(name.startswith("_") for name in globals_)
    assert any(e["displayName"] == "XlAxisType" for e in get_host_enums())


def test_library_tables() -> None:
    from pyvbaanalysis.host.event_signatures_data import HOST_EVENT_SIGNATURES
    from pyvbaanalysis.host.excel_library_names import EXCEL_LIBRARY_NAMES
    from pyvbaanalysis.host.excel_object_model import merge_host_constants
    from pyvbaanalysis.host.host_default_members import HOST_DEFAULT_MEMBERS
    from pyvbaanalysis.host.library_type_names import library_type_names
    from pyvbaanalysis.host.ms_forms_form_control_members import MSFORMS_FORM_CONTROL_MEMBERS
    from pyvbaanalysis.host.word_builtin_styles import WORD_BUILTIN_STYLES

    assert "activecell" in EXCEL_LIBRARY_NAMES and "version" not in EXCEL_LIBRARY_NAMES
    assert HOST_EVENT_SIGNATURES["Excel.Application"]["AfterCalculate"] == ""
    word_range = HOST_DEFAULT_MEMBERS["Word.Range"]
    assert (word_range.name, word_range.kind, word_range.writable) == ("Text", "property", True)
    assert "SetFocus" in MSFORMS_FORM_CONTROL_MEMBERS["MSForms.CheckBox"]
    assert "heading 1" in WORD_BUILTIN_STYLES
    vba = library_type_names("VBA")
    assert vba is not None and "collection" in vba and len(vba) == 29
    assert library_type_names("vba") is vba
    assert library_type_names("NoSuchLibrary") is None
    merged = merge_host_constants(
        {"xlA": {"name": "xlA", "value": 1}}, {"XLA": {"name": "XLA", "value": 2}}
    )
    assert merged == {"XLA": {"name": "XLA", "value": 2}}
