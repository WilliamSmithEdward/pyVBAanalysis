"""M6: lexical rule family (lexical.ts parity) + oracle-asserted validation."""

from __future__ import annotations

from pyvbaanalysis.diagnostics import analyze_module
from pyvbaanalysis.evidence import load_audit, load_oracle_cases

_AUDIT = {a.code: a for a in load_audit()}
_CASES = {c.id: c for c in load_oracle_cases()}
_LEXICAL_CODES = ("unterminated-string", "invalid-line-continuation")


def _codes(source: str) -> list[str]:
    return [d.code for d in analyze_module(source)]


def test_unterminated_string_fires() -> None:
    assert "unterminated-string" in _codes('s = "abc')
    assert "unterminated-string" in _codes('Debug.Print "no close')


def test_terminated_string_is_clean() -> None:
    assert "unterminated-string" not in _codes('s = "abc"')
    assert "unterminated-string" not in _codes('s = "he said ""hi"""')  # doubled quotes are escapes


def test_mid_line_continuation_underscore_fires() -> None:
    assert "invalid-line-continuation" in _codes("x = 1 _ 2")


def test_valid_line_continuation_is_clean() -> None:
    assert "invalid-line-continuation" not in _codes("x = 1 + _\n    2")


def test_trailing_identifier_underscore_is_clean() -> None:
    # `x_` is a valid identifier; the underscore is not a continuation marker.
    assert "invalid-line-continuation" not in _codes("Dim x_ As Long\nDim y As Long")


def test_continuation_inside_string_ignored() -> None:
    # An underscore inside a string literal is not a continuation marker.
    assert "invalid-line-continuation" not in _codes('s = "a _ b"')


def test_lexical_codes_are_spec_derived() -> None:
    # These are MS-VBAL spec-derived (not Excel/VBE-oracle), so they carry no
    # asserted oracle cases; the focused tests above encode the spec behavior.
    for code in _LEXICAL_CODES:
        assert _AUDIT[code].status == "spec-derived"
        assert _AUDIT[code].asserted_oracle_cases == ()


def test_no_lexical_false_positives_on_accepted_cases() -> None:
    # Valid (compile-accepted) VBA must never trigger a lexical diagnostic.
    lexical = set(_LEXICAL_CODES)
    accepted = [c for c in _CASES.values() if c.expected == "accepted"]
    assert accepted  # the corpus has accepted controls
    for case in accepted:
        for module in case.modules:
            emitted = {d.code for d in analyze_module(module.source)}
            spurious = emitted & lexical
            assert not spurious, f"{case.id}::{module.name}: lexical false positive {spurious}"
