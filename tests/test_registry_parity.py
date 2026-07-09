"""Drift gate: every catalogue rule must be emitted by the Python analyzer.

The rule catalogue (data/rule_metadata.json) is vendored verbatim from XLIDE, so
re-vendoring after an upstream release adds new rules to the catalogue as DATA
whether or not their rule functions have been ported. The metadata/audit
alignment tests all pass in that state, which is exactly the drift agent.md
section 6 says CI must flag: "a code in the audit ... has no Python rule
emitting it (unported code)".

Every ported rule reports through ``push("<ruleName>", ...)`` with a literal
rule-name string (the DIAGNOSTIC_RULES key), so a static scan of the diagnostics
package recovers the emitted-rule set without executing any rule. After an
upstream data re-pin, this gate turns red on precisely the rules that still
need porting.
"""

from __future__ import annotations

import re
from pathlib import Path

from pyvbaanalysis.diagnostics import DIAGNOSTIC_RULES

_DIAGNOSTICS_DIR = Path(__file__).resolve().parent.parent / "pyvbaanalysis" / "diagnostics"

# push("ruleName", ...) with a literal rule name.
_PUSH_RE = re.compile(r'push\(\s*"(\w+)"')
# rule = "a" if condition else "b" - a rule name selected at runtime between two
# literals and pushed via the variable (argument_inference.py's type-mismatch pair).
_SELECTED_RE = re.compile(r'"(\w+)"\s*\n?\s*if\s.+?\selse\s+"(\w+)"', re.DOTALL)

# Catalogue rules the Python port deliberately does not emit.
_NOT_PORTED = {
    # Emitted by the XLIDE extension layer (vbaModuleAnalysis/vbaTestRunner, the
    # '@xlide-test comment validator), not by the analyzer core this port mirrors.
    "vbaTestDirective",
}

# Rules emitted by the engine without going through a rule's push() callback,
# mapped to (file, marker) proving the emission still exists.
_DIRECT_EMITTERS = {
    # analyze_module.py constructs the malformed-suppression-directive
    # diagnostic directly from DIRECTIVE_DIAGNOSTIC_CODE after the
    # inline-suppression filter runs.
    "analysisSuppressionDirective": ("analyze_module.py", "DIRECTIVE_DIAGNOSTIC_CODE"),
}


def _scan() -> tuple[set[str], set[str]]:
    """(direct pushes, runtime-selected candidates) across the diagnostics package."""
    pushed: set[str] = set()
    selected: set[str] = set()
    for path in _DIAGNOSTICS_DIR.rglob("*.py"):
        text = path.read_text(encoding="utf-8")
        pushed |= set(_PUSH_RE.findall(text))
        for pair in _SELECTED_RE.findall(text):
            selected |= {name for name in pair if name in DIAGNOSTIC_RULES}
    return pushed, selected


def test_every_catalogue_rule_is_emitted() -> None:
    pushed, selected = _scan()
    emitted = pushed | selected | set(_DIRECT_EMITTERS)
    missing = set(DIAGNOSTIC_RULES) - emitted - _NOT_PORTED
    assert not missing, (
        f"catalogue rules with no Python emitter (unported after a data re-pin?): {sorted(missing)}"
    )


def test_no_orphan_pushed_rule_names() -> None:
    # A pushed name absent from the catalogue would KeyError at analysis time.
    pushed, _ = _scan()
    orphans = pushed - set(DIAGNOSTIC_RULES)
    assert not orphans, f"pushed rule names missing from the catalogue: {sorted(orphans)}"


def test_direct_emitters_still_exist() -> None:
    # The special-cased engine emissions must keep existing, or the exclusion
    # above would silently mask a genuinely unported rule.
    for rule_name, (filename, marker) in _DIRECT_EMITTERS.items():
        assert rule_name in DIAGNOSTIC_RULES
        text = (_DIAGNOSTICS_DIR / filename).read_text(encoding="utf-8")
        assert marker in text, f"{filename} no longer emits {rule_name} via {marker}"


def test_not_ported_set_stays_minimal() -> None:
    # If a deliberately-excluded rule gains a Python emitter, remove it from
    # _NOT_PORTED so the gate guards it again.
    pushed, selected = _scan()
    stale = _NOT_PORTED & (pushed | selected)
    assert not stale, f"_NOT_PORTED entries that now have emitters: {sorted(stale)}"
