"""Generate docs/diagnostics-catalogue.md from the diagnostic rule metadata.

The catalogue is generated so it cannot drift from the registry. Re-run after any
change to the rules or their metadata:

    python tools/generate_diagnostics_catalogue.py

It reads pyvbaanalysis.diagnostics.rule_metadata.DIAGNOSTIC_RULES (the same table
the engine uses for severities and the data/rule_metadata.json export) and writes a
grouped reference of every diagnostic code.
"""

from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_ROOT))

from pyvbaanalysis.diagnostics.rule_metadata import (  # noqa: E402
    DIAGNOSTIC_RULES,
    STRUCTURAL_DIAGNOSTIC_RULES,
)

_OUT = _ROOT / "docs" / "diagnostics-catalogue.md"

_KIND_LABEL = {
    "compile-error": "compile error",
    "deterministic-runtime-error": "runtime error",
    "runtime-risk": "runtime risk",
    "style": "style",
}


def _escape(text: str) -> str:
    return text.replace("|", r"\|")


def main() -> None:
    records = sorted(DIAGNOSTIC_RULES.values(), key=lambda m: (m.category.value, m.code))
    categories: dict[str, list[object]] = {}
    for record in records:
        categories.setdefault(record.category.value, []).append(record)

    lines: list[str] = []
    structural_codes = sorted(meta.code for meta in STRUCTURAL_DIAGNOSTIC_RULES.values())
    lines.append("# Diagnostic catalogue")
    lines.append("")
    lines.append(
        "The diagnostic codes pyVBAanalysis can emit, generated from the rule metadata "
        "(`tools/generate_diagnostics_catalogue.py`). This table lists the "
        f"{len(records)} rule-metadata codes across {len(categories)} categories. "
        f"A further {len(structural_codes)} structural block-balance codes ("
        + ", ".join(f"`{code}`" for code in structural_codes)
        + ") are emitted by the parser pass and are not in the metadata table, for a "
        f"full set of {len(records) + len(structural_codes)} codes."
    )
    lines.append("")
    lines.append(
        "Each code is reported only when it is provably correct; anything unknown or "
        "ambiguous stays quiet (the no-false-positive discipline). The **kind** column "
        "says what a code means: a *compile error* is rejected by the VBE compiler, a "
        "*runtime error* is a deterministic Run-time error, a *runtime risk* is a "
        "likely fault, and *style* is advisory."
    )
    lines.append("")
    lines.append(
        "Override a code's severity with `AnalyzeModuleOptions.severity_overrides` "
        "(or the `severity_overrides` argument of `analyze_project` / the reader "
        'functions), keyed by code. Use `"off"`, `"information"`, `"warning"`, or '
        '`"error"`; the allowed values per code are constrained by policy (some codes '
        "can be downgraded but not disabled). See [docs/usage.md](usage.md).")
    lines.append("")

    for category in sorted(categories):
        group = categories[category]
        lines.append(f"## {category.capitalize()} ({len(group)})")
        lines.append("")
        lines.append("| Code | Title | Default | Kind | Spec reference |")
        lines.append("| --- | --- | --- | --- | --- |")
        for record in group:
            kind = _KIND_LABEL.get(record.diagnostic_kind.value, record.diagnostic_kind.value)
            spec = _escape(record.spec_reference) if record.spec_reference else ""
            lines.append(
                f"| `{record.code}` | {_escape(record.title)} | "
                f"{record.default_severity.value} | {kind} | {spec} |"
            )
        lines.append("")

    _OUT.write_text("\n".join(lines), encoding="utf-8")
    print(f"wrote {_OUT.relative_to(_ROOT)} ({len(records)} codes)")


if __name__ == "__main__":
    main()
