"""Excel's verdict against the full-wrapper findings, one line per case,
with FP (finding on a clean run) and FN (raise or compile error with no
finding) marked.

    python table3.py ORACLE.json WRAPPER.jsonl [--diffs]
"""

import json
import sys

NOISE = ("unused", "variable-never", "missing-return", "option-explicit")


def rows(oracle_path: str, found_path: str) -> list[tuple[str, str, str, list[str]]]:
    """(flag, label, verdict, error codes) per oracle case; flag is ok, FP or FN."""
    oracle = json.load(open(oracle_path, encoding="utf-8"))
    found: dict[str, list[str]] = {}
    for line in open(found_path, encoding="utf-8"):
        row = json.loads(line)
        found.setdefault(row["label"], []).extend(
            c.split(":")[0] for c in row["codes"] if not c.startswith(NOISE))
    out = []
    for label, verdict in oracle.items():
        codes = found.get(label, [])
        run = verdict.get("run") or {}
        compiled = verdict.get("compile") == "accepted"
        raised = not compiled or run.get("outcome") != "passed"
        if not compiled:
            what = "COMPILE " + (verdict.get("message") or "").replace("\n", " ").replace("Compile error:", "").strip()[:40]
        elif run.get("outcome") != "passed":
            what = f"E{run.get('number')}"
        else:
            what = "ok " + str(run.get("value"))[:20]
        flag = "ok" if raised == bool(codes) else ("FP" if codes else "FN")
        out.append((flag, label, what, codes))
    return out


if __name__ == "__main__":
    table = rows(sys.argv[1], sys.argv[2])
    for flag, label, what, codes in table:
        if flag == "ok" and "--diffs" in sys.argv:
            continue
        print(f"{flag:4}{label:36} {what:50} {codes}".encode("ascii", "backslashreplace").decode())
    print(f"{sum(1 for r in table if r[0] == 'ok')} of {len(table)} agree")
