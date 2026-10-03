"""Print the shortest cases that raise in Excel with no error finding.

    python short_fn.py CASES.json ORACLE.json FOUND.jsonl [N]
"""

import json
import sys

cases = {c["label"]: c for c in json.load(open(sys.argv[1], encoding="utf-8"))}
oracle = json.load(open(sys.argv[2], encoding="utf-8"))
found = {}
for line in open(sys.argv[3], encoding="utf-8"):
    r = json.loads(line)
    found.setdefault(r["label"], []).extend(r["codes"])
QUIET = ("unused", "variable-never", "constant", "dead", "unreachable")
n = int(sys.argv[4]) if len(sys.argv) > 4 else 5
rows = []
for label, v in oracle.items():
    run = v.get("run") or {}
    if run.get("outcome") != "vba-error":
        continue
    if any(not c.startswith(QUIET) for c in found.get(label, [])):
        continue
    src = "\r\n".join(m["source"] for m in cases[label]["modules"])
    rows.append((len(src), label, run["number"], src))
for _, label, num, src in sorted(rows)[:n]:
    print("=" * 60)
    print(label, "raises", num)
    print(src.replace("\r\n", "\n"))
