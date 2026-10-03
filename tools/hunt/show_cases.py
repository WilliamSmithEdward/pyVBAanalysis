"""Print merged cases with Excel's verdict and a tree's findings.

    python show_cases.py ALL.json ALL_oracle.json FOUND.jsonl LABEL [...]
"""

import json
import sys

cases = {c["label"]: c for c in json.load(open(sys.argv[1], encoding="utf-8"))}
oracle = json.load(open(sys.argv[2], encoding="utf-8"))
found: dict = {}
for line in open(sys.argv[3], encoding="utf-8"):
    row = json.loads(line)
    found.setdefault(row["label"], []).extend(row["codes"])
for label in sys.argv[4:]:
    v = oracle.get(label, {})
    print("=" * 70)
    print(label, "| compile:", v.get("compile"), (v.get("message") or "").replace("\n", " ")[:60], "| run:", v.get("run"))
    for m in cases[label]["modules"]:
        print(f"--- {m['name']} ({m.get('type')})")
        print(m["source"].replace("\r\n", "\n").rstrip())
    for c in found.get(label, []):
        if not c.startswith(("unused", "variable-never")):
            print("  >>", c[:200])
