"""Print the shortest corpus cases that raise in Excel, were reported by OLD
and are not by NEW.

    python short_lost.py ALL.json ORACLE.json OLD.jsonl NEW.jsonl [N]
"""

import json
import sys

QUIET = ("unused", "variable-never", "constant", "dead", "unreachable")


def load(path):
    out = {}
    for line in open(path, encoding="utf-8"):
        r = json.loads(line)
        out.setdefault(r["label"], []).extend(c for c in r["codes"] if not c.startswith(QUIET))
    return out


cases = {c["label"]: c for c in json.load(open(sys.argv[1], encoding="utf-8"))}
oracle = json.load(open(sys.argv[2], encoding="utf-8"))
old, new = load(sys.argv[3]), load(sys.argv[4])
n = int(sys.argv[5]) if len(sys.argv) > 5 else 5
rows = []
for label, v in oracle.items():
    run = v.get("run") or {}
    if run.get("outcome") != "vba-error" or not old.get(label) or new.get(label):
        continue
    src = "\r\n".join(m["source"] for m in cases[label]["modules"])
    rows.append((len(src), label, run["number"], src))
for _, label, num, src in sorted(rows)[:n]:
    print("=" * 60)
    print(label, "raises", num, "old:", [c[:60] for c in old[label]])
    print(src.replace("\r\n", "\n"))
