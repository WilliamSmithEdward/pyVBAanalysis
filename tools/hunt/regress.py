"""Cases where an older analyzer agreed with Excel and a newer one does not.

    python regress.py ORACLE.json OLD.jsonl NEW.jsonl [ORACLE OLD NEW ...]

Agreement is table3's: a finding exactly when Excel raised or refused.
"""

import json
import sys

NOISE = ("unused", "variable-never", "missing-return", "option-explicit")


def load(path: str) -> dict:
    found: dict = {}
    for line in open(path, encoding="utf-8"):
        row = json.loads(line)
        codes = [c.split(":")[0].split(" ")[0] for c in row.get("codes", row.get("findings", []))]
        found.setdefault(row["label"], []).extend(c for c in codes if not c.startswith(NOISE))
    return found


args = sys.argv[1:]
for i in range(0, len(args), 3):
    oracle = json.load(open(args[i], encoding="utf-8"))
    old, new = load(args[i + 1]), load(args[i + 2])
    for label, v in oracle.items():
        if label not in new:
            continue
        raised = v.get("compile") != "accepted" or (v.get("run") or {}).get("outcome") not in ("passed", None)
        if (v.get("run") or {}).get("outcome") == "runner-error":
            continue
        ok_old = raised == bool(old.get(label))
        ok_new = raised == bool(new.get(label))
        if ok_old and not ok_new:
            print("REGRESSED", label, "excel", "raises" if raised else "runs", "old", old.get(label), "new", new.get(label))
