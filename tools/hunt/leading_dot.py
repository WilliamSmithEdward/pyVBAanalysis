"""List oracle verdicts the harness lost to a leading-dot number
(pyVBAharness serializes 0.5 as .5), and recover the value.

    python leading_dot.py ORACLE.json [FIXED.json]

With FIXED, writes a copy whose lost verdicts become passed runs.
"""

import json
import re
import sys

oracle = json.load(open(sys.argv[1], encoding="utf-8"))
LOST = re.compile(r"unreadable result: '(.*)'$", re.S)
fixed = 0
for label, verdict in oracle.items():
    run = verdict.get("run") or {}
    m = LOST.search(run.get("message") or "")
    if run.get("outcome") != "runner-error" or not m:
        continue
    text = re.sub(r'(?<=[:\[,])(-?)\.(\d)', r"\g<1>0.\2", m.group(1))
    try:
        payload = json.loads(text)
    except ValueError:
        print("still unreadable", label, m.group(1)[:80])
        continue
    print(label, payload.get("outcome"), payload.get("value"))
    run.update(outcome=payload.get("outcome"), value=payload.get("value"), message="")
    fixed += 1
print(fixed, "recovered")
if len(sys.argv) > 2:
    json.dump(oracle, open(sys.argv[2], "w", encoding="utf-8"), indent=1)
