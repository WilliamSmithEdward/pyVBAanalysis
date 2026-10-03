"""Drop harness-error and runner-error verdicts from an oracle file, so the
next oracle_run.py over it measures those cases again. Labels given after
the file are kept even when they failed.

    python drop_harness.py ORACLE.json [KEEP_LABEL ...]
"""

import json
import sys

path, keep = sys.argv[1], set(sys.argv[2:])
o = json.load(open(path, encoding="utf-8"))
out = {}
dropped = 0
for k, v in o.items():
    bad = v.get("compile") == "harness-error" or (v.get("run") or {}).get("outcome") == "runner-error"
    if bad and k not in keep:
        dropped += 1
        continue
    out[k] = v
json.dump(out, open(path, "w", encoding="utf-8"), indent=1)
print("dropped", dropped, "kept", len(out))
