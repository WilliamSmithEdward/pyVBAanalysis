"""Print the source around a span in one module of a corpus part, and the
module types in that project.

    python show_span.py PART.json LABEL_SUFFIX MODULE START [LINES]
"""

import json
import sys

part, suffix, mod, start = sys.argv[1], sys.argv[2], sys.argv[3], int(sys.argv[4])
lines = int(sys.argv[5]) if len(sys.argv) > 5 else 8
for c in json.load(open(part, encoding="utf-8")):
    if not c["label"].endswith(suffix):
        continue
    print("project:", c["label"])
    print("modules:", [(m["name"], m.get("type")) for m in c["modules"]][:40])
    for m in c["modules"]:
        if m["name"] == mod:
            s = m["source"]
            ln = s[:start].count("\n")
            src = s.split("\n")
            for i in range(max(0, ln - lines), min(len(src), ln + 3)):
                print(f"{i + 1:6}{'>' if i == ln else ' '} {src[i].rstrip()}")
    break
