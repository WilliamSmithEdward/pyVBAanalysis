"""Split a cases file into N chunks: IN.json -> IN.0.json ... IN.{N-1}.json.

    python split_cases.py IN.json N
"""

import json
import sys

cases = json.load(open(sys.argv[1], encoding="utf-8"))
n = int(sys.argv[2])
stem = sys.argv[1][: -len(".json")]
for i in range(n):
    json.dump(cases[i::n], open(f"{stem}.{i}.json", "w", encoding="utf-8"))
print(len(cases), "cases in", n, "chunks")
