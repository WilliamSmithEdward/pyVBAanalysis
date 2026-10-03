"""False positives by rule: corpus cases that compile and run clean in
Excel but get an error-class finding. Skips the advisory rules.

    python fp_rules.py ORACLE.json FOUND.jsonl [RULE]
"""

import json
import sys
from collections import Counter, defaultdict

ADVISORY = {"unreachable-code", "unused-variable", "unused-procedure", "variable-never-read",
            "unused-parameter", "unused-constant", "unused-label", "implicit-variant",
            "missing-option-explicit", "shadowed-name", "unused-enum", "unused-type"}
oracle = json.load(open(sys.argv[1], encoding="utf-8"))
found = defaultdict(list)
for line in open(sys.argv[2], encoding="utf-8"):
    if line.strip():
        row = json.loads(line)
        found[row["label"]].extend(row.get("codes") or [])
by_rule = Counter()
examples = defaultdict(list)
for label, codes in found.items():
    key = label.split("::", 1)[-1] if label not in oracle else label
    verdict = oracle.get(label) or oracle.get(key)
    if not verdict or verdict.get("compile") != "accepted":
        continue
    run = verdict.get("run") or {}
    if run.get("outcome") != "passed":
        continue
    rules = {c.split(":", 1)[0] for c in codes} - ADVISORY
    for r in rules:
        by_rule[r] += 1
        examples[r].append(label)
if len(sys.argv) > 3:
    print("\n".join(examples[sys.argv[3]]))
else:
    for r, n in by_rule.most_common():
        print(n, r, examples[r][:3])
