"""Cases that move from a miss or false positive to agreement (and back)
between two runs, by batch.

    python gains.py ORACLE.json OLD.jsonl NEW.jsonl
"""

import sys
from collections import Counter

from table3 import rows


def verdicts(found: str) -> dict[str, str]:
    return {label: flag for flag, label, _what, _codes in rows(sys.argv[1], found)}


old, new = verdicts(sys.argv[2]), verdicts(sys.argv[3])
gained: Counter[str] = Counter()
lost: Counter[str] = Counter()
examples: dict[str, str] = {}
for label, v in new.items():
    w = old.get(label)
    batch = label.split("::", 1)[0]
    if w and w != "ok" and v == "ok":
        gained[batch] += 1
        examples.setdefault(batch, label)
    elif w == "ok" and v != "ok":
        lost[batch] += 1
print("gained", sum(gained.values()), dict(gained.most_common(15)))
print("lost", sum(lost.values()), dict(lost.most_common(15)))
print("examples", examples)
