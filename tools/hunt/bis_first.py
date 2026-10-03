"""For each label, the first commit (in bis_commits.txt order) whose
non-advisory findings differ from the commit before it.

    python bis_first.py BASE_RUN.jsonl
"""

import json
import sys

NOISE = ("unused", "variable-never", "missing-return", "option-explicit", "unreachable")


def load(path: str) -> dict:
    out: dict = {}
    for line in open(path, encoding="utf-8"):
        if line.strip():
            row = json.loads(line)
            out.setdefault(row["label"], set()).update(
                c.split(":")[0] for c in row["codes"] if not c.startswith(NOISE))
    return out


commits = open("bis_commits.txt").read().split()
prev = load(sys.argv[1])
runs = [(s, load(f"bis_{s}.jsonl")) for s in commits]
labels = sorted(runs[-1][1])
for label in labels:
    before = prev.get(label, set())
    for sha, run in runs:
        now = run.get(label, set())
        if now != before:
            print(f"{label:40} {sha} {sorted(before)} -> {sorted(now)}")
        before = now
