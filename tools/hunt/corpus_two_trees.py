"""Run project_probe.mjs over corpus parts with two XLIDE trees and keep each
finding that only one of them reports.

    python corpus_two_trees.py OLD_ROOT NEW_ROOT OUT.jsonl PART.json [...]

Output rows: {part, label, module, side: "new-only"|"old-only", finding}.
"""

import json
import os
import subprocess
import sys
from collections import Counter
from pathlib import Path

PROBE = str(Path(__file__).resolve().parents[1] / "differential" / "upstream" / "project_probe.mjs")


def run(root: str, part: str) -> dict[tuple[str, str], Counter]:
    env = dict(os.environ, XLIDE_ROOT=root)
    # npx is a .cmd shim on Windows, which only a shell runs.
    out = subprocess.run(["npx", "-y", "tsx", PROBE, part], capture_output=True, text=True, encoding="utf-8",
                         env=env, shell=os.name == "nt", timeout=900)
    rows: dict[tuple[str, str], Counter] = {}
    for line in out.stdout.splitlines():
        try:
            row = json.loads(line)
        except json.JSONDecodeError:
            continue
        rows[(row["label"], row["module"])] = Counter(row["findings"])
    if not rows:
        print(f"{part}: no output from {root}; stderr: {out.stderr[-400:]}", flush=True)
    return rows


old_root, new_root, out_path = sys.argv[1:4]
with open(out_path, "a", encoding="utf-8") as out:
    for part in sys.argv[4:]:
        old, new = run(old_root, part), run(new_root, part)
        added = removed = 0
        for key in set(old) | set(new):
            a, b = old.get(key, Counter()), new.get(key, Counter())
            for finding in (b - a).elements():
                added += 1
                out.write(json.dumps({"part": os.path.basename(part), "label": key[0], "module": key[1],
                                      "side": "new-only", "finding": finding}) + "\n")
            for finding in (a - b).elements():
                removed += 1
                out.write(json.dumps({"part": os.path.basename(part), "label": key[0], "module": key[1],
                                      "side": "old-only", "finding": finding}) + "\n")
        print(f"{os.path.basename(part)}: {len(new)} modules, +{added} -{removed}", flush=True)
