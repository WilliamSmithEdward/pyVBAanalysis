"""Tally error-severity findings on real code by rule, skipping the
should-fail test workbook, and print a few samples per rule.

    python real_errs_tally.py FINDINGS.jsonl ERROR_CODES.txt [RULE] [N]
"""

import json
import sys
from collections import Counter, defaultdict

codes = set(open(sys.argv[2], encoding="utf-8").read().split())
only = sys.argv[3] if len(sys.argv) > 3 else None
n = int(sys.argv[4]) if len(sys.argv) > 4 else 3
tally = Counter()
samples = defaultdict(list)
for line in open(sys.argv[1], encoding="utf-8"):
    r = json.loads(line)
    if "Should Fail" in r["label"] or "should_fail" in r["label"]:
        continue
    for f in r["findings"]:
        code = f.split(" @", 1)[0]
        if code not in codes:
            continue
        tally[code] += 1
        samples[code].append((r["part"], r["label"], r["module"], f))
for code, cnt in tally.most_common():
    if only and code != only:
        continue
    print(f"{cnt:5} {code}")
    if only:
        for s in samples[code][:n]:
            print("     ", s[0], s[1][-60:], s[2], s[3][:200])
