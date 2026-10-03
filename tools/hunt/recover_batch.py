"""Recover a batch whose cases file was overwritten, from the merged
all.json and all_oracle.json, into NEWNAME.json and NEWNAME_oracle.json.
With --list, print which of the given batch names all.json holds.

    python recover_batch.py ALL.json ALL_oracle.json BATCH NEWNAME
    python recover_batch.py ALL.json ALL_oracle.json --list NAME [NAME ...]
"""

import json
import sys
from collections import Counter

cases = json.load(open(sys.argv[1], encoding="utf-8"))
oracle = json.load(open(sys.argv[2], encoding="utf-8"))
if sys.argv[3] == "--list":
    # Windows file names ignore case, so chc.json and chC.json are one file.
    have = Counter(c["label"].split("::", 1)[0].lower() for c in cases)
    for name in sys.argv[4:]:
        print(name, have.get(name.lower(), 0))
    sys.exit()
batch, new = sys.argv[3], sys.argv[4]
out, verdicts = [], {}
for c in cases:
    head, _, rest = c["label"].partition("::")
    if head != batch:
        continue
    c = dict(c, label=rest)
    out.append(c)
    verdicts[rest] = oracle[f"{batch}::{rest}"]
json.dump(out, open(f"{new}.json", "w", encoding="utf-8"), indent=1)
json.dump(verdicts, open(f"{new}_oracle.json", "w", encoding="utf-8"), indent=1)
print(len(out), "cases recovered into", new)
