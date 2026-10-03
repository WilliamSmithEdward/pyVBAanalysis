"""Merge every batch with Excel verdicts (X.json beside X_oracle.json) into
one cases file and one oracle, labels prefixed with the batch name, so a
tree can be probed in a single run.

    python merge_batches.py ALL.json ALL_oracle.json
"""

import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
cases_out, oracle_out = [], {}
for oracle in sorted(HERE.glob("*_oracle.json")):
    path = oracle.with_name(oracle.name.replace("_oracle.json", ".json"))
    if not path.exists() or path.name == Path(sys.argv[1]).name:
        continue
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        verdicts = json.loads(oracle.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, UnicodeDecodeError):
        continue
    if not (isinstance(data, list) and data and isinstance(data[0], dict) and "modules" in data[0]):
        continue
    batch = path.stem
    for case in data:
        verdict = verdicts.get(case["label"])
        if verdict is None:
            continue
        label = f"{batch}::{case['label']}"
        cases_out.append({**case, "label": label})
        oracle_out[label] = verdict
json.dump(cases_out, open(sys.argv[1], "w", encoding="utf-8"))
json.dump(oracle_out, open(sys.argv[2], "w", encoding="utf-8"))
print(len(cases_out), "cases from", len({c["label"].split("::")[0] for c in cases_out}), "batches")
