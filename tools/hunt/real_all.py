"""Run project_probe over corpus parts with one XLIDE tree and write every
module's findings, tagged with the part.

    python real_all.py ROOT OUT.jsonl PART.json [...]
"""

import json
import os
import subprocess
import sys
from pathlib import Path

PROBE = str(Path(__file__).resolve().parents[1] / "differential" / "upstream" / "project_probe.mjs")
root, out_path = sys.argv[1], sys.argv[2]
with open(out_path, "w", encoding="utf-8") as out:
    for part in sys.argv[3:]:
        env = dict(os.environ, XLIDE_ROOT=root)
        # npx is a .cmd shim on Windows, which only a shell runs.
        res = subprocess.run(["npx", "-y", "tsx", PROBE, part], capture_output=True, text=True,
                             encoding="utf-8", env=env, shell=os.name == "nt", timeout=1800)
        n = 0
        for line in res.stdout.splitlines():
            try:
                row = json.loads(line)
            except json.JSONDecodeError:
                continue
            row["part"] = os.path.basename(part)
            out.write(json.dumps(row) + "\n")
            n += 1
        print(part, n, "modules", flush=True)
