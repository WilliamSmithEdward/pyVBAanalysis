"""Wait until WAIT_ORACLE holds N verdicts, then run oracle_run.py over
each batch in turn (BATCH.json -> BATCH_oracle.json).

    python queue_oracle.py WAIT_ORACLE.json N BATCH [BATCH ...]
"""

import json
import subprocess
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
wait, n = sys.argv[1], int(sys.argv[2])
deadline = time.time() + 3 * 3600
while time.time() < deadline:
    try:
        if len(json.load(open(wait, encoding="utf-8"))) >= n:
            break
    except (OSError, ValueError):
        pass
    time.sleep(30)
for batch in sys.argv[3:]:
    subprocess.run([sys.executable, str(HERE / "oracle_run.py"), f"{batch}.json", f"{batch}_oracle.json"],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    subprocess.run([sys.executable, str(HERE / "leading_dot.py"), f"{batch}_oracle.json", f"{batch}_oracle.json"],
                   stdout=subprocess.DEVNULL)
    print(batch, "done", flush=True)
