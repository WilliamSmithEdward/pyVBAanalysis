"""Download one commit of xlide_vscode from GitHub as a tarball and extract
it into a folder, without touching the local checkout.

    python export_remote.py SHA OUTDIR
"""

import io
import subprocess
import sys
import tarfile
from pathlib import Path

sha, out = sys.argv[1], Path(sys.argv[2])
data = subprocess.run(["gh", "api", f"repos/WilliamSmithEdward/xlide_vscode/tarball/{sha}"],
                      capture_output=True, check=True).stdout
out.mkdir(parents=True, exist_ok=True)
with tarfile.open(fileobj=io.BytesIO(data), mode="r:gz") as tar:
    members = tar.getmembers()
    prefix = members[0].name.split("/")[0] + "/"
    for m in members:
        if m.name.startswith(prefix):
            m.name = m.name[len(prefix):]
    tar.extractall(out, members=[m for m in members if m.name], filter="data")
print("extracted", len(members), "entries to", out)
