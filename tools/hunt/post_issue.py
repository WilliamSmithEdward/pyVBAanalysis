"""Post an issue or comment to xlide_vscode only if its body passes the house
checks: no line over 78 characters, and no line outside a fence with an odd
number of backticks (a code span split across a wrap).

    python post_issue.py issue "TITLE" BODY.md
    python post_issue.py comment NUMBER BODY.md

Prints the problems and posts nothing when a check fails.
"""

import os
import subprocess
import sys
from pathlib import Path

REPO = os.environ.get("ISSUE_REPO", "WilliamSmithEdward/xlide_vscode")


def problems(text: str) -> list[str]:
    out = []
    fenced = False
    for n, line in enumerate(text.splitlines(), 1):
        if line.startswith("```"):
            fenced = not fenced
            continue
        if len(line) > 78:
            out.append(f"line {n} is {len(line)} characters")
        if not fenced and line.count("`") % 2:
            out.append(f"line {n} splits a code span: {line}")
    if fenced:
        out.append("a code fence is never closed")
    return out


kind, target, body_path = sys.argv[1], sys.argv[2], sys.argv[3]
body = Path(body_path).read_text(encoding="utf-8")
found = problems(body)
if found:
    print("NOT POSTED:")
    for p in found:
        print("  " + p)
    sys.exit(1)
if kind == "issue":
    cmd = ["gh", "issue", "create", "-R", REPO, "--title", target, "--body-file", body_path]
elif kind == "comment":
    cmd = ["gh", "issue", "comment", target, "-R", REPO, "--body-file", body_path]
else:
    sys.exit(f"unknown kind {kind!r}")
# The operator's own title and body file go to gh as list arguments, with no
# shell, which is this script's purpose (listed in SECURITY.md).
result = subprocess.run(cmd, capture_output=True, text=True)  # nosemgrep: python.lang.security.audit.dangerous-subprocess-use-tainted-env-args.dangerous-subprocess-use-tainted-env-args
print(result.stdout.strip() or result.stderr.strip())
sys.exit(result.returncode)
