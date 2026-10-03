"""List comments posted on an issue after it was closed (the ones that may
have gone unread), with the issue number, closed date and first line.

    python late_comments.py OUT.json
"""

import json
import subprocess
import sys

REPO = "WilliamSmithEdward/xlide_vscode"


def api(path):
    out = subprocess.run(["gh", "api", "--paginate", path], capture_output=True, text=True, encoding="utf-8", check=True).stdout
    # --paginate concatenates JSON arrays as ][ ; join them.
    return json.loads("[" + out.strip()[1:-1].replace("][", ",") + "]") if out.strip() else []


issues = api(f"repos/{REPO}/issues?state=closed&per_page=100")
closed = {i["number"]: i["closed_at"] for i in issues}
comments = api(f"repos/{REPO}/issues/comments?per_page=100")
late = []
for c in comments:
    n = int(c["issue_url"].rsplit("/", 1)[1])
    if n in closed and c["created_at"] > closed[n]:
        late.append({"issue": n, "closed": closed[n], "comment": c["created_at"], "url": c["html_url"],
                     "first": c["body"].splitlines()[0][:100] if c["body"] else ""})
late.sort(key=lambda r: r["issue"])
json.dump(late, open(sys.argv[1], "w", encoding="utf-8"), indent=1)
for r in late:
    print(r["issue"], r["comment"][:10], r["first"])
print(len(late), "late comments")
