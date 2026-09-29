"""Move the YARA Forge pin to the newest release that is at least a week old.

    python security/yara_forge_update.py

Reads security/yara-forge.json, asks GitHub for YARAHQ/yara-forge's releases,
and takes the newest one published seven or more days ago, the same cooldown
Dependabot applies. The asset's SHA-256 comes from the digest GitHub records
for it; the downloaded file must match it and hold the rule file the scan
reads. The pin is rewritten only when the release changes. Prints the release
it settled on, and writes `release=` and `changed=` to $GITHUB_OUTPUT when set.
Uses GH_TOKEN, when set, for the API.
"""

from __future__ import annotations

import datetime as dt
import hashlib
import io
import json
import os
import sys
import urllib.request
import zipfile
from pathlib import Path

PIN = Path(__file__).resolve().parent / "yara-forge.json"
API = "https://api.github.com/repos/YARAHQ/yara-forge/releases?per_page=30"
COOLDOWN = dt.timedelta(days=7)


# The only hosts fetched from: the GitHub API, and the release download host
# its asset URLs name (GitHub redirects that one to its object storage).
_HOSTS = ("https://api.github.com/", "https://github.com/YARAHQ/yara-forge/releases/download/")


def _get(url: str, accept: str = "application/vnd.github+json") -> bytes:
    if not url.startswith(_HOSTS):
        raise SystemExit(f"refusing to fetch {url}")
    request = urllib.request.Request(url, headers={"Accept": accept, "User-Agent": "pyvbaanalysis-yara-forge-update"})
    token = os.environ.get("GH_TOKEN")
    if token and url.startswith(_HOSTS[0]):
        request.add_header("Authorization", f"Bearer {token}")
    # Known acceptable (SECURITY.md): the URL is checked against _HOSTS above, so
    # no file: or other scheme reaches urlopen.
    with urllib.request.urlopen(request, timeout=120) as response:  # nosemgrep: python.lang.security.audit.dynamic-urllib-use-detected.dynamic-urllib-use-detected
        return response.read()


def main() -> int:
    pin = json.loads(PIN.read_text(encoding="utf-8"))
    now = dt.datetime.now(dt.timezone.utc)
    releases = json.loads(_get(API))
    eligible = [
        r
        for r in releases
        if not r["draft"]
        and not r["prerelease"]
        and now - dt.datetime.fromisoformat(r["published_at"].replace("Z", "+00:00")) >= COOLDOWN
    ]
    if not eligible:
        raise SystemExit("no YARA Forge release is a week old")
    newest = max(eligible, key=lambda r: r["published_at"])
    release = newest["tag_name"]
    changed = release != pin["release"]
    if changed:
        asset = next((a for a in newest["assets"] if a["name"] == pin["asset"]), None)
        if asset is None or not str(asset.get("digest", "")).startswith("sha256:"):
            raise SystemExit(f"{release} has no {pin['asset']} with a recorded SHA-256")
        expected = asset["digest"].removeprefix("sha256:")
        data = _get(asset["browser_download_url"], accept="application/octet-stream")
        actual = hashlib.sha256(data).hexdigest()
        if actual != expected:
            raise SystemExit(f"{pin['asset']} of {release}: SHA-256 {actual}, GitHub records {expected}")
        rule_file = f"packages/{pin['package']}/yara-rules-{pin['package']}.yar"
        if rule_file not in zipfile.ZipFile(io.BytesIO(data)).namelist():
            raise SystemExit(f"{pin['asset']} of {release} holds no {rule_file}")
        pin["release"] = release
        pin["sha256"] = expected
        PIN.write_text(json.dumps(pin, indent=2) + "\n", encoding="utf-8", newline="\n")
    print(f"YARA Forge {release} ({'new pin' if changed else 'already pinned'})")
    output = os.environ.get("GITHUB_OUTPUT")
    if output:
        with open(output, "a", encoding="utf-8") as handle:
            handle.write(f"release={release}\nchanged={'true' if changed else 'false'}\n")
    return 0


if __name__ == "__main__":
    sys.exit(main())
