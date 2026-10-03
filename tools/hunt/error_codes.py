"""Print the rule codes whose default severity is error, from a tree's
ruleMetadata.ts.

    python error_codes.py TREE
"""

import re
import sys

src = open(sys.argv[1] + "/src/analyzer/diagnostics/ruleMetadata.ts", encoding="utf-8").read()
codes = []
for m in re.finditer(r"code: '([^']+)'(.*?)\n\t\}", src, re.S):
    sev = re.search(r"defaultSeverity: '(\w+)'", m.group(2))
    if sev and sev.group(1) == "error":
        codes.append(m.group(1))
print("\n".join(sorted(codes)))
