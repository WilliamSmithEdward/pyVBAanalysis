"""Random callers and callees around a Long and a Collection: each callee
reads, writes, Sets or forwards its parameter, or changes a module
variable, ByRef or ByVal; the caller then divides by the Long, reads the
Collection's Count or indexes a fixed array. Tests #449's kept-across-
a-call values against VBA.

    python gen_cfz.py CASES.json N SEED
"""

import json
import random
import sys

from gen_dtc import std

LONG_BODIES = [
    "Debug.Print p",
    "p = 0",
    "p = 2",
    "p = p + 1",
    "If p > 0 Then p = 0",
    "mn = 0",
    "mn = 3",
    "Fwd p",
    "Debug.Print Abs(p)",
    "",
]
COLL_BODIES = [
    "Debug.Print p Is Nothing",
    "Set p = New Collection",
    "Set p = Nothing",
    "If p Is Nothing Then Set p = New Collection",
    "If Not p Is Nothing Then p.Add 1",
    "Set mc = New Collection",
    "FwdC p",
    "",
]
FIXED = (
    "Private Sub Fwd(ByRef q As Long)\r\nq = 1\r\nEnd Sub\r\n"
    "Private Sub FwdC(ByRef q As Collection)\r\nSet q = New Collection\r\nEnd Sub\r\n"
)


def one(rng: random.Random, k: int) -> dict:
    procs = []
    calls = []
    for j in range(rng.randint(1, 3)):
        if rng.random() < 0.5:
            mode = rng.choice(["ByRef ", "ByVal ", ""])
            body = rng.choice(LONG_BODIES)
            procs.append(f"Private Sub L{j}({mode}p As Long)\r\n{body}\r\nEnd Sub\r\n")
            arg = rng.choice(["n", "mn", "(n)"])
            form = rng.choice(["L{j} {a}", "Call L{j}({a})"])
            calls.append(form.format(j=j, a=arg))
        else:
            mode = rng.choice(["ByRef ", "ByVal ", ""])
            body = rng.choice(COLL_BODIES)
            procs.append(f"Private Sub C{j}({mode}p As Collection)\r\n{body}\r\nEnd Sub\r\n")
            arg = rng.choice(["c", "mc"])
            form = rng.choice(["C{j} {a}", "Call C{j}({a})"])
            calls.append(form.format(j=j, a=arg))
    lines = ["Dim n As Long, c As Collection, arr(0 To 3) As Long"]
    lines.append(rng.choice(["n = 0", "n = 2", "n = 5", ""]))
    lines.append(rng.choice(["mn = 0", "mn = 4", ""]))
    lines.append(rng.choice(["Set c = New Collection", "Set mc = New Collection", "", "Set c = Nothing"]))
    lines += calls
    lines.append(rng.choice([
        "Main = 10 \\ n", "Main = 10 \\ mn", "Main = c.Count", "Main = mc.Count",
        "Main = arr(n)", "Main = arr(mn)", "Main = c.Count + 10 \\ n",
    ]))
    lines = [ln for ln in lines if ln]
    case = std(f"kcz{SEED}/{k:04d}", "Private mn As Long\r\nPrivate mc As Collection\r\n",
               "\n".join(lines), "".join(procs) + FIXED)
    return case


if __name__ == "__main__":
    out, n, SEED = sys.argv[1], int(sys.argv[2]), int(sys.argv[3])
    rng = random.Random(SEED)
    cases = [one(rng, k) for k in range(n)]
    json.dump(cases, open(out, "w", encoding="utf-8"), indent=1)
    print(len(cases), "cases")
