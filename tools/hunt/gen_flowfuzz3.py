"""Control-flow fuzz, third kind: On Error Resume Next and On Error GoTo 0
in branches and loops around operations that raise for some values.

    python gen_flowfuzz3.py CASES.json N SEED
"""

import json
import random
import sys

from gen_dtc import std

LABEL = [0]
LOOPS = [0]
RAISERS = ["n = 10 \\ a", "n = arr(b)", "n = Len(Mid(s, a))", "n = CInt(b * 10000)", "n = c.Count"]


def val(rng: random.Random) -> str:
    return str(rng.choice([-1, 0, 1, 2, 5]))


def cond(rng: random.Random) -> str:
    return rng.choice([f"a = {val(rng)}", f"a > {val(rng)}", f"b <> {val(rng)}", "True", "False",
                       "c Is Nothing", "Len(s) > 1"])


def simple(rng: random.Random) -> str:
    p = rng.random()
    if p < 0.25:
        return rng.choice(["On Error Resume Next", "On Error GoTo 0"])
    if p < 0.5:
        return rng.choice(RAISERS)
    return rng.choice([f"a = {val(rng)}", f"b = {val(rng)}", 's = "abc"', "Set c = New Collection",
                       "a = a + 1"])


def stmt(rng: random.Random, depth: int) -> list:
    p = rng.random()
    if depth <= 0 or p < 0.5:
        return [simple(rng)]
    if p < 0.72:
        out = [f"If {cond(rng)} Then"] + stmt(rng, depth - 1)
        if rng.random() < 0.5:
            out += ["Else"] + stmt(rng, depth - 1)
        return out + ["End If"]
    if p < 0.86:
        LOOPS[0] += 1
        kv = f"k{LOOPS[0]}"
        return [f"For {kv} = 1 To {rng.choice([0, 1, 2])}"] + stmt(rng, depth - 1) + ["Next"]
    LABEL[0] += 1
    lab = f"Skip{LABEL[0]}"
    return [f"If {cond(rng)} Then GoTo {lab}"] + stmt(rng, depth - 1) + [f"{lab}:"]


def main(n: int, seed: int) -> list:
    rng = random.Random(seed)
    cases = []
    for k in range(n):
        start = LOOPS[0]
        body = []
        for _ in range(rng.randint(2, 4)):
            body += stmt(rng, 2)
        body.append(rng.choice(RAISERS))
        lines = ["Dim a As Long, b As Long, n As Long, s As String", "Dim arr(0 To 3) As Long, c As Collection"]
        counters = [f"k{j} As Long" for j in range(start + 1, LOOPS[0] + 1)]
        if counters:
            lines.append("Dim " + ", ".join(counters))
        lines += body
        lines.append("Main = n")
        cases.append(std(f"fle{seed}/{k:04d}", "", "\n".join(lines)))
    return cases


if __name__ == "__main__":
    out = main(int(sys.argv[2]), int(sys.argv[3]))
    json.dump(out, open(sys.argv[1], "w", encoding="utf-8"), indent=1)
    print(len(out), "cases")
