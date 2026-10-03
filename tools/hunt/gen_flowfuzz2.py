"""Control-flow fuzz, second kind: a dynamic array and a Collection whose
state changes inside branches and loops (ReDim, Erase, Set, Add, Remove),
then operations that raise for some states.

    python gen_flowfuzz2.py CASES.json N SEED
"""

import json
import random
import sys

from gen_dtc import std

LABEL = [0]
LOOPS = [0]


def val(rng: random.Random) -> str:
    return str(rng.choice([-1, 0, 1, 2, 3]))


def cond(rng: random.Random) -> str:
    return rng.choice([f"a = {val(rng)}", f"a > {val(rng)}", "c Is Nothing", "Not c Is Nothing", "True", "False",
                       f"b <> {val(rng)}"])


def simple(rng: random.Random) -> str:
    return rng.choice([
        f"a = {val(rng)}", f"b = {val(rng)}", "a = a + 1", "b = a",
        f"ReDim arr({rng.choice([0, 1, 2, 3])})", f"ReDim Preserve arr({rng.choice([0, 2, 4])})", "Erase arr",
        "Set c = New Collection", "Set c = Nothing", "c.Add a", "c.Add b",
        "If Not c Is Nothing Then If c.Count > 0 Then c.Remove 1",
    ])


def stmt(rng: random.Random, depth: int) -> list:
    p = rng.random()
    if depth <= 0 or p < 0.45:
        return [simple(rng)]
    if p < 0.65:
        out = [f"If {cond(rng)} Then"] + stmt(rng, depth - 1)
        if rng.random() < 0.5:
            out += ["Else"] + stmt(rng, depth - 1)
        return out + ["End If"]
    if p < 0.75:
        return ["Select Case a", f"Case {val(rng)}"] + stmt(rng, depth - 1) + ["Case Else"] + stmt(rng, depth - 1) + ["End Select"]
    if p < 0.85:
        LOOPS[0] += 1
        kv = f"k{LOOPS[0]}"
        return [f"For {kv} = 1 To {rng.choice([0, 1, 2])}"] + stmt(rng, depth - 1) + ["Next"]
    if p < 0.93:
        # Each Do gets its own counter so a nested loop cannot reset an outer one.
        LOOPS[0] += 1
        kv = f"k{LOOPS[0]}"
        return [f"{kv} = 0", f"Do While {kv} < " + str(rng.choice([0, 1, 2]))] + stmt(rng, depth - 1) + [f"{kv} = {kv} + 1", "Loop"]
    LABEL[0] += 1
    lab = f"Skip{LABEL[0]}"
    return [f"If {cond(rng)} Then GoTo {lab}"] + stmt(rng, depth - 1) + [f"{lab}:"]


FINAL = ["Main = arr(b)", "Main = UBound(arr)", "Main = c.Count", "Main = c(1)", "Main = 10 \\ c.Count",
         "Main = arr(UBound(arr))", "Main = c(b)"]


def main(n: int, seed: int) -> list:
    rng = random.Random(seed)
    cases = []
    for k in range(n):
        start = LOOPS[0]
        body = []
        for _ in range(rng.randint(2, 4)):
            body += stmt(rng, 2)
        counters = [f"k{j} As Long" for j in range(start + 1, LOOPS[0] + 1)]
        lines = ["Dim a As Long, b As Long, i As Long", "Dim arr() As Long, c As Collection"]
        if counters:
            lines.append("Dim " + ", ".join(counters))
        lines += body
        lines.append(rng.choice(FINAL))
        cases.append(std(f"flx{seed}/{k:04d}", "", "\n".join(lines)))
    return cases


if __name__ == "__main__":
    out = main(int(sys.argv[2]), int(sys.argv[3]))
    json.dump(out, open(sys.argv[1], "w", encoding="utf-8"), indent=1)
    print(len(out), "cases")
