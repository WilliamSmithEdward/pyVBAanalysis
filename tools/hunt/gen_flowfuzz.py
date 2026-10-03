"""Random control flow over Long, String and a fixed array, ending in
operations that raise for some values: 10 \\ a, arr(b), Mid(s, a),
CInt(b * 10000). Tests flow-sensitive value tracking against VBA.

    python gen_flowfuzz.py CASES.json N SEED
"""

import json
import random
import sys

from gen_dtc import std

VARS = ["a", "b"]


def val(rng: random.Random) -> str:
    return str(rng.choice([-1, 0, 1, 2, 3, 4, 5]))


def cond(rng: random.Random) -> str:
    v = rng.choice(VARS)
    return rng.choice([f"{v} = {val(rng)}", f"{v} > {val(rng)}", f"{v} <> {val(rng)}", f'Len(s) > {val(rng)}',
                       "True", "False", f"{v} Mod 2 = 0"])


def assign(rng: random.Random) -> str:
    p = rng.random()
    v = rng.choice(VARS)
    if p < 0.5:
        return f"{v} = {val(rng)}"
    if p < 0.7:
        return f"{v} = {v} + {rng.choice([-1, 1, 2])}"
    if p < 0.85:
        return f's = "{rng.choice(["", "a", "abc", "abcdef"])}"'
    return f"{v} = {rng.choice(VARS)}"


def stmt(rng: random.Random, depth: int) -> list:
    p = rng.random()
    if depth <= 0 or p < 0.45:
        return [assign(rng)]
    if p < 0.65:
        out = [f"If {cond(rng)} Then"] + stmt(rng, depth - 1)
        if rng.random() < 0.5:
            out += ["Else"] + stmt(rng, depth - 1)
        return out + ["End If"]
    if p < 0.75:
        out = [f"Select Case {rng.choice(VARS)}", f"Case {val(rng)}"] + stmt(rng, depth - 1)
        out += ["Case Else"] + stmt(rng, depth - 1)
        return out + ["End Select"]
    if p < 0.88:
        body = stmt(rng, depth - 1)
        if rng.random() < 0.4:
            body += [f"If {cond(rng)} Then Exit For"]
        return [f"For i = {rng.choice([0, 1])} To {rng.choice([0, 1, 2, 3])}"] + body + ["Next"]
    LABEL[0] += 1
    lab = f"Skip{LABEL[0]}"
    return [f"If {cond(rng)} Then GoTo {lab}"] + stmt(rng, depth - 1) + [f"{lab}:"]


LABEL = [0]


FINAL = ["Main = 10 \\ a", "Main = arr(b)", "Main = Mid(s, a)", "Main = CInt(b * 10000)",
         "Main = Left(s, b)", "Main = 10 / (a - b)"]


def main(n: int, seed: int) -> list:
    rng = random.Random(seed)
    cases = []
    for k in range(n):
        lines = ["Dim a As Long, b As Long, i As Long, s As String", "Dim arr(0 To 3) As Long"]
        lines += [assign(rng) for _ in range(2)]
        for _ in range(rng.randint(1, 3)):
            lines += stmt(rng, 2)
        lines.append(rng.choice(FINAL))
        cases.append(std(f"flz{seed}/{k:04d}", "", "\n".join(lines)))
    return cases


if __name__ == "__main__":
    out = main(int(sys.argv[2]), int(sys.argv[3]))
    json.dump(out, open(sys.argv[1], "w", encoding="utf-8"), indent=1)
    print(len(out), "cases")
