"""Random nested string functions over literals and String locals, with
numeric arguments that are themselves built from Len/InStr/arithmetic.
Tests the analyzer's string folding against VBA.

    python gen_strfuzz.py CASES.json N SEED
"""

import json
import random
import sys

from gen_dtc import std

LITS = ['"abc"', '""', '"a,b,c"', '" x "', '"ABCdef"', '"12"', 's', 't']
INSTR_FINDS = ['","', '"b"', '"z"', '""']
STRING_CHARS = ['"x"', '""', 'a']


def snum(rng: random.Random, depth: int) -> str:
    pick = rng.random()
    if depth <= 0 or pick < 0.35:
        return str(rng.choice([-1, 0, 1, 2, 3, 5, 10]))
    if pick < 0.55:
        return f"Len({sexpr(rng, depth - 1)})"
    if pick < 0.75:
        return f"InStr({sexpr(rng, depth - 1)}, {rng.choice(INSTR_FINDS)})"
    return f"({snum(rng, depth - 1)} {rng.choice(['+', '-', '*'])} {snum(rng, depth - 1)})"


def sexpr(rng: random.Random, depth: int) -> str:
    if depth <= 0 or rng.random() < 0.3:
        return rng.choice(LITS)
    f = rng.choice(["Left", "Right", "Mid", "Mid3", "Replace", "Trim", "UCase", "Space", "String",
                    "StrReverse", "Concat", "LCase", "Split"])
    a = sexpr(rng, depth - 1)
    if f in ("Left", "Right"):
        return f"{f}({a}, {snum(rng, depth - 1)})"
    if f == "Mid":
        return f"Mid({a}, {snum(rng, depth - 1)})"
    if f == "Mid3":
        return f"Mid({a}, {snum(rng, depth - 1)}, {snum(rng, depth - 1)})"
    if f == "Replace":
        return f"Replace({a}, {rng.choice(LITS[:6])}, {rng.choice(LITS[:6])})"
    if f == "Space":
        return f"Space({snum(rng, depth - 1)})"
    if f == "String":
        return f"String({snum(rng, depth - 1)}, {rng.choice(STRING_CHARS)})"
    if f == "Concat":
        return f"({a} & {sexpr(rng, depth - 1)})"
    if f == "Split":
        return f"Split({a}, \",\")({rng.choice([0, 1, 2, 3, -1])})"
    return f"{f}({a})"


def main(n: int, seed: int) -> list:
    rng = random.Random(seed)
    cases = []
    for k in range(n):
        e = sexpr(rng, 3)
        body = ('Dim s As String, t As String, a As String\ns = "q,r"\nt = ""\na = "z"\n'
                f"Main = {e}")
        cases.append(std(f"sfz{seed}/{k:04d}", "", body))
    return cases


if __name__ == "__main__":
    out = main(int(sys.argv[2]), int(sys.argv[3]))
    json.dump(out, open(sys.argv[1], "w", encoding="utf-8"), indent=1)
    print(len(out), "cases")
