"""Random guards over known locals, each with a division by zero behind
it. Excel decides whether the guard holds; a report on a guard that is
false is a false positive, a silence on one that is true a miss.

    python gen_gdfuzz.py CASES.json N SEED
"""

import json
import random
import sys

from gen_dtc import std

DECL = "Dim k As Long, j As Integer, x As Double, s As String, t As String, b As Boolean, d As Date, z As Long, n As Long\n"
INIT = "k = {k}\nj = {j}\nx = {x}\ns = \"{s}\"\nt = \"{t}\"\nb = {b}\nd = #{d}#\n"

NUM_ATOMS = ["k", "j", "x", "Len(s)", "Len(t)", "Abs(k)", "Int(x)", "Fix(x)", "Round(x)", "Sgn(k)",
             "k Mod 3", "k \\ 2", "Year(d)", "Month(d)", "Day(d)", "Asc(s)", "Val(t)", "InStr(s, \"b\")",
             "CLng(x)", "CInt(j)", "k + j", "k * 2", "x * 2", "-k", "UBound(Split(s, \",\"))"]
STR_ATOMS = ["s", "t", "UCase(s)", "LCase(t)", "Left(s, 1)", "Mid(s, 2)", "s & t", "Trim(t)", "CStr(k)",
             "Format(k)", "Replace(s, \"a\", \"x\")", "StrReverse(s)"]
BOOL_ATOMS = ["b", "Not b", "IsNumeric(t)", "IsEmpty(s)", "IsDate(t)", "s Like \"a*\"", "t Like \"*1\"",
              "IsNumeric(s)"]
CMP = ["=", "<>", "<", ">", "<=", ">="]


def num_cmp(rng):
    return f"{rng.choice(NUM_ATOMS)} {rng.choice(CMP)} {rng.choice(['0', '1', '2', '3', '5', '7', '10', '-1', '2000', 'k', 'j'])}"


def str_cmp(rng):
    return f"{rng.choice(STR_ATOMS)} {rng.choice(['=', '<>', '<', '>'])} \"{rng.choice(['a', 'ab', 'b', 'A', 'abc', '', '12', 'x'])}\""


def guard(rng, depth=0):
    r = rng.random()
    if depth < 1 and r < 0.25:
        return f"({guard(rng, depth + 1)}) {rng.choice(['And', 'Or'])} ({guard(rng, depth + 1)})"
    if r < 0.6:
        return num_cmp(rng)
    if r < 0.85:
        return str_cmp(rng)
    return rng.choice(BOOL_ATOMS)


def one(rng, k, seed):
    init = INIT.format(k=rng.choice([0, 1, 3, 7, -2]), j=rng.choice([0, 2, 5]), x=rng.choice(["0", "1.5", "2.5", "-0.5", "3"]),
                       s=rng.choice(["a", "ab", "abc", "a,b", "B", ""]), t=rng.choice(["", "12", "x", " 1 ", "1/2/2000"]),
                       b=rng.choice(["True", "False"]), d=rng.choice(["1/2/2000", "12/31/1999", "6/15/2010"]))
    form = rng.random()
    g = guard(rng)
    if form < 0.5:
        body = f"If {g} Then\nn = 10 \\ z\nEnd If"
    elif form < 0.75:
        body = f"If {g} Then n = 10 \\ z"
    else:
        body = f"If {g} Then\nn = 1\nElse\nn = 10 \\ z\nEnd If"
    return std(f"gfz{seed}/{k:04d}", "", DECL + init + body + "\nMain = n")


if __name__ == "__main__":
    out, n, seed = sys.argv[1], int(sys.argv[2]), int(sys.argv[3])
    rng = random.Random(seed)
    cases = [one(rng, k, seed) for k in range(n)]
    json.dump(cases, open(out, "w", encoding="utf-8"), indent=1)
    print(len(cases), "cases")
