"""Random expressions over typed locals, stored into a typed target.

    python gen_exprfuzz.py CASES.json N SEED
"""

import json
import random
import sys

from gen_dtc import std

n, seed = int(sys.argv[2]), int(sys.argv[3])
rng = random.Random(seed)
LOCALS = {
    "i": ("Integer", ["0", "1", "-1", "300", "32767", "-32768", "200"]),
    "l": ("Long", ["0", "2", "-5", "70000", "2147483647", "100000"]),
    "b": ("Byte", ["0", "1", "200", "255"]),
    "d": ("Double", ["0", "0.5", "-2.5", "1E10", "3.7", "1E300"]),
    "c": ("Currency", ["0", "1.25", "922337203685477"]),
    "t": ("Date", ["#1/2/2000#", "#12/31/9999#", "0"]),
    "o": ("Boolean", ["True", "False"]),
    "s": ("String", ['"12"', '"abc"', '""', '"1E3"', '"-4"', '" 7 "']),
    "v": ("Variant", ["Empty", "Null", "5", '"x"', "2.5", "32767"]),
}
BIN = ["+", "-", "*", "/", "\\", "Mod", "^", "&", "=", "<", "And", "Or"]
FUNCS = ["Abs({})", "Int({})", "Fix({})", "CInt({})", "CLng({})", "CDbl({})", "Len({})", "Val({})", "CStr({})",
         "Sgn({})", "-({})", "Not {}"]
TARGETS = ["Integer", "Long", "Double", "String", "Variant", "Byte", "Boolean", "Currency"]


def expr(depth: int) -> str:
    r = rng.random()
    if depth <= 0 or r < 0.3:
        return rng.choice(list(LOCALS)) if rng.random() < 0.8 else rng.choice(["2", "0", "1.5", '"3"', "40000"])
    if r < 0.5:
        return rng.choice(FUNCS).format(expr(depth - 1))
    return f"({expr(depth - 1)} {rng.choice(BIN)} {expr(depth - 1)})"


CASES = []
for k in range(n):
    e = expr(rng.choice([1, 2, 3]))
    used = [name for name in LOCALS if any(tok == name for tok in e.replace("(", " ").replace(")", " ").split())]
    lines = []
    for name in used:
        t, vals = LOCALS[name]
        lines.append(f"Dim {name} As {t}")
        lines.append(f"{name} = {rng.choice(vals)}")
    target = rng.choice(TARGETS)
    lines.append(f"Dim r As {target}")
    lines.append(f"r = {e}")
    lines.append("Main = r")
    CASES.append(std(f"xf{seed}/{k:04d}", "", "\n".join(lines)))

if __name__ == "__main__":
    json.dump(CASES, open(sys.argv[1], "w", encoding="utf-8"), indent=1)
    print(len(CASES), "cases")
