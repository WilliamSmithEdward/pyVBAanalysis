"""Random chains of Range geometry on literal addresses: Offset, Resize,
Cells, Rows, Columns, Item, EntireRow/EntireColumn, ending in .Address.
Tests the analyzer's range folding against Excel's edges.

    python gen_rgfuzz.py CASES.json N SEED
"""

import json
import random
import sys

from gen_dtc import std

STARTS = ['"A1"', '"B2"', '"C3:E7"', '"XFD1"', '"A1048576"', '"XFC1048575:XFD1048576"', '"D4"',
          '"A1:A3"', '"Z100"', '"B2:C2"']


def num(rng: random.Random, lo: int, hi: int) -> int:
    pick = rng.random()
    if pick < 0.15:
        return rng.choice([0, -1, 1])
    if pick < 0.25:
        return rng.choice([16384, 16385, 1048576, 1048577, -16384])
    return rng.randint(lo, hi)


def op(rng: random.Random) -> str:
    kind = rng.choice(["offset", "offset", "resize", "resize", "cells", "rows", "columns", "item",
                       "entirerow", "entirecolumn", "offset1", "resize1"])
    if kind == "offset":
        return f".Offset({num(rng, -5, 5)}, {num(rng, -5, 5)})"
    if kind == "offset1":
        return f".Offset({num(rng, -5, 5)})"
    if kind == "resize":
        return f".Resize({num(rng, 1, 6)}, {num(rng, 1, 6)})"
    if kind == "resize1":
        return f".Resize({num(rng, 1, 6)})"
    if kind == "cells":
        return f".Cells({num(rng, -2, 6)}, {num(rng, -2, 6)})"
    if kind == "rows":
        return f".Rows({num(rng, 0, 4)})"
    if kind == "columns":
        return f".Columns({num(rng, 0, 4)})"
    if kind == "item":
        return f".Item({num(rng, 0, 8)})"
    return ".EntireRow" if kind == "entirerow" else ".EntireColumn"


def main(n: int, seed: int) -> list:
    rng = random.Random(seed)
    cases = []
    for k in range(n):
        chain = "".join(op(rng) for _ in range(rng.randint(1, 3)))
        expr = f"ActiveSheet.Range({rng.choice(STARTS)}){chain}.Address"
        cases.append(std(f"rgz{seed}/{k:04d}", "", f"Main = {expr}"))
    return cases


if __name__ == "__main__":
    out = main(int(sys.argv[2]), int(sys.argv[3]))
    json.dump(out, open(sys.argv[1], "w", encoding="utf-8"), indent=1)
    print(len(out), "cases")
