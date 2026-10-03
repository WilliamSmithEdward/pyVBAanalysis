"""Random nested date functions over date literals near the ends of the
date range: DateAdd, DateSerial, TimeSerial, date arithmetic, Year/
Month/Day, DateDiff, Weekday, CDate of numbers.

    python gen_datefuzz.py CASES.json N SEED
"""

import json
import random
import sys

from gen_dtc import std

DATES = ["#1/1/100#", "#12/31/9999#", "#2/29/2000#", "#1/31/2020#", "#12/30/1899#", "t", "#6/15/5000#"]
INTERVALS = ['"yyyy"', '"q"', '"m"', '"d"', '"ww"', '"h"', '"n"', '"s"', '"y"', '"w"']
DATE_TEXTS = ['"2020-02-30"', '"2020-02-29"', '"1/1/100"', '"12/31/9999"']


def num(rng: random.Random, depth: int) -> str:
    p = rng.random()
    if depth <= 0 or p < 0.4:
        return str(rng.choice([-1, 0, 1, 12, 13, 31, 32, 366, 9999, 10000, -10000, 2958465, 3000000]))
    if p < 0.6:
        return f"{rng.choice(['Year', 'Month', 'Day', 'Weekday'])}({dexpr(rng, depth - 1)})"
    if p < 0.8:
        return f"DateDiff({rng.choice(INTERVALS)}, {dexpr(rng, depth - 1)}, {dexpr(rng, depth - 1)})"
    return f"({num(rng, depth - 1)} {rng.choice(['+', '-', '*'])} {num(rng, depth - 1)})"


def dexpr(rng: random.Random, depth: int) -> str:
    if depth <= 0 or rng.random() < 0.3:
        return rng.choice(DATES)
    f = rng.choice(["add", "add", "serial", "time", "plus", "minus", "cdate", "dateval"])
    if f == "add":
        return f"DateAdd({rng.choice(INTERVALS)}, {num(rng, depth - 1)}, {dexpr(rng, depth - 1)})"
    if f == "serial":
        return f"DateSerial({num(rng, depth - 1)}, {num(rng, depth - 1)}, {num(rng, depth - 1)})"
    if f == "time":
        return f"TimeSerial({num(rng, depth - 1)}, {num(rng, depth - 1)}, 0)"
    if f == "plus":
        return f"({dexpr(rng, depth - 1)} + {num(rng, depth - 1)})"
    if f == "minus":
        return f"({dexpr(rng, depth - 1)} - {num(rng, depth - 1)})"
    if f == "cdate":
        return f"CDate({num(rng, depth - 1)})"
    return f"DateValue({rng.choice(DATE_TEXTS)})"


def main(n: int, seed: int) -> list:
    rng = random.Random(seed)
    cases = []
    for k in range(n):
        e = num(rng, 3) if rng.random() < 0.3 else dexpr(rng, 3)
        body = f"Dim t As Date\nt = #3/15/2023#\nDim r As Variant\nr = {e}\nMain = CStr(r)"
        cases.append(std(f"dtz{seed}/{k:04d}", "", body))
    return cases


if __name__ == "__main__":
    out = main(int(sys.argv[2]), int(sys.argv[3]))
    json.dump(out, open(sys.argv[1], "w", encoding="utf-8"), indent=1)
    print(len(out), "cases")
