"""Random sheet-state programs: two added sheets, random Name, Visible,
Activate, Select, Protect/Unprotect, cell edits and Delete, then cleanup
that restores the workbook whatever happened.

    python gen_shfuzz.py CASES.json N SEED
"""

import json
import random
import sys

from gen_dtc import std

OPS = [
    'w2.Name = "zzA{k}"', 'w3.Name = "zzA{k}"', 'w2.Name = "zzB{k}"',
    "w2.Visible = xlSheetHidden", "w2.Visible = xlSheetVisible", "w2.Visible = xlSheetVeryHidden",
    "w3.Visible = xlSheetHidden",
    "w2.Activate", "w3.Activate", "w1.Activate",
    "w2.Select", "w3.Select",
    'w2.Protect "pw"', 'w2.Unprotect "pw"', "w2.Protect", "w2.Unprotect",
    'w2.Range("A1").Value = 1', 'w3.Range("B2").Value = "x"',
    'w2.Range("A1").Select',
    "Application.DisplayAlerts = False: w3.Delete: Application.DisplayAlerts = True",
    "n = w2.Index", "n = Worksheets.Count",
    "w2.Move After:=w1", "w3.Move Before:=w2",
]
PRE = ("Dim w1 As Worksheet, w2 As Worksheet, w3 As Worksheet, n As Long, e As Long\n"
       "Set w1 = ActiveSheet\nSet w2 = Worksheets.Add\nSet w3 = Worksheets.Add\n"
       "On Error GoTo H\n")
TAIL = ("\nC:\nOn Error Resume Next\nApplication.DisplayAlerts = False\n"
        'w2.Unprotect "pw": w2.Unprotect\nw2.Visible = xlSheetVisible: w3.Visible = xlSheetVisible\n'
        "w2.Delete: w3.Delete\nw1.Activate\nApplication.DisplayAlerts = True\nOn Error GoTo 0\n"
        "If e <> 0 Then Err.Raise e\nMain = n\nExit Function\nH:\ne = Err.Number\nResume C")


def main(count: int, seed: int) -> list:
    rng = random.Random(seed)
    cases = []
    for k in range(count):
        ops = [rng.choice(OPS).format(k=k) for _ in range(rng.randint(2, 5))]
        cases.append(std(f"shz{seed}/{k:04d}", "", PRE + "\n".join(ops) + TAIL))
    return cases


if __name__ == "__main__":
    out = main(int(sys.argv[2]), int(sys.argv[3]))
    json.dump(out, open(sys.argv[1], "w", encoding="utf-8"), indent=1)
    print(len(out), "cases")
