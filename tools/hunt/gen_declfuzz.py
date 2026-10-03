"""Random declarations, meant to be valid: module-level and procedure-level
Dim, Const, Enum, Type, Declare, Event and WithEvents, used by Main.

    python gen_declfuzz.py CASES.json N SEED
"""

import json
import random
import sys

n, seed = int(sys.argv[2]), int(sys.argv[3])
rng = random.Random(seed)
SCALARS = ["Long", "Integer", "Byte", "Double", "Single", "Currency", "Date", "String", "Boolean", "Variant", "LongLong"]


def const_expr(names):
    parts = [str(rng.randint(-5, 50))] + names
    a, b = rng.choice(parts), rng.choice(parts)
    return rng.choice([a, f"{a} + {b}", f"({a}) * 2", f"{a} \\ 2", f"-{a}", f"{a} Mod 7 + 1", f"Abs({a})", f"{a} And 255"])


def module(k):
    std, cls = [], []
    consts, enums = [], []
    for j in range(rng.randint(1, 4)):
        name = f"K{j}"
        t = rng.choice(["", " As Long", " As Integer", " As Double", " As Variant"])
        std.append(f"{rng.choice(['Private', 'Public', ''])} Const {name}{t} = {const_expr(consts)}".strip())
        consts.append(name)
    if rng.random() < 0.6:
        members, mnames = [], []
        for j in range(rng.randint(1, 4)):
            m = f"eM{j}"
            members.append(f"    {m}" + (f" = {const_expr(mnames + consts)}" if rng.random() < 0.6 else ""))
            mnames.append(m)
        std.append(f"{rng.choice(['Private', 'Public'])} Enum En{k}")
        std += members
        std.append("End Enum")
        enums += mnames
    if rng.random() < 0.6:
        std.append("Private Type Rec")
        for j in range(rng.randint(1, 4)):
            t = rng.choice(SCALARS)
            shape = rng.choice(["", f"({rng.randint(0, 4)})", f"(1 To {rng.randint(1, 5)})", "()"]) if t != "String" else rng.choice(["", " * 5"])
            if shape == " * 5":
                std.append(f"    f{j} As String * 5")
            else:
                std.append(f"    f{j}{shape} As {t}")
        std.append("End Type")
    if rng.random() < 0.4:
        std.append(rng.choice([
            'Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long',
            'Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)',
            'Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long',
            'Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal p As LongPtr) As Long',
        ]))
    for j in range(rng.randint(0, 3)):
        t = rng.choice(SCALARS + ["Collection", "Object", "New Collection"])
        shape = rng.choice(["", "()", f"({rng.randint(0, 3)})", f"({rng.choice(consts)} To {rng.choice(consts)} + 3)" if consts else ""])
        if t == "New Collection" and shape:
            shape = ""
        std.append(f"{rng.choice(['Private', 'Public', 'Dim'])} m{j}{shape} As {t}")
    body = []
    for j in range(rng.randint(0, 4)):
        t = rng.choice(SCALARS + ["Collection", "New Collection"])
        shape = rng.choice(["", "()", f"({rng.randint(0, 3)})"]) if t != "New Collection" else ""
        body.append(f"{rng.choice(['Dim', 'Static'])} v{j}{shape} As {t}")
    if rng.random() < 0.3:
        body.append(f"Const L1 As Long = {const_expr(consts)}")
    uses = consts + enums
    body.append("Main = " + (" + ".join(rng.sample(uses, min(len(uses), 3))) if uses else "0"))
    src = "Option Explicit\r\n" + "\r\n".join(std) + "\r\nPublic Function Main() As Variant\r\n" + "".join(f"    {b}\r\n" for b in body) + "End Function\r\n"
    mods = [{"name": "Module1", "type": "standard", "source": src}]
    if rng.random() < 0.3:
        cls = ["Public Event Changed(ByVal v As Long)", "Private WithEvents mC As Class2", "Public Sub Fire()", "    RaiseEvent Changed(1)", "End Sub"]
        mods.append({"name": "Class1", "type": "class", "source": "Option Explicit\r\n" + "\r\n".join(cls) + "\r\n"})
        mods.append({"name": "Class2", "type": "class", "source": "Option Explicit\r\nPublic Event Ping()\r\n"})
    return {"label": f"dz{seed}/{k:04d}", "run": "Module1.Main", "modules": mods}


CASES = [module(k) for k in range(n)]
if __name__ == "__main__":
    json.dump(CASES, open(sys.argv[1], "w", encoding="utf-8"), indent=1)
    print(len(CASES), "cases")
