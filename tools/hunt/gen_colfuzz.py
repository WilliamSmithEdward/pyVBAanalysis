"""Name collisions: one name declared twice, as two kinds drawn from a pool
(module Dim, Public Dim, Const, Sub, Function, Type, Enum, Enum member,
Declare, procedure Dim, Static, parameter, label), in one or two modules,
then used from Main. VBA either compiles it or refuses it; the cases
measure which.

    python gen_colfuzz.py CASES.json N SEED
"""

import json
import random
import sys

HEAD = "Option Explicit\r\n"
MODULE_KINDS = ["dim", "pubdim", "const", "pubconst", "sub", "func", "pubfunc", "type", "enum", "enummember", "declare"]
LOCAL_KINDS = ["localdim", "localconst", "static", "param", "label"]
NAME = "Zq"


def module_decl(kind: str, name: str) -> tuple[str, str]:
    """(declarations section text, procedures section text)"""
    return {
        "dim": (f"Private {name} As Long\r\n", ""),
        "pubdim": (f"Public {name} As Long\r\n", ""),
        "const": (f"Private Const {name} As Long = 3\r\n", ""),
        "pubconst": (f"Public Const {name} As Long = 3\r\n", ""),
        "sub": ("", f"Private Sub {name}()\r\nEnd Sub\r\n"),
        "func": ("", f"Private Function {name}() As Long\r\n{name} = 2\r\nEnd Function\r\n"),
        "pubfunc": ("", f"Public Function {name}() As Long\r\n{name} = 2\r\nEnd Function\r\n"),
        "type": (f"Private Type {name}\r\n    v As Long\r\nEnd Type\r\n", ""),
        "enum": (f"Private Enum {name}\r\n    {name}A = 1\r\nEnd Enum\r\n", ""),
        "enummember": (f"Private Enum E{name}\r\n    {name} = 1\r\nEnd Enum\r\n", ""),
        "declare": (f"Private Declare PtrSafe Function {name} Lib \"kernel32\" Alias \"GetTickCount\" () As Long\r\n", ""),
    }[kind]


def local_decl(kind: str, name: str) -> tuple[str, str]:
    """(parameter list text, body lines)"""
    return {
        "localdim": ("", f"Dim {name} As Long\r\n{name} = 1\r\n"),
        "localconst": ("", f"Const {name} As Long = 4\r\n"),
        "static": ("", f"Static {name} As Long\r\n"),
        "param": (f"ByVal {name} As Long", ""),
        "label": ("", f"{name}:\r\n"),
    }[kind]


def one(rng: random.Random, k: int, seed: int) -> dict:
    a = rng.choice(MODULE_KINDS)
    b = rng.choice(MODULE_KINDS + LOCAL_KINDS)
    two_modules = rng.random() < 0.35 and b in MODULE_KINDS
    d1, p1 = module_decl(a, NAME)
    use = rng.choice(["Main = 1", f"Main = {NAME}", f"Main = TypeName({NAME})"])
    params, body = "", ""
    d2 = p2 = ""
    if b in LOCAL_KINDS:
        params, body = local_decl(b, NAME)
        helper = f"Private Function Helper({params}) As Long\r\n{body}Helper = 1\r\nEnd Function\r\n"
        p1 += helper
    else:
        d2, p2 = module_decl(b, NAME)
    main = f"Public Function Main() As Variant\r\n    {use}\r\nEnd Function\r\n"
    mods = []
    if two_modules:
        mods.append({"name": "Module1", "type": "standard", "source": HEAD + d1 + main + p1})
        mods.append({"name": "Module2", "type": "standard", "source": HEAD + d2 + p2})
    else:
        mods.append({"name": "Module1", "type": "standard", "source": HEAD + d1 + d2 + main + p1 + p2})
    return {"label": f"clz{seed}/{k:04d}-{a}-{b}{'-2mod' if two_modules else ''}", "run": "Module1.Main", "modules": mods}


if __name__ == "__main__":
    out, n, seed = sys.argv[1], int(sys.argv[2]), int(sys.argv[3])
    rng = random.Random(seed)
    cases = [one(rng, k, seed) for k in range(n)]
    json.dump(cases, open(out, "w", encoding="utf-8"), indent=1)
    print(len(cases), "cases")
