"""Random calls to a project procedure: random signatures (ByRef, ByVal,
Optional, ParamArray, typed and array parameters) and random arguments
(literals, locals, expressions, parenthesized locals, array elements, Type
fields, omitted and named), in every call form.

    python gen_callfuzz.py CASES.json N SEED
"""

import json
import random
import sys

from gen_dtc import std

TYPES = ["Integer", "Long", "Double", "String", "Boolean", "Variant", "Date", "Byte", "Currency", "Collection"]
VALS = {
    "Integer": ["5", "-3", "32767"], "Long": ["5", "40000", "-1"], "Double": ["1.5", "0", "1E+20"],
    "String": ['"abc"', '"5"', '""'], "Boolean": ["True", "False"], "Variant": ["5", '"x"', "Null", "Empty"],
    "Date": ["#1/2/2000#"], "Byte": ["200", "0"], "Currency": ["1.25"], "Collection": ["New Collection"],
}
LITS = ['"abc"', '"7"', "5", "40000", "1.5", "True", "#1/2/2000#", "Nothing", "Empty", "Null", "-1", "300"]


def main(n: int, seed: int) -> list:
    out = []
    for k in range(n):
        rng = random.Random(seed * 100000 + k)
        np_ = rng.randint(1, 3)
        params, sig = [], []
        for j in range(np_):
            last = j == np_ - 1
            any_opt = any(q.get("opt") for q in params)
            if last and not any_opt and rng.random() < 0.12:
                params.append({"name": f"p{j}", "kind": "paramarray"})
                sig.append(f"ParamArray p{j}() As Variant")
                continue
            t = rng.choice(TYPES)
            mode = rng.choice(["", "ByRef ", "ByVal "])
            arr = t not in ("Collection",) and rng.random() < 0.1
            opt = not arr and (any_opt or rng.random() < 0.25)
            p = {"name": f"p{j}", "kind": "normal", "type": t, "mode": mode, "arr": arr, "opt": opt}
            if arr:
                mode = "" if mode == "ByVal " else mode
                p["mode"] = mode
                sig.append(f"{mode}p{j}() As {t}")
            elif opt:
                dflt = "" if t in ("Collection", "Variant") else f" = {VALS[t][0]}" if t != "Collection" else ""
                sig.append(f"Optional {mode}p{j} As {t}{dflt}")
            else:
                sig.append(f"{mode}p{j} As {t}")
            params.append(p)
        is_fn = rng.random() < 0.5
        callee = (f"Private Function Callee({', '.join(sig)}) As Long\r\n    Callee = 1\r\nEnd Function\r\n" if is_fn
                  else f"Private Sub Callee({', '.join(sig)})\r\nEnd Sub\r\n")
        lines, args = [], []
        nargs = len(params) + rng.choice([0, 0, 0, 0, -1, 1])
        named = rng.random() < 0.15
        for j in range(max(0, nargs)):
            p = params[j] if j < len(params) else None
            r = rng.random()
            if p is not None and p.get("opt") and r < 0.1:
                args.append("")
                continue
            if p is not None and p.get("arr") and r < 0.7:
                lines.append(f"Dim a{j}() As {rng.choice([p['type'], p['type'], 'Long', 'Variant'])}")
                lines.append(f"ReDim a{j}(1)")
                args.append(f"a{j}")
                continue
            r = rng.random()
            if r < 0.3:
                args.append(rng.choice(LITS))
            else:
                t = rng.choice(TYPES + [p["type"]] * 3 if p and p.get("type") else TYPES)
                obj = t == "Collection"
                lines.append(f"Dim v{j} As {t}")
                lines.append(f"{'Set ' if obj else ''}v{j} = {rng.choice(VALS[t])}")
                form = rng.random()
                if form < 0.6 or obj:
                    args.append(f"v{j}")
                elif form < 0.8:
                    args.append(f"(v{j})")
                else:
                    args.append(f"v{j} + 0" if t not in ("String",) else f'v{j} & ""')
        if named:
            args = [f"p{j}:={a}" for j, a in enumerate(args) if a and j < len(params)
                    and params[j]["kind"] != "paramarray"]
        arglist = ", ".join(args)
        form = rng.random()
        if is_fn and form < 0.4:
            call = f"Main = Callee({arglist})"
        elif form < 0.7:
            call = f"Call Callee({arglist})" if arglist else "Call Callee"
        else:
            call = f"Callee {arglist}".rstrip()
        lines.append(call)
        if not call.startswith("Main ="):
            lines.append("Main = 1")
        out.append(std(f"cz{seed}/{k:04d}", "", "\n".join(lines), callee))
    return out


if __name__ == "__main__":
    cases = main(int(sys.argv[2]), int(sys.argv[3]))
    json.dump(cases, open(sys.argv[1], "w", encoding="utf-8"), indent=1)
    print(len(cases), "cases")
