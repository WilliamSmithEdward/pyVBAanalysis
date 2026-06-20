# pyVBAanalysis

Static analysis for Excel VBA. It reads your macros and reports likely bugs and the
errors the VBA compiler would catch, without opening Excel or running any code.

Point it at a workbook or a set of exported module files, and it returns the
problems it finds, each with the exact line and a plain explanation.

## What it checks

It looks for more than a hundred kinds of problem, including:

- Type errors, such as assigning a string to a `Long` or passing the wrong type to
  a procedure.
- Undeclared variables and calls to procedures or members that do not exist.
- Code the VBA compiler rejects: duplicate declarations, a missing `Option
  Explicit`, malformed statements, or a `Declare` that lacks `PtrSafe` on 64-bit
  Office.
- Likely run-time failures, such as dividing by a constant zero or a type mismatch
  from a bad conversion.

It only reports a problem when it can prove one, and stays quiet otherwise, so the
output does not bury you in false alarms.

## Install

```
pip install pyvbaanalysis
```

Python 3.10 or later. Nothing else to set up.

## Use it from Python

Analyze a workbook:

```python
from pyvbaanalysis import analyze_workbook

for module, problems in analyze_workbook("Budget.xlsm").items():
    for p in problems:
        print(module, p.severity.value, p.code, p.message)
```

Analyze a single module's source:

```python
from pyvbaanalysis import analyze_module

source = "Sub Test()\n    Dim n As Long\n    n = \"oops\"\nEnd Sub\n"
for p in analyze_module(source):
    print(p.code, p.message)
```

Each result has a `code`, a `message`, a `severity` (`error`, `warning`, or
`information`), and a `span` giving the character offsets in the source.

Analyze several exported files together, so references between them resolve:

```python
from pyvbaanalysis import analyze_loose_files

analyze_loose_files(["Module1.bas", "Sheet1.cls", "UserForm1.frm"])
```

## Use it from the command line

```
pyvbaanalysis Budget.xlsm
pyvbaanalysis ./exported_modules --format json
pyvbaanalysis Budget.xlsm --only Sheet1
```

A path can be a workbook, a folder of exported `.bas` / `.cls` / `.frm` files, or a
single file. The command exits 1 when it finds problems and 0 when the code is
clean, so it drops into a CI check.

## Scope

This analyzes Excel VBA. It does not run macros and does not need Excel installed.
Word and PowerPoint are not supported.

## Documentation

- [Usage guide](docs/usage.md)
- [API reference](docs/api-reference.md)
- [Diagnostic catalogue](docs/diagnostics-catalogue.md)

## License

MIT. See [LICENSE](LICENSE).
