# Usage

A task-oriented guide to pyVBAanalysis. Every example uses the real API; see
[api-reference.md](api-reference.md) for the full surface and
[diagnostics-catalogue.md](diagnostics-catalogue.md) for the codes.

## Analyze one module from a source string

`analyze_module` is the engine entry point. It takes one module's source text and
returns that module's diagnostics. It never raises; on an internal failure it
returns an empty list.

```python
from pyvbaanalysis import analyze_module

source = "Sub S()\n    Dim n As Long\n    n = \"oops\"\nEnd Sub\n"
diagnostics = analyze_module(source)
for diag in diagnostics:
    print(diag.severity.value, diag.code, diag.message)
```

### Reading a result

Each item is a `VbaDiagnostic`:

| Field | Meaning |
| --- | --- |
| `code` | Stable diagnostic code, e.g. `assignment-type-mismatch`. |
| `message` | Human-readable explanation. |
| `severity` | A `DiagnosticSeverity` (`error`, `warning`, `information`). |
| `span` | A `Span` with `start` and `end` character offsets into the source. |
| `spec_reference` | The MS-VBAL or VBE-compiler basis, when one is recorded. |

To turn an offset into a 1-based line and column, use the `line_col` helper:

```python
from pyvbaanalysis import line_col

line, column = line_col(source, diag.span.start)
```

## Pass options

`AnalyzeModuleOptions` carries the module's name and kind plus the cross-module
context. For a single module you usually only set the name and kind:

```python
from pyvbaanalysis import analyze_module, AnalyzeModuleOptions, ModuleSymbolKind

opts = AnalyzeModuleOptions(module_name="ThisWorkbook", module_kind=ModuleSymbolKind.DOCUMENT)
analyze_module(source, opts)
```

`module_kind` is one of `ModuleSymbolKind.STANDARD`, `CLASS`, `DOCUMENT`, or
`USERFORM`. It changes object-module behavior: what `Me` resolves to, whether
public members are exposed, and the document/object rules. When you do not know
the kind, `STANDARD` is the default.

The `project_*` fields (procedures, class members, type names, visible symbols)
are how cross-module rules see the rest of the project. You rarely set these by
hand; `analyze_project` builds them for you (see below).

## Cross-module and whole-project analysis

Many rules need the rest of the project to be precise (member-not-found resolves
a receiver against its class, the type-name rules resolve project types, and so
on). `analyze_project` indexes every module first, then analyzes each one with
that shared context.

```python
from pyvbaanalysis import analyze_project, ModuleInput, ModuleSymbolKind

modules = [
    ModuleInput("Person", ModuleSymbolKind.CLASS, "Public Sub Save()\nEnd Sub\n"),
    ModuleInput("Module1", ModuleSymbolKind.STANDARD,
                "Sub S()\n    Dim p As Person\n    Set p = New Person\n    p.Delete\nEnd Sub\n"),
]
results = analyze_project(modules)   # dict: module name -> list[VbaDiagnostic]
# results["Module1"] reports member-not-found on p.Delete, resolved against Person.
```

Analyze a subset by name while still indexing the whole project for context:

```python
analyze_project(modules, only=["Module1"])   # names match case-insensitively
```

If you need the per-module options yourself (for a custom pass), build the index
and derive them:

```python
from pyvbaanalysis import build_project_index, analyze_module_options_for, ModuleSymbolKind

index = build_project_index(modules)
opts = analyze_module_options_for(index, "Module1", ModuleSymbolKind.STANDARD)
```

## Conditional compilation

`#If` / `#Const` directives decide which code is live. The defaults are `VBA7` and
`Win64` true, `Win32` and `Mac` false; an undefined `#Const` is treated as live so
nothing is missed. To set your own baseline, pass a `ConditionalCompilationEnvironment`:

```python
from pyvbaanalysis import analyze_project, analyze_workbook, ConditionalCompilationEnvironment

env = ConditionalCompilationEnvironment(
    compiler_constants={"Win64": False, "Mac": True},
    project_constants={"DebugMode": True},
)
analyze_project(modules, conditional_compilation=env)
analyze_workbook("Book.xlsm", conditional_compilation=env)
```

A `#Const` set to `False` makes its branch inactive, so diagnostics inside it are
not reported. The keyword is accepted by `analyze_project`, `analyze_workbook`, and
`analyze_loose_file` / `analyze_loose_files`; a per-module `ModuleInput.conditional_compilation`
overrides it for that module.

## Analyze files on disk

### Loose export files (.bas / .cls / .frm)

The reader strips the VBE export header (the `VERSION ... CLASS` and
`Begin {GUID} ... End` designer blocks), infers the module kind from the
extension and header, and derives the module name from the `VB_Name` attribute.

```python
from pyvbaanalysis import analyze_loose_file, analyze_loose_files

analyze_loose_file("Widget.cls")                       # one file
analyze_loose_files(["Module1.bas", "Widget.cls"])     # several, as one project
```

### Excel workbooks (.xlsm / .xlsb / .xlam / .xls)

The workbook reader reads VBA directly out of an Excel file via pyOpenVBA (the one
runtime dependency). pyOpenVBA is imported lazily, so `import pyvbaanalysis` stays
light. The analyzer targets Excel VBA; Word and PowerPoint are out of scope.

```python
from pyvbaanalysis import analyze_workbook

analyze_workbook("Book.xlsm")                  # dict: module name -> diagnostics
analyze_workbook("Book.xlsm", only=["Sheet1"]) # one module by name, full context
```

A path that is not an Excel workbook, or a container with no readable VBA, raises
`WorkbookReadError`.

The reader loads the workbook and its VBA into memory and does not bound the input
size, so impose your own limit (for example a maximum file size) before pointing it
at untrusted files.

## Adjust or silence diagnostics

`severity_overrides` maps a code to `"off"`, `"information"`, `"warning"`, or
`"error"`. The allowed values per code are constrained by policy (some codes can
be downgraded but not turned off). It is accepted by `analyze_project` and the
reader functions, and lives on `AnalyzeModuleOptions` for `analyze_module`.

```python
analyze_project(modules, severity_overrides={"option-explicit-missing": "off"})
```

An invalid override (an unknown code, or a value a code does not allow) is silently
ignored during analysis. Call `validate_severity_overrides` to catch a typo before
it quietly does nothing; it raises `ValueError` with the offending entries.

## Command line

```
pyvbaanalysis path/to/Module1.bas
pyvbaanalysis ./exported_modules --only Sheet1
pyvbaanalysis Book.xlsm --format json
```

A path may be a loose file, a folder of loose files (analyzed together as one
project), or an Excel workbook. `pyvbaanalysis --version` prints the version, and
`python -m pyvbaanalysis` is equivalent to the `pyvbaanalysis` command.

Flags:

| Flag | Effect |
| --- | --- |
| `--only NAME` | Report only the named module(s); repeatable. Project context still uses every module. |
| `--severity CODE=LEVEL` | Override a code's severity (`off`/`information`/`warning`/`error`); repeatable. An invalid code or value exits 2. |
| `--select CODE` | Report only these codes; repeatable. Codes match case-insensitively; an unknown code exits 2. |
| `--ignore CODE` | Hide these codes from the report; repeatable. Codes match case-insensitively; an unknown code exits 2. |
| `--fail-level LEVEL` | Exit non-zero only when a diagnostic at or above `error`/`warning`/`information` is reported (default `information`, meaning any). |
| `--format text\|json` | Output format (default `text`). |

Exit codes: `0` when nothing is reported at or above the fail level, `1` when
diagnostics are reported or a file cannot be read, and `2` for a usage error.

### JSON output

`--format json` prints a list of projects. The shape is stable as of 1.0.0:

```json
[
  {
    "project": "Book.xlsm",
    "modules": [
      {
        "module": "Sheet1",
        "diagnostics": [
          {
            "code": "assignment-type-mismatch",
            "severity": "error",
            "message": "...",
            "start": 42,
            "end": 47,
            "line": 4,
            "column": 9,
            "spec_reference": "MS-VBAL 5.4.3 / ..."
          }
        ]
      }
    ]
  }
]
```

`start` and `end` are character offsets; `line` and `column` are 1-based.
