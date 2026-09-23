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

`analyze_loose_file` analyzes one file on its own, so the whole-project checks are
skipped for it (see "Whole project vs a single file" below). Use `analyze_loose_files`
to analyze several files together with shared cross-module context.

### Office macro containers

The container reader reads VBA directly out of an Office file via pyOpenVBA (the one
runtime dependency). pyOpenVBA is imported lazily, so `import pyvbaanalysis` stays
light.

| Host | Extensions |
| --- | --- |
| Excel | `.xlsm`, `.xlsb`, `.xlam`, `.xls` |
| Word | `.docm`, `.dotm`, `.doc` |
| PowerPoint | `.pptm`, `.potm` |
| Access | `.accdb`, `.mdb` (read-only) |

```python
from pyvbaanalysis import analyze_office_file

analyze_office_file("Book.xlsm")                  # dict: module name -> diagnostics
analyze_office_file("Report.docm")                # resolved against Word's model
analyze_office_file("Book.xlsm", only=["Sheet1"]) # one module by name, full context
```

The extension selects the host, so Word code is measured against Word's object model
and never against Excel's. Analyzing a Word module under Excel's model reports
members and constants that are perfectly legal in Word (`Selection.TypeText`,
`ActiveDocument`, `wdOrientPortrait`), which is the false positive this avoids.

The project's reference list is read with its modules. A workbook that references
the Word object library can declare `Dim doc As Word.Document` and use Word's
constants, and those are checked against Word's model. A workbook that names
`Word.Document` without the reference reports `missing-library-reference`: the VBE
refuses that declaration with "User-defined type not defined". To read a container
without analyzing it, `read_office_project(path)` in `pyvbaanalysis.reader` returns
its modules, its host, and the other Office libraries it references.

`analyze_workbook` is the Excel-only form of the same call: it refuses any other
extension and otherwise returns exactly what `analyze_office_file` does.

A path that is not a readable container, or a container with no readable VBA, raises
`WorkbookReadError`. Legacy `.ppt` is not readable: pyOpenVBA reads it as a plain
CFB, but its VBA project sits inside a compressed record, so the extension is
rejected rather than failing later with a parse error
([pyOpenVBA #17](https://github.com/WilliamSmithEdward/pyOpenVBA/issues/17)).

### Choosing the host yourself

When you already have the module text, name the host directly. Absent means Excel,
so existing calls are unchanged; a named host with no model asserts no host
knowledge at all rather than falling back to Excel's.

```python
from pyvbaanalysis import analyze_project

analyze_project(modules, host="word")      # Word's object model
analyze_project(modules, host="vb6")       # VB6: App, Screen, Printer, the intrinsic controls
analyze_project(modules, host="outlook")   # no model yet: stays quiet, never Excel's
analyze_project(modules)                   # Excel, exactly as before
```

A VB6 project is no Office container, so nothing reads one from a file; pass its
modules with `host="vb6"`. The VB6 model offers and describes, and never proves a
member absent.

Name the other Office libraries the project references the same way, in the order
its References dialog lists them. The project's own host wins any name two
libraries share, as it does in VBA:

```python
analyze_project(modules, referenced_hosts=["word"])   # an Excel project referencing Word
analyze_project(modules, referenced_hosts=[])         # known to reference nothing else
analyze_project(modules)                              # reference list unknown
```

`missing-library-reference` only reports against a known list. Left unset, the list
is unknown, as it is for loose `.bas` files, and the check stays silent rather than
guess that a reference is missing.

The reader loads the workbook and its VBA into memory and does not bound the input
size, so impose your own limit (for example a maximum file size) before pointing it
at untrusted files.

## Whole project vs a single file

Three checks need every module to be correct: `undeclared-variable`, `unknown-call`,
and `member-not-found` resolve a name across the whole project, so a symbol declared
in a module the analyzer cannot see would look undefined. To stay false-positive-free,
those checks run only when the analyzed set is the complete project and are skipped
for a partial view:

* `analyze_workbook` and `analyze_loose_files` treat their input as the whole project
  (a workbook holds every module; a set of files is taken as the project).
* `analyze_loose_file` analyzes one file in isolation, so it is partial by default.
  Pass `whole_project=True` if that single file really is the entire project.
* `analyze_project` defaults to whole-project; pass `whole_project=False` for a fragment.
* On the command line, a folder or several files is a whole project, a single targeted
  file is partial automatically, and `--partial-project` forces partial for any input.

Every other check is local or resolves positively (it reports only when it can prove
the problem), so a partial view never turns it into a false positive.

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

### Inline suppression

Suppress diagnostics from within the source with `'@pyvba-ignore` comment directives:

```vba
'@pyvba-ignore-file: option-explicit-missing

Sub Demo()
    Dim a(10 To 1) As Long  '@pyvba-ignore: array-declaration-impossible-bounds
End Sub
```

* `'@pyvba-ignore` suppresses diagnostics on its own line (write it as a trailing comment).
* `'@pyvba-ignore-next-line` suppresses the following line.
* `'@pyvba-ignore-file` suppresses the whole module; place it before the first
  non-comment, non-attribute line.

Each takes an optional `: code1, code2` list (omit it, or write `all`, to suppress every
code); codes match case-insensitively, and a `-- reason` trailer is free text. A
malformed directive (an unknown code, an unknown verb, or a misplaced `-ignore-file`) is
reported as `analysis-suppression-directive` and suppresses nothing. Directives are
single-apostrophe comments; `'''` doc comments and `Rem` comments are not directives, and
the two structural codes (`missing-block-closer`, `unmatched-block-closer`) are not
suppressible this way.

Pass `inline_suppression=False` (library) or `--no-inline-suppression` (CLI) to ignore
every directive and report all diagnostics, for an audit run.

## Command line

```
pyvbaanalysis path/to/Module1.bas
pyvbaanalysis ./exported_modules --only Sheet1
pyvbaanalysis Book.xlsm --format json
```

A path may be a loose file, a folder of loose files (analyzed together as one
project), or an Office macro container, whose extension selects the host model.
`pyvbaanalysis --version` prints the version, and `python -m pyvbaanalysis` is
equivalent to the `pyvbaanalysis` command.

Flags:

| Flag | Effect |
| --- | --- |
| `--only NAME` | Report only the named module(s); repeatable. Project context still uses every module. |
| `--severity CODE=LEVEL` | Override a code's severity (`off`/`information`/`warning`/`error`); repeatable. An invalid code or value exits 2. |
| `--select CODE` | Report only these codes; repeatable. Codes match case-insensitively; an unknown code exits 2. |
| `--ignore CODE` | Hide these codes from the report; repeatable. Codes match case-insensitively; an unknown code exits 2. |
| `--partial-project` | Treat the input as a fragment of a larger project: skip the whole-project checks (`undeclared-variable`, `unknown-call`, `member-not-found`). A single targeted file is treated as partial automatically. |
| `--no-inline-suppression` | Ignore `'@pyvba-ignore` directives in the source and report every diagnostic (an audit run). |
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

## Use in CI

`pyvbaanalysis` exits non-zero when it reports a diagnostic, so it gates a build
directly. A GitHub Actions job that fails on errors and lets warnings through:

```yaml
name: VBA lint
on: [push, pull_request]
jobs:
  vba:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: "3.12"
      - run: pip install pyvbaanalysis
      - run: pyvbaanalysis path/to/vba --fail-level error
```

Point the last step at a folder of exported `.bas` / `.cls` / `.frm` files or at a
workbook. Tune the gate with the flags above: `--fail-level warning` to also fail on
warnings, `--severity option-explicit-missing=off` or `--ignore <code>` to mute a
code, and `--partial-project` when the checked-in files are only a fragment of the
project. Use `--format json` if a later step needs to parse the results.

The dead-code checks (`unused-variable`, `variable-never-read`, `unused-procedure`,
`unreachable-code`) report at `information` level. Without `--fail-level`, the gate
fails on any diagnostic, so a single unused variable fails the run.
