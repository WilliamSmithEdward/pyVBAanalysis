"""Record the analyzeModule calls upstream's own test suite makes, for replay.py.

The recording happens inside a pin (tools/pin_analyzer.py). Hooks are appended to
two upstream files, upstream's vitest suite runs, and `git checkout` puts both
files back, so the pin is pristine again before anything else reads it:
tools/vendor_data.py reuses an existing pin.

Each call is appended to a JSONL file as {simple, keys, source, opts, project,
unreplayable, result}. A call made with project context also records the module
inputs of the ProjectIndex its options came from and, for every project option,
the index method and arguments that produced it, so the port can rebuild the
same index and ask its own methods (replay.py).

The suite needs upstream's node_modules, borrowed from the sibling xlide_vscode
checkout through one link PER ENTRY inside a node_modules folder the pin owns.
One link for the whole folder is not safe: Vite writes caches beside the nearest
node_modules (`.vite-temp` for the bundled config), and a whole-folder link sent
those writes into the sibling checkout. Its entries are compared before and
after every run, and a run that changed them fails.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import stat
import subprocess
import sys
from collections import Counter
from pathlib import Path

ANALYZE_MODULE = "src/analyzer/diagnostics/analyzeModule.ts"
PROJECT_INDEX = "src/analyzer/symbols/projectIndex.ts"
CONFIG = "vitest.record.config.mts"
CACHE = ".vite-record-cache"

# ProjectIndex methods whose results the recorder stamps, with the port's names
# for them, which replay.py calls to rebuild an option.
STAMPED_METHODS: dict[str, str] = {
    "visibleIdentifierNames": "visible_identifier_names",
    "visibleProcedureNames": "visible_procedure_names",
    "visibleNonTypeNames": "visible_non_type_names",
    "procedureSignatures": "procedure_signatures",
    "projectMemberSurfaces": "project_member_surfaces",
    "projectClassMembers": "project_class_members",
    "visibleTypeNames": "visible_type_names",
    "visibleIdentifierSymbols": "visible_identifier_symbols",
    "visibleExternalIntegerConstantExpressions": "visible_external_integer_constant_expressions",
    "stringLiteralWords": "string_literal_words",
    "implementedInterfaceNames": "implemented_interface_names",
}

# The exported function is renamed and a recording wrapper takes its name, so the
# hook depends on one declaration line rather than on the function's body.
_EXPORT = "export function analyzeModule("
_RENAMED = "function __differentialAnalyzeModule("

# Vite creates these beside the nearest node_modules. They are never linked, so
# a run creates them in the pin's own folder.
_UNLINKED = frozenset({".vite", ".vite-temp", ".vitest", ".cache"})

_ANALYZE_MODULE_HOOK = """

// DIFFERENTIAL RECORDER, appended by pyVBAanalysis tools/differential/recorder.py
// and removed with `git checkout` after the run. Never part of upstream.
// analyzeModule is renamed above and wrapped here, so every call, from a test or
// from inside the analyzer, is appended to the JSONL file XLIDE_RECORD names.
import { appendFileSync as __recordAppend } from 'node:fs';
import { getExcelObjectModel as __recordExcelModel } from '../host/excelObjectModel';
import { hostObjectModelForToken as __recordModelForToken } from '../host/hostRegistry';

export function analyzeModule(source: string, opts: AnalyzeModuleOptions = {}): VbaDiagnostic[] {
\tconst result = __differentialAnalyzeModule(source, opts);
\t__recordAnalysis(source, opts, result);
\treturn result;
}

// Options that carry project context. A call without any is replayed with its
// options as they are; a call with some is replayed by rebuilding its index.
const __RECORD_PROJECT_KEYS = [
\t'knownProcedures', 'knownIdentifiers', 'projectProcedures', 'projectClassMembers',
\t'projectTypes', 'implementedInterfaces', 'projectVisibleSymbols', 'knownNonTypeNames',
\t'projectIntegerConstants', 'projectStringLiteralWords', 'hostModel', 'parsedModule',
\t'walkProcedureFilter',
];

function __recordAnalysis(source: string, opts: AnalyzeModuleOptions, result: VbaDiagnostic[]): void {
\tconst path = process.env.XLIDE_RECORD;
\tif (!path) {
\t\treturn;
\t}
\tconst record = opts as Record<string, unknown>;
\tconst present = Object.keys(record).filter((key) => record[key] !== undefined);
\tconst simple = present.every((key) => !__RECORD_PROJECT_KEYS.includes(key));
\tlet project: unknown;
\tlet unreplayable: string | undefined;
\tif (!simple) {
\t\ttype Stamp = { index: unknown; method: string; args: unknown[] };
\t\tconst derived: Record<string, { method: string; args: unknown[] }> = {};
\t\tlet index: unknown;
\t\tfor (const key of present.filter((k) => __RECORD_PROJECT_KEYS.includes(k))) {
\t\t\tif (key === 'parsedModule' || key === 'hostModel') {
\t\t\t\tcontinue;
\t\t\t}
\t\t\tif (key === 'walkProcedureFilter') {
\t\t\t\tunreplayable = 'walkProcedureFilter';
\t\t\t\tbreak;
\t\t\t}
\t\t\tconst value = record[key];
\t\t\tconst stamp = (value as { __recordSource?: Stamp } | undefined)?.__recordSource;
\t\t\tif (!stamp) {
\t\t\t\t// A hand-built set of names or map of constants records as it is.
\t\t\t\tif (value instanceof Set && [...value].every((item) => typeof item === 'string')) {
\t\t\t\t\tderived[key] = { method: '__literal', args: [[...value]] };
\t\t\t\t\tcontinue;
\t\t\t\t}
\t\t\t\tif (value instanceof Map && [...value.keys()].every((item) => typeof item === 'string')
\t\t\t\t\t&& [...value.values()].every((item) => item === undefined || typeof item === 'string')) {
\t\t\t\t\tderived[key] = { method: '__literal', args: [Object.fromEntries(value)] };
\t\t\t\t\tcontinue;
\t\t\t\t}
\t\t\t\tunreplayable = `unstamped ${key}`;
\t\t\t\tbreak;
\t\t\t}
\t\t\tif (index !== undefined && stamp.index !== index) {
\t\t\t\tunreplayable = 'two indexes';
\t\t\t\tbreak;
\t\t\t}
\t\t\tindex = stamp.index;
\t\t\tderived[key] = { method: stamp.method, args: stamp.args };
\t\t}
\t\tlet hostModel: unknown;
\t\tif (!unreplayable && opts.hostModel !== undefined) {
\t\t\tconst token = opts.hostModel === __recordExcelModel()
\t\t\t\t? 'excel'
\t\t\t\t: ['word', 'powerpoint', 'access', 'vb6'].find((t) => __recordModelForToken(t) === opts.hostModel);
\t\t\tif (token) {
\t\t\t\thostModel = { token };
\t\t\t} else if (JSON.stringify(opts.hostModel).length < 2_000_000) {
\t\t\t\thostModel = { model: opts.hostModel };
\t\t\t} else {
\t\t\t\tunreplayable = 'large hostModel';
\t\t\t}
\t\t}
\t\tif (!unreplayable) {
\t\t\tconst holder = index as { __recordedInputs?: Map<string, unknown>; options?: unknown } | undefined;
\t\t\tproject = {
\t\t\t\tinputs: [...(holder?.__recordedInputs?.values() ?? [])],
\t\t\t\tindexOptions: holder?.options,
\t\t\t\tderived,
\t\t\t\thostModel,
\t\t\t};
\t\t}
\t}
\tconst plainOpts = Object.fromEntries(
\t\tpresent.filter((key) => !__RECORD_PROJECT_KEYS.includes(key)).map((key) => [key, record[key]]),
\t);
\ttry {
\t\t__recordAppend(path, JSON.stringify({
\t\t\tsimple,
\t\t\tkeys: present,
\t\t\tsource,
\t\t\topts: plainOpts,
\t\t\tproject,
\t\t\tunreplayable,
\t\t\tresult: result.map((d) => ({ code: d.code, message: d.message, start: d.span.start, end: d.span.end })),
\t\t}) + '\\n');
\t} catch {
\t\t// A record that cannot be written is skipped, never a test failure.
\t}
}
"""

_PROJECT_INDEX_HOOK = """

// DIFFERENTIAL RECORDER, appended by pyVBAanalysis tools/differential/recorder.py
// and removed with `git checkout` after the run. Never part of upstream.
// Keeps each index's module inputs, and stamps each query result with the index,
// method and arguments that made it, so a recorded call can be rebuilt.
{
\tconst proto = ProjectIndex.prototype as unknown as Record<string, (...args: unknown[]) => unknown>;
\tconst required = (name: string): ((...args: unknown[]) => unknown) => {
\t\tconst original = proto[name];
\t\tif (typeof original !== 'function') {
\t\t\tthrow new Error(`differential recorder: ProjectIndex has no method ${name}`);
\t\t}
\t\treturn original;
\t};
\ttype Recorded = { __recordedInputs?: Map<string, unknown> };
\tconst setModule = required('setModule');
\tproto.setModule = function (this: Recorded, ...args: unknown[]): unknown {
\t\tconst result = setModule.apply(this, args);
\t\tconst input = args[0] as { moduleName: string };
\t\t(this.__recordedInputs ??= new Map()).set(input.moduleName.toLowerCase(), input);
\t\treturn result;
\t};
\tconst removeModule = required('removeModule');
\tproto.removeModule = function (this: Recorded, ...args: unknown[]): unknown {
\t\tthis.__recordedInputs?.delete(String(args[0]).toLowerCase());
\t\treturn removeModule.apply(this, args);
\t};
\tfor (const method of [__STAMPED__]) {
\t\tconst original = required(method);
\t\tproto[method] = function (this: unknown, ...args: unknown[]): unknown {
\t\t\tconst result = original.apply(this, args);
\t\t\tif (result && typeof result === 'object') {
\t\t\t\ttry {
\t\t\t\t\tObject.defineProperty(result, '__recordSource', {
\t\t\t\t\t\tvalue: { index: this, method, args }, enumerable: false, configurable: true, writable: true,
\t\t\t\t\t});
\t\t\t\t} catch {
\t\t\t\t\t// A frozen result simply goes unstamped.
\t\t\t\t}
\t\t\t}
\t\t\treturn result;
\t\t};
\t}
}
""".replace("__STAMPED__", ", ".join(f"'{name}'" for name in STAMPED_METHODS))

_VITEST_CONFIG = """// Written by pyVBAanalysis tools/differential/recorder.py and deleted after the run.
import { configDefaults, defineConfig } from 'vitest/config';

export default defineConfig({
    // Every cache stays inside the pin.
    cacheDir: './.vite-record-cache',
    test: {
        environment: 'node',
        include: ['tests/**/*.test.ts'],
        // The browser-bundle test fails once the recorder imports node:fs, which
        // says nothing about the analyzer.
        exclude: [...configDefaults.exclude, 'tests/webBundle.test.ts'],
        testTimeout: 30000,
    },
});
"""


def _git(pin: Path, *args: str) -> str:
    result = subprocess.run(["git", *args], cwd=pin, capture_output=True, text=True, timeout=300)
    if result.returncode != 0:
        raise SystemExit(f"git {' '.join(args)} failed in {pin}: {result.stderr.strip()}")
    return result.stdout


def _declares_method(source: str, name: str) -> bool:
    return re.search(rf"^\s+(?:public\s+)?{name}\s*\(", source, re.MULTILINE) is not None


def patch(pin: Path) -> None:
    """Append the recorder hooks and write the vitest config into a pristine pin."""
    dirty = _git(pin, "status", "--porcelain").splitlines()
    if dirty:
        raise SystemExit(
            f"The pin at {pin} has local changes in {len(dirty)} path(s). Run "
            "`harness.py unpatch`, or re-pin it with tools/pin_analyzer.py --force."
        )
    analyze = pin / ANALYZE_MODULE
    analyze_text = analyze.read_text(encoding="utf-8")
    if analyze_text.count(_EXPORT) != 1:
        raise SystemExit(
            f"{ANALYZE_MODULE} does not declare `{_EXPORT}` exactly once; "
            "the recorder needs updating for this upstream version."
        )
    index = pin / PROJECT_INDEX
    index_text = index.read_text(encoding="utf-8")
    missing = [
        name
        for name in ("setModule", "removeModule", *STAMPED_METHODS)
        if not _declares_method(index_text, name)
    ]
    if not re.search(r"^export class ProjectIndex\b", index_text, re.MULTILINE) or missing:
        raise SystemExit(
            f"{PROJECT_INDEX} lacks class ProjectIndex or its methods {missing}; "
            "the recorder needs updating for this upstream version."
        )
    analyze.write_text(analyze_text.replace(_EXPORT, _RENAMED) + _ANALYZE_MODULE_HOOK, encoding="utf-8", newline="")
    index.write_text(index_text + _PROJECT_INDEX_HOOK, encoding="utf-8", newline="")
    (pin / CONFIG).write_text(_VITEST_CONFIG, encoding="utf-8", newline="")


def unpatch(pin: Path) -> None:
    """Put the pin back as its commit has it: the two files, the config, the
    cache, and the node_modules links. Safe to run on a pin that is not patched."""
    _git(pin, "checkout", "--", ANALYZE_MODULE, PROJECT_INDEX)
    (pin / CONFIG).unlink(missing_ok=True)
    if (pin / CACHE).is_dir():
        _remove_tree(pin / CACHE)
    unlink_node_modules(pin)
    leftover = _git(pin, "status", "--porcelain", "--ignored").splitlines()
    if leftover:
        print(f"note: the pin still differs from its commit in {len(leftover)} path(s):", file=sys.stderr)
        for line in leftover[:10]:
            print(f"  {line}", file=sys.stderr)


# -- links -------------------------------------------------------------------


def _is_link(path: Path) -> bool:
    """A symlink, or a Windows directory junction, which os.path.islink does not
    count as one."""
    if path.is_symlink():
        return True
    isjunction = getattr(os.path, "isjunction", None)  # Python 3.12 and later
    if isjunction is not None:
        return bool(isjunction(path))
    try:
        reparse_tag = getattr(os.lstat(path), "st_reparse_tag", 0)
    except FileNotFoundError:
        return False
    return reparse_tag == getattr(stat, "IO_REPARSE_TAG_MOUNT_POINT", -1)


def _remove_link(path: Path) -> None:
    """Remove a link and never anything it points at."""
    if os.name == "nt":
        # RemoveDirectoryW, the call [System.IO.Directory]::Delete(path, $false)
        # makes: it deletes a junction itself, whatever its target holds.
        os.rmdir(path)
    else:
        os.unlink(path)


def _make_link(link: Path, target: Path) -> None:
    if os.name == "nt":
        # A junction needs no privilege, where a directory symlink does.
        subprocess.run(
            ["cmd", "/c", "mklink", "/J", str(link), str(target)],
            check=True,
            capture_output=True,
            timeout=60,
        )
    else:
        os.symlink(target, link, target_is_directory=True)


def _remove_tree(path: Path) -> None:
    """Delete a real directory tree. A link inside it is removed as a link, so
    nothing it points at is ever visited."""
    for entry in os.scandir(path):
        child = Path(entry.path)
        if _is_link(child):
            _remove_link(child)
        elif entry.is_dir(follow_symlinks=False):
            _remove_tree(child)
        else:
            child.unlink()
    path.rmdir()


def link_node_modules(pin: Path, source: Path) -> None:
    """Give the pin a node_modules of its own, holding one link per package of
    ``source``."""
    folder = pin / "node_modules"
    if not (source / "vitest").is_dir():
        raise SystemExit(f"No vitest in {source}; install the XLIDE checkout's dependencies there.")
    folder.mkdir()
    for entry in sorted(os.scandir(source), key=lambda one: one.name):
        if entry.name in _UNLINKED or not entry.is_dir():
            continue
        _make_link(folder / entry.name, Path(entry.path))


def unlink_node_modules(pin: Path) -> None:
    """Remove the pin's node_modules when it is the recorder's. Each linked package
    goes as a link, never through it, and the caches a run created inside the
    folder go with it. A folder of installed packages is left alone."""
    folder = pin / "node_modules"
    if _is_link(folder):
        # A link to a whole folder, as earlier runs made by hand: the link alone goes.
        _remove_link(folder)
        return
    if not folder.is_dir():
        return
    installed = [
        entry.name
        for entry in os.scandir(folder)
        if entry.name not in _UNLINKED and not _is_link(Path(entry.path))
    ]
    if installed:
        print(f"note: {folder} holds packages of its own ({', '.join(installed[:3])}, ...); left alone.", file=sys.stderr)
        return
    _remove_tree(folder)


def _snapshot(folder: Path) -> dict[str, int]:
    """The folder's own modification time and its entries', to show that a run
    left it alone."""
    entries = {entry.name: entry.stat(follow_symlinks=False).st_mtime_ns for entry in os.scandir(folder)}
    entries["."] = folder.stat().st_mtime_ns
    return entries


# -- the run -----------------------------------------------------------------


def record(pin: Path, out: Path, node_modules: Path, timeout: int) -> None:
    """Record every analyzeModule call of upstream's suite into ``out``."""
    node = shutil.which("node")
    if node is None:
        raise SystemExit("node is not on PATH.")
    own = pin / "node_modules"
    if own.exists() or _is_link(own):
        raise SystemExit(f"{own} already exists; run `harness.py unpatch` first.")
    before = _snapshot(node_modules)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.unlink(missing_ok=True)
    patch(pin)
    try:
        link_node_modules(pin, node_modules)
        print("==> upstream's test suite, recording", file=sys.stderr)
        result = subprocess.run(
            [node, str(pin / "node_modules" / "vitest" / "vitest.mjs"), "run", "--config", CONFIG, "--reporter=dot"],
            cwd=pin,
            env=dict(os.environ, XLIDE_RECORD=str(out.resolve())),
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
        )
        # The summary, or on a failure the report of it; the dot reporter's
        # progress rows carry no letters and are dropped.
        worded = [line for line in (result.stdout + result.stderr).splitlines() if any(c.isalpha() for c in line)]
        for line in worded[-(40 if result.returncode else 5):]:
            print(line.encode("ascii", "replace").decode("ascii"), file=sys.stderr)
        if result.returncode != 0:
            print(
                f"note: vitest exited {result.returncode}; the calls of the tests that ran are recorded",
                file=sys.stderr,
            )
    finally:
        unpatch(pin)
        after = _snapshot(node_modules)
        if after != before:
            changed = sorted(name for name in set(before) | set(after) if before.get(name) != after.get(name))
            raise SystemExit(f"{node_modules} changed during the run, in {changed[:10]}. Check it by hand.")
    summarize(out)


def summarize(calls: Path) -> None:
    rows = [json.loads(line) for line in calls.read_text(encoding="utf-8").splitlines() if line]
    standalone = sum(1 for row in rows if row["simple"])
    project = sum(1 for row in rows if not row["simple"] and row.get("project") is not None)
    reasons = Counter(row.get("unreplayable") for row in rows if not row["simple"] and row.get("project") is None)
    print(f"recorded {len(rows)} calls: {standalone} standalone, {project} with project context")
    if reasons:
        print(f"not replayable: {dict(reasons)}")
