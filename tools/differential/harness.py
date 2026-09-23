"""Compare the port with the upstream XLIDE analyzer it mirrors.

The upstream side runs from a pin (tools/pin_analyzer.py): by default the pin of
the commit the vendored data came from, as manifest.json records it, so both
sides read the same data. Run with the dev environment's Python:

    python tools/differential/harness.py record      # record upstream's own test calls
    python tools/differential/harness.py replay      # replay them through the port
    python tools/differential/harness.py corpus      # the oracle corpus through both
    python tools/differential/harness.py cases OUT.json PATH ...   # projects from files
    python tools/differential/harness.py projects OUT.json         # those projects through both
    python tools/differential/harness.py unpatch     # restore a pin a run left patched

replay, corpus and projects exit 1 when anything differs. Outputs go under
artifacts/differential/<commit>/, which is not in source control. See "Checking a
sync against upstream" in CONTRIBUTING.md.
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from collections.abc import Callable
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Any

_HERE = Path(__file__).resolve().parent
_ROOT = _HERE.parent.parent
# After this folder, so the working tree's package is the one imported.
sys.path.insert(1, str(_ROOT))

import cases  # noqa: E402
import probes  # noqa: E402
import recorder  # noqa: E402
import replay  # noqa: E402

_MARKER = Path("src") / "analyzer" / "diagnostics" / "ruleMetadata.ts"

Records = list[dict[str, Any]]


def _vendored_commit() -> str:
    manifest = json.loads((_ROOT / "pyvbaanalysis" / "data" / "manifest.json").read_text(encoding="utf-8"))
    return str(manifest["xlideCommit"])


def _pin(given: str | None) -> Path:
    """The pin to run upstream from, checked to be one, with a warning when its
    commit is not the one the vendored data came from."""
    vendored = _vendored_commit()
    pin = Path(given).resolve() if given else _ROOT / "artifacts" / "analyzer-pin" / vendored[:7]
    if not (pin / _MARKER).is_file():
        raise SystemExit(
            f"No XLIDE pin at {pin}. Make one with `python tools/vendor_data.py --ref <tag>` "
            "or tools/pin_analyzer.py, or pass --pin."
        )
    head = subprocess.run(
        ["git", "rev-parse", "HEAD"], cwd=pin, capture_output=True, text=True, timeout=60, check=True
    ).stdout.strip()
    if head != vendored:
        print(
            f"warning: the pin is at {head[:7]} but the vendored data came from {vendored[:7]}; "
            "a difference may be the data rather than the code.",
            file=sys.stderr,
        )
    return pin


def _outputs(pin: Path) -> Path:
    head = subprocess.run(
        ["git", "rev-parse", "--short=7", "HEAD"], cwd=pin, capture_output=True, text=True, timeout=60, check=True
    ).stdout.strip()
    return _ROOT / "artifacts" / "differential" / head


def _record(args: argparse.Namespace) -> int:
    pin = _pin(args.pin)
    out = Path(args.out).resolve() if args.out else _outputs(pin) / "calls.jsonl"
    recorder.record(pin, out, Path(args.node_modules).resolve(), args.timeout)
    print(f"calls in {out}")
    return 0


def _replay(args: argparse.Namespace) -> int:
    calls = Path(args.calls) if args.calls else _outputs(_pin(args.pin)) / "calls.jsonl"
    if not calls.is_file():
        raise SystemExit(f"No recorded calls at {calls}; run `harness.py record` first.")
    return replay.replay(calls, args.show, args.code)


def _both(script: str, pin: Path, inputs: Path, timeout: int, port: Callable[[Path], Records]) -> tuple[Records, Records]:
    """Upstream in a thread, since it only waits on its process, while the port
    runs here: the two halves take about as long as each other."""
    with ThreadPoolExecutor(max_workers=1) as pool:
        upstream_future = pool.submit(probes.upstream, script, pin, inputs, timeout)
        port_records = port(inputs)
        return upstream_future.result(), port_records


def _corpus(args: argparse.Namespace) -> int:
    pin = _pin(args.pin)
    corpus = Path(args.corpus).resolve() if args.corpus else pin / "syntax_corpus" / "oracle" / "vbe_oracle_cases.json"
    out = _outputs(pin)
    upstream_records, port_records = _both("corpus_messages.mjs", pin, corpus, args.timeout, probes.port_corpus)
    probes.write_jsonl(out / "corpus_upstream.jsonl", upstream_records)
    probes.write_jsonl(out / "corpus_port.jsonl", port_records)
    return probes.compare(upstream_records, port_records, "corpus", args.show)


def _cases(args: argparse.Namespace) -> int:
    markdown_pin = _pin(args.pin) if args.markdown else None
    built = cases.build_cases([Path(p) for p in args.paths], markdown_pin, args.host)
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(built, ensure_ascii=True), encoding="utf-8")
    modules = sum(len(case["modules"]) for case in built)
    print(f"wrote {len(built)} cases with {modules} modules to {out}")
    return 0


def _projects(args: argparse.Namespace) -> int:
    pin = _pin(args.pin)
    case_file = Path(args.cases).resolve()
    out = _outputs(pin)
    upstream_records, port_records = _both("project_probe.mjs", pin, case_file, args.timeout, probes.port_projects)
    probes.write_jsonl(out / f"{case_file.stem}_upstream.jsonl", upstream_records)
    probes.write_jsonl(out / f"{case_file.stem}_port.jsonl", port_records)
    return probes.compare(upstream_records, port_records, "projects", args.show)


def _unpatch(args: argparse.Namespace) -> int:
    recorder.unpatch(_pin(args.pin))
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    commands = parser.add_subparsers(dest="command", required=True)

    def command(name: str, help_text: str) -> argparse.ArgumentParser:
        sub = commands.add_parser(name, help=help_text, description=help_text)
        sub.add_argument("--pin", help="the pin to run upstream from (default: the vendored commit's)")
        return sub

    record = command("record", "record every analyzeModule call of upstream's own test suite")
    record.add_argument("--out", help="the JSONL file to write (default: artifacts/differential/<commit>/calls.jsonl)")
    record.add_argument(
        "--node-modules",
        default=str(_ROOT.parent / "xlide_vscode" / "node_modules"),
        help="the installed packages the suite borrows, read-only (default: the sibling checkout's)",
    )
    record.add_argument("--timeout", type=int, default=1800, help="seconds before the suite is stopped")
    record.set_defaults(run=_record)

    replay_command = command("replay", "replay the recorded calls through the port")
    replay_command.add_argument("--calls", help="the recorded calls (default: the pin's calls.jsonl)")
    replay_command.add_argument("--show", type=int, default=10, help="differing calls to print in full")
    replay_command.add_argument("--code", help="only count calls whose differences include this code")
    replay_command.set_defaults(run=_replay)

    corpus = command("corpus", "run the oracle corpus through both analyzers and compare")
    corpus.add_argument("--corpus", help="the cases file (default: the pin's syntax_corpus/oracle one)")
    corpus.add_argument("--show", type=int, default=20, help="differing modules to print")
    corpus.add_argument("--timeout", type=int, default=1800, help="seconds before upstream is stopped")
    corpus.set_defaults(run=_corpus)

    build = command("cases", "build project cases from Office files and folders of exported modules")
    build.add_argument("out", help="the cases file to write")
    build.add_argument("paths", nargs="*", help="Office files, or folders to search for them and for exports")
    build.add_argument("--markdown", action="store_true", help="add every VBA block of the pin's syntax_corpus")
    build.add_argument(
        "--host", help="the host for exported modules and markdown blocks, which name none (Office files name their own)"
    )
    build.set_defaults(run=_cases)

    projects = command("projects", "run project cases through both analyzers and compare")
    projects.add_argument("cases", help="a cases file from `harness.py cases`")
    projects.add_argument("--show", type=int, default=20, help="differing modules to print")
    projects.add_argument("--timeout", type=int, default=1800, help="seconds before upstream is stopped")
    projects.set_defaults(run=_projects)

    unpatch = command("unpatch", "restore a pin a recording left patched")
    unpatch.set_defaults(run=_unpatch)

    args = parser.parse_args()
    return int(args.run(args))


if __name__ == "__main__":
    sys.exit(main())
