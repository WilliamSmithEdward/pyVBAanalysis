"""Ask the real Office host about each hunt case: does the project compile, and
what does running its entry point do.

    python oracle_run.py CASES.json RESULTS.json [--only LABEL_SUBSTRING] [--limit N]

A case is {label, host?, modules: [{name, type, source}], run?: "Module.Proc"}.
The host (excel by default, or word, powerpoint, access) picks the session.
Only standard and class modules can be injected; a case with any other module
type is recorded as unsupported. Results are keyed by label and written after
every case, so a rerun resumes where the last one stopped.
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from collections import defaultdict
from pathlib import Path

from pyvbaharness import AccessSession, ExcelSession, PowerPointSession, WordSession
from pyvbaharness.results import SessionLockHeld
from pyvbaharness.session import HarnessConfig

SESSIONS = {
    "excel": ExcelSession,
    "word": WordSession,
    "powerpoint": PowerPointSession,
    "access": AccessSession,
}


def _load(path: Path) -> dict:
    if path.is_file():
        return json.loads(path.read_text(encoding="utf-8"))
    return {}


def _open_session(host: str, config: HarnessConfig, wait: float):
    """Another session on this machine may hold the host's lock; wait for it
    rather than opting out of exclusivity, which could disturb its run."""
    started = time.monotonic()
    while True:
        try:
            session = SESSIONS[host](config)
            print(f"{host}: lock acquired after {time.monotonic() - started:.0f}s", flush=True)
            return session
        except SessionLockHeld:
            if time.monotonic() - started > wait:
                return None
            time.sleep(5)


def _injected_source(host: str, source: str) -> str:
    """Access writes `Option Compare Database` into every module it creates, so
    the case's own copy of that line would be a duplicate there. Dropping it
    leaves the module Access holds identical to the source the analyzer reads."""
    if host != "access":
        return source
    lines = source.split("\n")
    kept = [line for line in lines if line.strip().lower() != "option compare database"]
    return "\n".join(kept)


def _run_case(session, case: dict, watch: float) -> dict:
    kinds = {m.get("type", "standard") for m in case["modules"]}
    if not kinds <= {"standard", "class"}:
        return {"compile": "unsupported", "message": f"module types {sorted(kinds)}"}
    host = case.get("host") or "excel"
    try:
        session.new_document()
        for m in case["modules"]:
            session.add_module(m["name"], _injected_source(host, m["source"]), kind=m.get("type", "standard"))
        compiled = session.compile_project(watch_seconds=watch)
        record: dict = {"compile": compiled.outcome, "message": compiled.message}
        if case.get("run") and compiled.outcome == "accepted":
            ran = session.run_macro(case["run"], timeout=30)
            record["run"] = {
                "outcome": ran.outcome,
                "value": ran.value,
                "number": ran.error.number if ran.error else None,
                "description": ran.error.description if ran.error else None,
                "message": ran.message,
            }
        return record
    except Exception as error:  # recorded, so one broken case does not end the batch
        return {"compile": "harness-error", "message": f"{type(error).__name__}: {error}"}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("cases")
    parser.add_argument("results")
    parser.add_argument("--only", default="")
    parser.add_argument("--limit", type=int, default=0)
    parser.add_argument("--watch", type=float, default=10.0)
    parser.add_argument("--rerun", action="store_true", help="ignore results already recorded")
    parser.add_argument("--wait", type=float, default=420.0, help="seconds to wait for a host's lock")
    args = parser.parse_args()

    cases = json.loads(Path(args.cases).read_text(encoding="utf-8"))
    out = Path(args.results)
    results = {} if args.rerun else _load(out)
    todo = [c for c in cases if args.only in c["label"] and c["label"] not in results]
    if args.limit:
        todo = todo[: args.limit]
    print(f"{len(todo)} cases to run", flush=True)
    by_host: dict[str, list[dict]] = defaultdict(list)
    for case in todo:
        by_host[case.get("host") or "excel"].append(case)

    started = time.monotonic()
    config = HarnessConfig(compile_watch_s=args.watch)
    n = 0
    for host, group in by_host.items():
        session = _open_session(host, config, args.wait)
        if session is None:
            print(f"{host}: the lock stayed held; skipped {len(group)} cases", flush=True)
            continue
        with session as s:
            for case in group:
                n += 1
                record = _run_case(s, case, args.watch)
                results[case["label"]] = record
                out.write_text(json.dumps(results, indent=1, ensure_ascii=True, default=str), encoding="utf-8")
                run = record.get("run")
                tail = f" run={run['outcome']}:{run['number']}" if run else ""
                print(f"[{n}/{len(todo)}] {case['label']}: {record['compile']}{tail} {record.get('message', '')[:90]!r}", flush=True)
    print(f"done in {time.monotonic() - started:.0f}s", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
