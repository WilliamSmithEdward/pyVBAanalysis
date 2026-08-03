"""Build the VBE-oracle probes for combining-mark identifiers.

Reproduces the evidence behind the Mn/Mc branch in ``_is_ident_part``
(pyvbaanalysis/lexer/tokenize.py). Run on Windows with Excel installed::

    python tools/oracle/build_combining_mark_probe.py build
    powershell -ExecutionPolicy Bypass -File RUNNER -WorkbookPath WORKBOOK -MacroName RunGate

where RUNNER is pyOpenVBA's bounded popup-aware harness
(``tools/live_excel/run_macro.ps1``) and WORKBOOK is one of the two files this
script writes. The harness prints a JSON line and dismisses any VBE modal, so a
compile error surfaces as a non "run-ok" outcome with the dialog text captured
rather than hanging on a dialog.

Recorded result (2026-08-03, Excel 16, Office16): both probes report
``{"outcome": "run-ok", "popups": []}`` and write ``ran=ok``, so the VBE accepts
an identifier containing a Thai tone mark. Splitting it, as the lexer did
before, was a false positive on valid VBA.

pyVBAanalysis consumes oracle evidence rather than generating it (agent.md
section 2), so this script exists to make the finding reproducible; the corpus
entry itself belongs upstream in XLIDE.

Two probes, so a failure can be attributed:

* control   - a Thai identifier of base letters only (no combining marks).
* combining - the same shape but with a tone mark plus a vowel sign.

Module source is stored in the project's code page, so both need a genuine
cp874 project: the encoder has to use cp874 AND the PROJECTCODEPAGE record in
the dir stream has to say 874, or Excel decodes the bytes as cp1252 and the
probe tests mojibake instead of Thai. The build asserts the identifier
round-trips back out of storage before Excel ever opens the file.
"""

from __future__ import annotations

import struct
import sys
from pathlib import Path

from pyopenvba import ExcelFile

CONTROL_IDENT = "ผล"  # pho + lo: two base letters
COMBINING_IDENT = "ค่า"  # kho + mai ek (Mn) + sara aa
REC_PROJECTCODEPAGE = 0x0003
CODE_PAGE = 874


def patch_dir_code_page(dir_raw: bytearray, code_page: int) -> bool:
    """Rewrite the PROJECTCODEPAGE record in a decompressed dir stream."""
    pos = 0
    while pos + 6 <= len(dir_raw):
        rec_id, size = struct.unpack_from("<HI", dir_raw, pos)
        if rec_id == REC_PROJECTCODEPAGE and size >= 2:
            struct.pack_into("<H", dir_raw, pos + 6, code_page)
            return True
        pos += 6 + size
    return False


def module_source(ident: str) -> str:
    return (
        'Attribute VB_Name = "Module1"\r\n'
        "\r\n"
        "Sub RunGate()\r\n"
        f"    Dim {ident} As String\r\n"
        f'    {ident} = "ok"\r\n'
        "    Dim f As Integer\r\n"
        "    f = FreeFile\r\n"
        '    Open ThisWorkbook.Path & "\\thai_sentinel.txt" For Output As #f\r\n'
        f'    Print #f, "ran=" & {ident}\r\n'
        "    Close #f\r\n"
        "End Sub\r\n"
    )


def main() -> None:
    if len(sys.argv) != 2:
        raise SystemExit("usage: python tools/oracle/build_combining_mark_probe.py OUTDIR")
    out_dir = Path(sys.argv[1])
    out_dir.mkdir(parents=True, exist_ok=True)

    for label, ident in (("control", CONTROL_IDENT), ("combining", COMBINING_IDENT)):
        target = out_dir / f"thai_{label}.xlsm"
        if target.exists():
            target.unlink()
        with ExcelFile.create_new(target) as wb:
            project = wb.vba_project()
            raw = bytearray(project.dir_raw)
            if not patch_dir_code_page(raw, CODE_PAGE):
                raise SystemExit(f"{label}: PROJECTCODEPAGE record not found")
            project.dir_raw = bytes(raw)
            project.dir_structure_dirty = True
            project.code_page = CODE_PAGE
            wb.set_module("Module1", module_source(ident))
            wb.save()

        with ExcelFile(target) as wb:
            stored = wb.get_module("Module1")
            page = wb.vba_project().code_page
        if ident not in stored:
            raise SystemExit(
                f"{label}: identifier did not survive storage (code_page={page}); "
                "the probe would test mojibake, not Thai"
            )
        print(f"{label}: code_page={page} identifier_stored=True path={target}")


if __name__ == "__main__":
    main()
