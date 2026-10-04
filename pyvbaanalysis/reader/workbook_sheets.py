"""Saved sheet names and kinds, ported from XLIDE vba/workbookSheets.ts."""

from __future__ import annotations

import struct
from collections.abc import Iterator
from typing import Any, Literal

from defusedxml import ElementTree

from ..symbols.sheet_changes import WorkbookSheetInfo

SheetKind = Literal["worksheet", "chartsheet", "dialogsheet", "macrosheet"]
_KINDS: dict[str, SheetKind] = {
    "worksheets": "worksheet", "chartsheets": "chartsheet",
    "dialogsheets": "dialogsheet", "macrosheets": "macrosheet",
}
_BIFF_KINDS: dict[int, SheetKind] = {0: "worksheet", 1: "macrosheet", 2: "chartsheet"}


def _kind(target: str) -> SheetKind | None:
    return next((kind for folder, kind in _KINDS.items() if f"{folder}/" in target), None)


def _relationships(zip_file: Any, part: str) -> dict[str, str]:
    root = ElementTree.fromstring(zip_file.read(part))
    return {node.attrib["Id"]: node.attrib["Target"] for node in root if "Id" in node.attrib and "Target" in node.attrib}


def _varint(data: bytes, pos: int, limit: int) -> tuple[int, int]:
    value = 0
    for shift in range(limit):
        if pos >= len(data):
            raise ValueError("Truncated XLSB record header")
        byte = data[pos]
        pos += 1
        value |= (byte & 127) << (7 * shift)
        if not byte & 128:
            return value, pos
    return value, pos


def _xlsb_records(data: bytes) -> Iterator[tuple[int, bytes]]:
    pos = 0
    while pos < len(data):
        kind, pos = _varint(data, pos, 2)
        size, pos = _varint(data, pos, 4)
        if pos + size > len(data):
            raise ValueError("Truncated XLSB record")
        yield kind, data[pos:pos + size]
        pos += size


def _wide_string(data: bytes, pos: int) -> tuple[str, int]:
    count = struct.unpack_from("<I", data, pos)[0]
    pos += 4
    if count == 0xFFFFFFFF:
        return "", pos
    end = pos + count * 2
    if end > len(data):
        raise ValueError("Truncated XLSB string")
    return data[pos:end].decode("utf-16-le"), end


def _zip_sheets(zip_file: Any, binary: bool) -> list[WorkbookSheetInfo]:
    part = "workbook.bin" if binary else "workbook.xml"
    relationships = _relationships(zip_file, f"xl/_rels/{part}.rels")
    data = zip_file.read(f"xl/{part}")
    sheets = []
    if binary:
        for record, body in _xlsb_records(data):
            if record != 0x9C:
                continue
            rel_id, pos = _wide_string(body, 8)
            name, _ = _wide_string(body, pos)
            kind = _kind(relationships.get(rel_id, ""))
            if kind is not None:
                sheets.append(WorkbookSheetInfo(name, kind))
    else:
        root = ElementTree.fromstring(data)
        for node in root.iter():
            if node.tag.rsplit("}", 1)[-1] != "sheet":
                continue
            rel_id = next((value for key, value in node.attrib.items() if key.rsplit("}", 1)[-1] == "id"), "")
            kind = _kind(relationships.get(rel_id, ""))
            if kind is not None:
                sheets.append(WorkbookSheetInfo(node.attrib.get("name", ""), kind))
    return sheets


def _biff_sheets(data: bytes) -> list[WorkbookSheetInfo]:
    sheets = []
    pos = 0
    while pos + 4 <= len(data):
        record, size = struct.unpack_from("<HH", data, pos)
        pos += 4
        if pos + size > len(data):
            raise ValueError("Truncated BIFF record")
        body = data[pos:pos + size]
        pos += size
        if record == 0xA:
            break
        if record != 0x85 or len(body) < 8:
            continue
        kind = _BIFF_KINDS.get(body[5])
        if kind is not None:
            count, wide = body[6], body[7] & 1
            end = 8 + count * (2 if wide else 1)
            if end > len(body):
                raise ValueError("Truncated BIFF sheet name")
            name = body[8:end].decode("utf-16-le" if wide else "latin-1")
            sheets.append(WorkbookSheetInfo(name, kind))
    return sheets


def workbook_sheets(container: Any, suffix: str) -> list[WorkbookSheetInfo] | None:
    """Unknown metadata stays None; an empty list means a readable empty list."""
    try:
        if suffix == ".xls":
            return _biff_sheets(container._get_cfb().get_stream("Workbook"))
        zip_file = getattr(container, "_zip", None)
        return None if zip_file is None else _zip_sheets(zip_file, suffix == ".xlsb")
    except Exception:
        return None
