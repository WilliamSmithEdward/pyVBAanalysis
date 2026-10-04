"""Translate Python source offsets to the UTF-16 offsets used by TypeScript."""

from __future__ import annotations


def utf16_offsets(source: str) -> list[int] | None:
    if all(ord(char) <= 0xFFFF for char in source):
        return None
    offsets = [0]
    for char in source:
        offsets.append(offsets[-1] + (2 if ord(char) > 0xFFFF else 1))
    return offsets
