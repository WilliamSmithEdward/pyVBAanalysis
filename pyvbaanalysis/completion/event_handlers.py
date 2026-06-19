"""Excel event-handler catalogue (port of completion/eventHandlers.ts).

Only the seams the ``eventHandlerModuleScope`` diagnostic consumes are ported:
``event_handler_procedure_for_name`` (name -> owning document type) and
``event_handler_document_type_for_context`` (module facts -> the document type
Excel wires events for). The catalogue itself is vendored as
``data/event_definitions.json``, mechanically extracted from XLIDE by
``tools/extract_event_definitions.mjs`` (no hand-transcription). The completion-UX
stub generation is intentionally not ported.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path

from ..symbols.symbol_model import ModuleSymbolKind

_DATA_DIR = Path(__file__).resolve().parent.parent / "data"

# Document types an event can be wired for; mirrors EventHandlerDocumentType.
EventHandlerDocumentType = str  # 'workbook' | 'worksheet' | 'chart'

_CHART_NAME_RE = re.compile(r"^chart\d*$", re.IGNORECASE)


@dataclass(frozen=True, slots=True)
class EventHandlerProcedureMatch:
    """A procedure name that matches a known Excel event handler."""

    name: str
    owner: str  # 'Workbook' | 'Worksheet' | 'Chart'
    document_type: EventHandlerDocumentType


@lru_cache(maxsize=1)
def _event_definitions_by_lower_name() -> dict[str, EventHandlerProcedureMatch]:
    raw = json.loads((_DATA_DIR / "event_definitions.json").read_text(encoding="utf-8"))
    out: dict[str, EventHandlerProcedureMatch] = {}
    for entry in raw["events"]:
        out[entry["name"].lower()] = EventHandlerProcedureMatch(
            name=entry["name"], owner=entry["owner"], document_type=entry["documentType"]
        )
    return out


def event_handler_procedure_for_name(name: str) -> EventHandlerProcedureMatch | None:
    """The Excel event a procedure name matches (case-insensitive), or None.

    Port of eventHandlerProcedureForName. Every catalogue owner maps to a document
    type, so a name match always yields a populated match.
    """
    return _event_definitions_by_lower_name().get(name.lower())


def _infer_document_type(module_name: str | None) -> EventHandlerDocumentType:
    """Port of inferDocumentType: name-based fallback when documentType is unset."""
    lower = (module_name or "").lower()
    if lower == "thisworkbook":
        return "workbook"
    if _CHART_NAME_RE.match(module_name or ""):
        return "chart"
    return "worksheet"


def event_handler_document_type_for_context(
    module_name: str | None,
    module_kind: ModuleSymbolKind | None,
    document_type: EventHandlerDocumentType | None,
) -> EventHandlerDocumentType | None:
    """Port of eventHandlerDocumentTypeForContext.

    Only document modules wire events; for those the caller-supplied document type
    wins, falling back to a name heuristic (ThisWorkbook -> workbook, Chart* ->
    chart, otherwise worksheet).
    """
    if module_kind is not ModuleSymbolKind.DOCUMENT:
        return None
    return document_type if document_type is not None else _infer_document_type(module_name)
