"""Readers that turn files into analyzable VBA modules.

* loose_file: load and analyze loose .bas / .cls / .frm export files.
* workbook: load and analyze VBA out of Office macro containers via pyOpenVBA
  (Excel, Word, PowerPoint, and read-only Access).
* vbe_module: the shared export-format helper both readers build on.

Importing this package does not import pyOpenVBA; only calling the workbook
functions does.
"""

from __future__ import annotations

from .loose_file import (
    LOOSE_EXTENSIONS,
    LooseFileReadError,
    analyze_loose_file,
    analyze_loose_files,
    load_loose_module,
)
from .vbe_module import (
    LoadedModule,
    classify_module_kind,
    loaded_module_from_text,
    module_name_from_text,
    strip_export_header,
)
from .workbook import (
    ACCESS_EXTENSIONS,
    EXCEL_EXTENSIONS,
    OFFICE_EXTENSIONS,
    POWERPOINT_EXTENSIONS,
    WORD_EXTENSIONS,
    WorkbookReadError,
    analyze_office_file,
    analyze_workbook,
    read_office_modules,
    read_workbook_modules,
)

__all__ = [
    "ACCESS_EXTENSIONS",
    "EXCEL_EXTENSIONS",
    "LOOSE_EXTENSIONS",
    "OFFICE_EXTENSIONS",
    "POWERPOINT_EXTENSIONS",
    "WORD_EXTENSIONS",
    "LoadedModule",
    "LooseFileReadError",
    "WorkbookReadError",
    "analyze_loose_file",
    "analyze_loose_files",
    "analyze_office_file",
    "analyze_workbook",
    "classify_module_kind",
    "load_loose_module",
    "loaded_module_from_text",
    "module_name_from_text",
    "read_office_modules",
    "read_workbook_modules",
    "strip_export_header",
]
