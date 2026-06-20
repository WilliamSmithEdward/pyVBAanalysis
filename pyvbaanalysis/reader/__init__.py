"""Readers that turn files into analyzable VBA modules.

* loose_file: load and analyze loose .bas / .cls / .frm export files.
* workbook: load and analyze VBA out of Excel files via pyOpenVBA.
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
    EXCEL_EXTENSIONS,
    WorkbookReadError,
    analyze_workbook,
    read_workbook_modules,
)

__all__ = [
    "EXCEL_EXTENSIONS",
    "LOOSE_EXTENSIONS",
    "LoadedModule",
    "LooseFileReadError",
    "WorkbookReadError",
    "analyze_loose_file",
    "analyze_loose_files",
    "analyze_workbook",
    "classify_module_kind",
    "load_loose_module",
    "loaded_module_from_text",
    "module_name_from_text",
    "read_workbook_modules",
    "strip_export_header",
]
