"""Canonical VBA keyword table.

Ported from xlide_vscode/src/analyzer/lexer/keywordTable.ts. Verified against
MS-VBAL v20250520, section 3.3.5.2 (Reserved Identifiers and IDENTIFIER).

VBA identifiers are case-insensitive. This table maps a lowercased keyword to the
canonical capitalization the VBE emits. Two kinds are tracked: reserved
identifiers (the closed MS-VBAL 3.3.5.2 set) and contextual keywords (capitalized
by the VBE in a statement context but not reserved).
"""

from __future__ import annotations

# statement-keyword (MS-VBAL 3.3.5.2).
STATEMENT_KEYWORDS: tuple[str, ...] = (
    "Call", "Case", "Close", "Const", "Declare", "DefBool", "DefByte", "DefCur",
    "DefDate", "DefDbl", "DefInt", "DefLng", "DefLngLng", "DefLngPtr", "DefObj",
    "DefSng", "DefStr", "DefVar", "Dim", "Do", "Else", "ElseIf", "End", "EndIf",
    "Enum", "Erase", "Event", "Exit", "For", "Friend", "Function", "Get",
    "Global", "GoSub", "GoTo", "If", "Implements", "Input", "Let", "Lock",
    "Loop", "LSet", "Next", "On", "Open", "Option", "Print", "Private", "Public",
    "Put", "RaiseEvent", "ReDim", "Resume", "Return", "RSet", "Seek", "Select",
    "Set", "Static", "Stop", "Sub", "Type", "Unlock", "Wend", "While", "With",
    "Write",
)

# rem-keyword (MS-VBAL 3.3.5.2).
REM_KEYWORD = "Rem"

# marker-keyword (MS-VBAL 3.3.5.2).
MARKER_KEYWORDS: tuple[str, ...] = (
    "Any", "As", "ByRef", "ByVal", "Case", "Each", "Else", "In", "New", "Shared",
    "Until", "WithEvents", "Write", "Optional", "ParamArray", "Preserve", "Spc",
    "Tab", "Then", "To",
)

# operator-identifier (MS-VBAL 3.3.5.2).
OPERATOR_IDENTIFIERS: tuple[str, ...] = (
    "AddressOf", "And", "Eqv", "Imp", "Is", "Like", "New", "Mod", "Not", "Or",
    "TypeOf", "Xor",
)

# reserved-name (MS-VBAL 3.3.5.2).
RESERVED_NAMES: tuple[str, ...] = (
    "Abs", "CBool", "CByte", "CCur", "CDate", "CDbl", "CDec", "CInt", "CLng",
    "CLngLng", "CLngPtr", "CSng", "CStr", "CVar", "CVErr", "Date", "Debug",
    "DoEvents", "Fix", "Int", "Len", "LenB", "Me", "PSet", "Scale", "Sgn",
    "String",
)

# special-form (MS-VBAL 3.3.5.2).
SPECIAL_FORMS: tuple[str, ...] = (
    "Array", "Circle", "Input", "InputB", "LBound", "Scale", "UBound",
)

# reserved-type-identifier (MS-VBAL 3.3.5.2).
RESERVED_TYPE_IDENTIFIERS: tuple[str, ...] = (
    "Boolean", "Byte", "Currency", "Date", "Double", "Integer", "Long",
    "LongLong", "LongPtr", "Single", "String", "Variant",
)

# literal-identifier (MS-VBAL 3.3.5.2). Rendered capitalized to match the VBE.
LITERAL_IDENTIFIERS: tuple[str, ...] = (
    "True", "False", "Nothing", "Empty", "Null",
)

# future-reserved (MS-VBAL 3.3.5.2).
FUTURE_RESERVED: tuple[str, ...] = (
    "CDecl", "Decimal", "DefDec",
)

# reserved-for-implementation-use (MS-VBAL 3.3.5.2). Reserved for declaration
# validation but intentionally NOT keyword-cased, because exported source
# metadata uses them raw (e.g. `Attribute VB_Name = "Module1"`).
RESERVED_FOR_IMPLEMENTATION_USE: tuple[str, ...] = (
    "Attribute", "LINEINPUT", "VB_Base", "VB_Control", "VB_Creatable",
    "VB_Customizable", "VB_Description", "VB_Exposed", "VB_Ext_KEY",
    "VB_GlobalNameSpace", "VB_HelpID", "VB_Invoke_Func", "VB_Invoke_Property",
    "VB_Invoke_PropertyPut", "VB_Invoke_PropertyPutRef", "VB_MemberFlags",
    "VB_Name", "VB_PredeclaredId", "VB_ProcData", "VB_TemplateDerived",
    "VB_UserMemId", "VB_VarDescription", "VB_VarHelpID", "VB_VarMemberFlags",
    "VB_VarProcData", "VB_VarUserMemId",
)

# Contextual keywords: not reserved per MS-VBAL 3.3.5.2, but capitalized by the
# VBE in their statement context (VBE-convention casing).
CONTEXTUAL_KEYWORDS: tuple[str, ...] = (
    "Explicit", "Base", "Compare", "Binary", "Text",  # Option statements
    "Lib", "Alias",  # Declare
    "Property",  # Property statement keyword
    "Step",  # For ... Step
    "Error",  # On Error / Error
    "Output", "Append", "Random", "Read",  # Open ... For <mode>
    "Object",  # used as a type name
)

_KEYWORD_CASING_LISTS: tuple[tuple[str, ...], ...] = (
    STATEMENT_KEYWORDS,
    (REM_KEYWORD,),
    MARKER_KEYWORDS,
    OPERATOR_IDENTIFIERS,
    RESERVED_NAMES,
    SPECIAL_FORMS,
    RESERVED_TYPE_IDENTIFIERS,
    LITERAL_IDENTIFIERS,
    FUTURE_RESERVED,
)

# Lowercased names of every reserved identifier (MS-VBAL 3.3.5.2). A name in this
# set is never an <IDENTIFIER>.
RESERVED_IDENTIFIERS: frozenset[str] = frozenset(
    word.lower()
    for word_list in (*_KEYWORD_CASING_LISTS, RESERVED_FOR_IMPLEMENTATION_USE)
    for word in word_list
)


def _build_keyword_map() -> dict[str, str]:
    out: dict[str, str] = {}
    for word_list in (*_KEYWORD_CASING_LISTS, CONTEXTUAL_KEYWORDS):
        for word in word_list:
            # Reserved-identifier casing wins over contextual on conflict; there
            # are no conflicting spellings between the two sets today.
            out.setdefault(word.lower(), word)
    return out


# Lowercase keyword -> canonical capitalization. Covers all single-token reserved
# identifiers plus VBE-convention contextual keywords. Excludes
# reserved-for-implementation-use names, which stay raw Attribute-line spellings.
VBA_KEYWORDS: dict[str, str] = _build_keyword_map()


def canonical_keyword(word: str) -> str | None:
    """Canonical capitalization for a keyword, or None if not a known keyword."""
    return VBA_KEYWORDS.get(word.lower())


def is_reserved_identifier(word: str) -> bool:
    """True when word is a reserved identifier (MS-VBAL 3.3.5.2)."""
    return word.lower() in RESERVED_IDENTIFIERS
