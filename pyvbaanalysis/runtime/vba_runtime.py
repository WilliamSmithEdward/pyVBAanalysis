"""Built-in VBA runtime function/statement metadata (MS-VBAL Phase 9).

Ported from the function table of xlide_vscode/src/analyzer/runtime/vbaRuntime.ts:
the verified intrinsic functions and statements available in every VBA project
(MsgBox, Left, CLng, Now, Array, RGB, ...). Signatures are transcribed from the
Microsoft VBA language reference; they are never invented. The runtime constants
and global objects feed host/external-constant inference and are deferred (M9).
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class VbaRuntimeParam:
    name: str
    type_: str | None = None
    optional: bool = False
    param_array: bool = False


@dataclass(frozen=True, slots=True)
class VbaRuntimeFunction:
    name: str
    signature: str
    kind: str  # 'function' | 'statement'
    returns: str | None = None
    params: tuple[VbaRuntimeParam, ...] | None = None
    explicit_call: str | None = None  # 'allowed' | 'forbidden'


def _p(name: str, type_: str | None = None, optional: bool = False, param_array: bool = False) -> VbaRuntimeParam:
    return VbaRuntimeParam(name, type_, optional, param_array)


def _fn(
    name: str,
    signature: str,
    returns: str | None = None,
    params: tuple[VbaRuntimeParam, ...] | None = None,
    explicit_call: str | None = None,
) -> VbaRuntimeFunction:
    return VbaRuntimeFunction(name, signature, "function", returns, params, explicit_call)


def _stmt(
    name: str,
    signature: str,
    params: tuple[VbaRuntimeParam, ...] | None = None,
    explicit_call: str | None = None,
) -> VbaRuntimeFunction:
    return VbaRuntimeFunction(name, signature, "statement", None, params, explicit_call)


_LEFT_PARAMS = (_p("String", "String"), _p("Length", "Long"))
_RIGHT_PARAMS = (_p("String", "String"), _p("Length", "Long"))
_MID_PARAMS = (_p("String", "String"), _p("Start", "Long"), _p("Length", "Long", optional=True))
_SPACE_PARAMS = (_p("Number", "Long"),)
_STRING_PARAMS = (_p("Number", "Long"), _p("Character", "Variant"))
_REPLACE_PARAMS = (
    _p("Expression", "String"),
    _p("Find", "String"),
    _p("Replace", "String"),
    _p("Start", "Long", optional=True),
    _p("Count", "Long", optional=True),
    _p("Compare", "VbCompareMethod", optional=True),
)


VBA_RUNTIME_FUNCTIONS: tuple[VbaRuntimeFunction, ...] = (
    # -- Interaction --------------------------------------------------------
    _fn("MsgBox", "MsgBox(Prompt, [Buttons As VbMsgBoxStyle = vbOKOnly], [Title], [HelpFile], [Context]) As VbMsgBoxResult", "VbMsgBoxResult"),
    _fn("InputBox", "InputBox(Prompt, [Title], [Default], [XPos], [YPos], [HelpFile], [Context]) As String", "String"),
    _fn("Environ", "Environ(Expression) As String", "String"),
    _fn("Shell", "Shell(PathName, [WindowStyle As VbAppWinStyle = vbMinimizedFocus]) As Double", "Double"),
    _fn("DoEvents", "DoEvents() As Integer", "Integer", None, "forbidden"),
    _fn("Erl", "Erl() As Integer", "Integer", None, "forbidden"),
    _fn("CreateObject", "CreateObject(Class, [ServerName]) As Object", "Object"),
    _fn("GetObject", "GetObject([PathName], [Class]) As Object", "Object"),
    _stmt("Beep", "Beep"),
    _stmt("Randomize", "Randomize [Number]"),
    # -- String functions ---------------------------------------------------
    _fn("Len", "Len(Expression) As Long", "Long"),
    _fn("Left", "Left(String, Length) As String", "String", _LEFT_PARAMS),
    _fn("Left$", "Left$(String, Length) As String", "String", _LEFT_PARAMS),
    _fn("Right", "Right(String, Length) As String", "String", _RIGHT_PARAMS),
    _fn("Right$", "Right$(String, Length) As String", "String", _RIGHT_PARAMS),
    _fn("Mid", "Mid(String, Start, [Length]) As String", "String", _MID_PARAMS),
    _fn("Mid$", "Mid$(String, Start, [Length]) As String", "String", _MID_PARAMS),
    _fn("Trim", "Trim(String) As String", "String"),
    _fn("LTrim", "LTrim(String) As String", "String"),
    _fn("RTrim", "RTrim(String) As String", "String"),
    _fn("UCase", "UCase(String) As String", "String"),
    _fn("LCase", "LCase(String) As String", "String"),
    _fn("Replace", "Replace(Expression, Find, Replace, [Start = 1], [Count = -1], [Compare As VbCompareMethod = vbBinaryCompare]) As String", "String", _REPLACE_PARAMS),
    _fn("Replace$", "Replace$(Expression, Find, Replace, [Start = 1], [Count = -1], [Compare As VbCompareMethod = vbBinaryCompare]) As String", "String", _REPLACE_PARAMS),
    _fn("InStr", "InStr([Start], String1, String2, [Compare As VbCompareMethod]) As Long", "Long"),
    _fn("InStrRev", "InStrRev(StringCheck, StringMatch, [Start = -1], [Compare As VbCompareMethod = vbBinaryCompare]) As Long", "Long"),
    _fn("Split", "Split(Expression, [Delimiter], [Limit = -1], [Compare As VbCompareMethod = vbBinaryCompare]) As String()", "String()"),
    _fn("Join", "Join(SourceArray, [Delimiter]) As String", "String"),
    _fn("StrComp", "StrComp(String1, String2, [Compare As VbCompareMethod]) As Integer", "Integer"),
    _fn("StrConv", "StrConv(String, Conversion As VbStrConv, [LCID]) As String", "String"),
    _fn("String$", "String$(Number, Character) As String", "String", _STRING_PARAMS),
    _fn("Space", "Space(Number) As String", "String", _SPACE_PARAMS),
    _fn("Space$", "Space$(Number) As String", "String", _SPACE_PARAMS),
    _fn("Format", "Format(Expression, [Format], [FirstDayOfWeek As VbDayOfWeek], [FirstWeekOfYear As VbFirstWeekOfYear]) As String", "String"),
    _fn("Chr", "Chr(CharCode) As String", "String"),
    _fn("ChrW", "ChrW(CharCode) As String", "String"),
    _fn("Asc", "Asc(String) As Integer", "Integer"),
    _fn("AscW", "AscW(String) As Integer", "Integer"),
    # -- Type conversion ----------------------------------------------------
    _fn("CBool", "CBool(Expression) As Boolean", "Boolean"),
    _fn("CByte", "CByte(Expression) As Byte", "Byte"),
    _fn("CCur", "CCur(Expression) As Currency", "Currency"),
    _fn("CDate", "CDate(Expression) As Date", "Date"),
    _fn("CDbl", "CDbl(Expression) As Double", "Double"),
    _fn("CDec", "CDec(Expression) As Variant", "Variant"),
    _fn("CInt", "CInt(Expression) As Integer", "Integer"),
    _fn("CLng", "CLng(Expression) As Long", "Long"),
    _fn("CLngLng", "CLngLng(Expression) As LongLong", "LongLong"),
    _fn("CSng", "CSng(Expression) As Single", "Single"),
    _fn("CStr", "CStr(Expression) As String", "String"),
    _fn("CVar", "CVar(Expression) As Variant", "Variant"),
    _fn("Val", "Val(String) As Double", "Double"),
    _fn("Hex", "Hex(Number) As String", "String"),
    _fn("Oct", "Oct(Number) As String", "String"),
    # -- Math ---------------------------------------------------------------
    _fn("Abs", "Abs(Number) As Variant", "Variant"),
    _fn("Int", "Int(Number) As Variant", "Variant"),
    _fn("Fix", "Fix(Number) As Variant", "Variant"),
    _fn("Sgn", "Sgn(Number) As Integer", "Integer"),
    _fn("Sqr", "Sqr(Number) As Double", "Double"),
    _fn("Exp", "Exp(Number) As Double", "Double"),
    _fn("Log", "Log(Number) As Double", "Double"),
    _fn("Sin", "Sin(Number) As Double", "Double"),
    _fn("Cos", "Cos(Number) As Double", "Double"),
    _fn("Tan", "Tan(Number) As Double", "Double"),
    _fn("Atn", "Atn(Number) As Double", "Double"),
    _fn("Round", "Round(Number, [NumDigitsAfterDecimal]) As Double", "Double"),
    _fn("Rnd", "Rnd([Number]) As Single", "Single"),
    # -- Date / time --------------------------------------------------------
    _fn("Now", "Now() As Date", "Date"),
    _fn("Timer", "Timer() As Single", "Single"),
    _fn("Year", "Year(Date) As Integer", "Integer"),
    _fn("Month", "Month(Date) As Integer", "Integer"),
    _fn("Day", "Day(Date) As Integer", "Integer"),
    _fn("Hour", "Hour(Time) As Integer", "Integer"),
    _fn("Minute", "Minute(Time) As Integer", "Integer"),
    _fn("Second", "Second(Time) As Integer", "Integer"),
    _fn("Weekday", "Weekday(Date, [FirstDayOfWeek As VbDayOfWeek = vbSunday]) As Integer", "Integer"),
    _fn("MonthName", "MonthName(Month, [Abbreviate As Boolean = False]) As String", "String"),
    _fn("WeekdayName", "WeekdayName(Weekday, [Abbreviate As Boolean = False], [FirstDayOfWeek As VbDayOfWeek = vbUseSystemDayOfWeek]) As String", "String"),
    _fn("DateAdd", "DateAdd(Interval, Number, Date) As Date", "Date"),
    _fn("DateDiff", "DateDiff(Interval, Date1, Date2, [FirstDayOfWeek As VbDayOfWeek], [FirstWeekOfYear As VbFirstWeekOfYear]) As Long", "Long"),
    _fn("DatePart", "DatePart(Interval, Date, [FirstDayOfWeek As VbDayOfWeek], [FirstWeekOfYear As VbFirstWeekOfYear]) As Integer", "Integer"),
    _fn("DateSerial", "DateSerial(Year, Month, Day) As Date", "Date"),
    _fn("TimeSerial", "TimeSerial(Hour, Minute, Second) As Date", "Date"),
    _fn("DateValue", "DateValue(Date) As Date", "Date"),
    _fn("TimeValue", "TimeValue(Time) As Date", "Date"),
    # -- Arrays / inspection ------------------------------------------------
    _fn("Array", "Array(ArgList) As Variant", "Variant", (_p("ArgList", "Variant", optional=True, param_array=True),)),
    _fn("LBound", "LBound(ArrayName, [Dimension = 1]) As Long", "Long"),
    _fn("UBound", "UBound(ArrayName, [Dimension = 1]) As Long", "Long"),
    _fn("IsArray", "IsArray(VarName) As Boolean", "Boolean"),
    _fn("IsDate", "IsDate(Expression) As Boolean", "Boolean"),
    _fn("IsEmpty", "IsEmpty(Expression) As Boolean", "Boolean"),
    _fn("IsError", "IsError(Expression) As Boolean", "Boolean"),
    _fn("IsNull", "IsNull(Expression) As Boolean", "Boolean"),
    _fn("IsNumeric", "IsNumeric(Expression) As Boolean", "Boolean"),
    _fn("IsObject", "IsObject(Expression) As Boolean", "Boolean"),
    _fn("VarType", "VarType(VarName) As VbVarType", "VbVarType"),
    _fn("TypeName", "TypeName(VarName) As String", "String"),
    # -- Selection / colour -------------------------------------------------
    _fn("RGB", "RGB(Red, Green, Blue) As Long", "Long"),
    _fn("QBColor", "QBColor(Color) As Long", "Long"),
    _fn("IIf", "IIf(Expression, TruePart, FalsePart) As Variant", "Variant"),
    _fn("Choose", "Choose(Index, ArgList) As Variant", "Variant"),
    _fn("Switch", "Switch(ArgList) As Variant", "Variant"),
    _fn("IsMissing", "IsMissing(ArgName) As Boolean", "Boolean"),
    _fn("CallByName", "CallByName(Object, ProcName As String, CallType As VbCallType, [Args]) As Variant", "Variant", (
        _p("Object", "Object"),
        _p("ProcName", "String"),
        _p("CallType", "VbCallType"),
        _p("Args", "Variant", optional=True, param_array=True),
    )),
    # -- Additional string functions ---------------------------------------
    _fn("StrReverse", "StrReverse(Expression) As String", "String"),
    _fn("LenB", "LenB(Expression) As Long", "Long"),
    _fn("Str", "Str(Number) As String", "String"),
    _fn("Filter", "Filter(SourceArray, Match, [Include As Boolean = True], [Compare As VbCompareMethod = vbBinaryCompare]) As String()", "String()"),
    _fn("FormatCurrency", "FormatCurrency(Expression, [NumDigitsAfterDecimal = -1], [IncludeLeadingDigit As VbTriState = vbUseDefault], [UseParensForNegativeNumbers As VbTriState = vbUseDefault], [GroupDigits As VbTriState = vbUseDefault]) As String", "String"),
    _fn("FormatNumber", "FormatNumber(Expression, [NumDigitsAfterDecimal = -1], [IncludeLeadingDigit As VbTriState = vbUseDefault], [UseParensForNegativeNumbers As VbTriState = vbUseDefault], [GroupDigits As VbTriState = vbUseDefault]) As String", "String"),
    _fn("FormatPercent", "FormatPercent(Expression, [NumDigitsAfterDecimal = -1], [IncludeLeadingDigit As VbTriState = vbUseDefault], [UseParensForNegativeNumbers As VbTriState = vbUseDefault], [GroupDigits As VbTriState = vbUseDefault]) As String", "String"),
    _fn("FormatDateTime", "FormatDateTime(Date, [NamedFormat As VbDateTimeFormat = vbGeneralDate]) As String", "String"),
    # -- Conversion helpers -------------------------------------------------
    _fn("CVErr", "CVErr(ErrorNumber) As Variant", "Variant"),
    # -- Interaction / registry --------------------------------------------
    _fn("Command", "Command() As String", "String"),
    _fn("Partition", "Partition(Number, Start, Stop, Interval) As String", "String"),
    _fn("GetSetting", "GetSetting(AppName, Section, Key, [Default]) As String", "String"),
    _fn("GetAllSettings", "GetAllSettings(AppName, Section) As Variant", "Variant"),
    _stmt("SaveSetting", "SaveSetting AppName, Section, Key, Setting"),
    _stmt("DeleteSetting", "DeleteSetting AppName, [Section], [Key]"),
    _stmt("AppActivate", "AppActivate Title, [Wait As Boolean]"),
    _stmt("SendKeys", "SendKeys String, [Wait]"),
    # -- File system --------------------------------------------------------
    _fn("Dir", "Dir([PathName], [Attributes As VbFileAttribute = vbNormal]) As String", "String"),
    _fn("FreeFile", "FreeFile([RangeNumber]) As Integer", "Integer"),
    _fn("EOF", "EOF(FileNumber) As Boolean", "Boolean"),
    _fn("LOF", "LOF(FileNumber) As Long", "Long"),
    _fn("Loc", "Loc(FileNumber) As Long", "Long"),
    _fn("Seek", "Seek(FileNumber) As Long", "Long"),
    _fn("FileLen", "FileLen(PathName) As Long", "Long"),
    _fn("FileDateTime", "FileDateTime(PathName) As Date", "Date"),
    _fn("GetAttr", "GetAttr(PathName) As VbFileAttribute", "VbFileAttribute"),
    _fn("CurDir", "CurDir([Drive]) As String", "String"),
    _stmt("ChDir", "ChDir Path"),
    _stmt("ChDrive", "ChDrive Drive"),
    _stmt("MkDir", "MkDir Path"),
    _stmt("RmDir", "RmDir Path"),
    _stmt("Kill", "Kill PathName"),
    _stmt("FileCopy", "FileCopy Source, Destination"),
    _stmt("SetAttr", "SetAttr PathName, Attributes As VbFileAttribute"),
    # -- Financial ----------------------------------------------------------
    _fn("PV", "PV(Rate, NPer, Pmt, [FV = 0], [Type = 0]) As Double", "Double"),
    _fn("FV", "FV(Rate, NPer, Pmt, [PV = 0], [Type = 0]) As Double", "Double"),
    _fn("Pmt", "Pmt(Rate, NPer, PV, [FV = 0], [Type = 0]) As Double", "Double"),
    _fn("IPmt", "IPmt(Rate, Per, NPer, PV, [FV = 0], [Type = 0]) As Double", "Double"),
    _fn("PPmt", "PPmt(Rate, Per, NPer, PV, [FV = 0], [Type = 0]) As Double", "Double"),
    _fn("NPer", "NPer(Rate, Pmt, PV, [FV = 0], [Type = 0]) As Double", "Double"),
    _fn("Rate", "Rate(NPer, Pmt, PV, [FV = 0], [Type = 0], [Guess = 0.1]) As Double", "Double"),
    _fn("NPV", "NPV(Rate, ValueArray()) As Double", "Double"),
    _fn("IRR", "IRR(ValueArray(), [Guess = 0.1]) As Double", "Double"),
    _fn("MIRR", "MIRR(ValueArray(), FinanceRate, ReinvestRate) As Double", "Double"),
    _fn("SLN", "SLN(Cost, Salvage, Life) As Double", "Double"),
    _fn("SYD", "SYD(Cost, Salvage, Life, Period) As Double", "Double"),
    _fn("DDB", "DDB(Cost, Salvage, Life, Period, [Factor = 2]) As Double", "Double"),
)


_BY_LOWER = {f.name.lower(): f for f in VBA_RUNTIME_FUNCTIONS}


def resolve_runtime_function(name: str) -> VbaRuntimeFunction | None:
    """Resolve a built-in VBA runtime function/statement by name (case-insensitive)."""
    return _BY_LOWER.get(name.lower())


def runtime_allows_explicit_call(fn: VbaRuntimeFunction) -> bool:
    """Whether this runtime entry may be the target of an explicit `Call` statement."""
    return fn.explicit_call != "forbidden"
