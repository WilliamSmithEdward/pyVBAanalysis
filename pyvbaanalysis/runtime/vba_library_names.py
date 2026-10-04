"""Ported from xlide_vscode/src/analyzer/runtime/vbaLibraryNames.ts.

Its tables were generated from the TS module's own values, not transcribed by hand.
"""

from __future__ import annotations

from collections.abc import Mapping

# Generated from the VBA type library, VBE7.DLL (VBA 7.1), for issue #369
# (scratchpad dump_vba_typelib.py + gen_vba_library_names.py, 2026-10-01).
# The library is closed: the VBE refuses any other name after `VBA.`, after
# a VBA module or enum, or after VBA.Global, as "Method or data member not
# found" (measured in Excel 16.0). `_B_str_Left` is spelled Left$ and
# `_B_var_Left` Left. Every name is lowercased.


# Every name `VBA.` reaches: modules, enums, classes, and their members.
VBA_LIBRARY_NAMES: frozenset[str] = frozenset({
    "_collection", "_errobject", "_formshowconstants", "_hiddeninterface", "_hiddenmodule", "abs",
    "appactivate", "array", "asc", "ascb", "ascw", "atn", "beep", "calendar", "callbyname", "cbool",
    "cbyte", "ccur", "cdate", "cdbl", "cdec", "chdir", "chdrive", "choose", "chr", "chr$", "chrb",
    "chrb$", "chrw", "chrw$", "cint", "clng", "clnglng", "clngptr", "collection", "colorconstants",
    "command", "command$", "constants", "conversion", "cos", "createobject", "csng", "cstr",
    "cstr$", "curdir", "curdir$", "cvar", "cvdate", "cverr", "date", "date$", "dateadd", "datediff",
    "datepart", "dateserial", "datetime", "datevalue", "day", "ddb", "deletesetting", "dir", "dir$",
    "doevents", "environ", "environ$", "eof", "erl", "err", "errobject", "error", "error$", "exp",
    "fileattr", "filecopy", "filedatetime", "filelen", "filesystem", "filter", "financial", "fix",
    "format", "format$", "formatcurrency", "formatcurrency$", "formatdatetime", "formatdatetime$",
    "formatnumber", "formatnumber$", "formatpercent", "formatpercent$", "formshowconstants",
    "freefile", "fv", "getallsettings", "getattr", "getobject", "getsetting", "getsetting$",
    "global", "hex", "hex$", "hour", "iif", "imatch", "imatchcollection", "imestatus",
    "information", "input", "input$", "inputb", "inputb$", "inputbox", "inputbox$", "instr",
    "instrb", "instrrev", "int", "interaction", "ipmt", "iregexp", "irr", "isarray", "isdate",
    "isempty", "iserror", "ismissing", "isnull", "isnumeric", "isobject", "isubmatches", "join",
    "join$", "keycodeconstants", "kill", "lcase", "lcase$", "left", "left$", "leftb", "leftb$",
    "len", "lenb", "load", "loc", "lof", "log", "long_ptr", "ltrim", "ltrim$", "macid", "macscript",
    "macscript$", "match", "matchcollection", "math", "mid", "mid$", "midb", "midb$", "minute",
    "mirr", "mkdir", "month", "monthname", "monthname$", "msgbox", "now", "nper", "npv", "objptr",
    "oct", "oct$", "partition", "pmt", "ppmt", "pv", "qbcolor", "randomize", "rate", "regexp",
    "replace", "replace$", "reset", "rgb", "right", "right$", "rightb", "rightb$", "rmdir", "rnd",
    "round", "rtrim", "rtrim$", "savesetting", "second", "seek", "sendkeys", "setattr", "sgn",
    "shell", "sin", "sln", "space", "space$", "split", "sqr", "str", "str$", "strcomp", "strconv",
    "string", "string$", "strings", "strptr", "strreverse", "strreverse$", "submatches", "switch",
    "syd", "systemcolorconstants", "tan", "time", "time$", "timer", "timeserial", "timevalue",
    "trim", "trim$", "typename", "typename$", "ucase", "ucase$", "unload", "userforms", "val",
    "varptr", "vartype", "vb3ddkshadow", "vb3dface", "vb3dhighlight", "vb3dlight", "vb3dshadow",
    "vbabort", "vbabortretryignore", "vbactiveborder", "vbactivetitlebar", "vbalias",
    "vbapplicationmodal", "vbapplicationworkspace", "vbapptaskmanager", "vbappwindows",
    "vbappwinstyle", "vbarchive", "vbarray", "vbback", "vbbinarycompare", "vbblack", "vbblue",
    "vbboolean", "vbbuttonface", "vbbuttonshadow", "vbbuttontext", "vbbyte", "vbcalendar",
    "vbcalgreg", "vbcalhijri", "vbcalltype", "vbcancel", "vbcomparemethod", "vbcr", "vbcritical",
    "vbcrlf", "vbcurrency", "vbcyan", "vbdatabasecompare", "vbdataobject", "vbdate",
    "vbdatetimeformat", "vbdayofweek", "vbdecimal", "vbdefaultbutton1", "vbdefaultbutton2",
    "vbdefaultbutton3", "vbdefaultbutton4", "vbdesktop", "vbdirectory", "vbdouble", "vbeglobal",
    "vbempty", "vberror", "vbexclamation", "vbfalse", "vbfileattribute", "vbfirstfourdays",
    "vbfirstfullweek", "vbfirstjan1", "vbfirstweekofyear", "vbformcode", "vbformcontrolmenu",
    "vbformfeed", "vbformmdiform", "vbfriday", "vbfromunicode", "vbgeneraldate", "vbget",
    "vbgraytext", "vbgreen", "vbhidden", "vbhide", "vbhighlight", "vbhighlighttext", "vbhiragana",
    "vbignore", "vbimealphadbl", "vbimealphasng", "vbimedisable", "vbimehiragana",
    "vbimekatakanadbl", "vbimekatakanasng", "vbimemodealpha", "vbimemodealphafull",
    "vbimemodedisable", "vbimemodehangul", "vbimemodehangulfull", "vbimemodehiragana",
    "vbimemodekatakana", "vbimemodekatakanahalf", "vbimemodenocontrol", "vbimemodeoff",
    "vbimemodeon", "vbimenoop", "vbimeoff", "vbimeon", "vbimestatus", "vbinactiveborder",
    "vbinactivecaptiontext", "vbinactivetitlebar", "vbinfobackground", "vbinformation",
    "vbinfotext", "vbinteger", "vbkatakana", "vbkey0", "vbkey1", "vbkey2", "vbkey3", "vbkey4",
    "vbkey5", "vbkey6", "vbkey7", "vbkey8", "vbkey9", "vbkeya", "vbkeyadd", "vbkeyb", "vbkeyback",
    "vbkeyc", "vbkeycancel", "vbkeycapital", "vbkeyclear", "vbkeycontrol", "vbkeyd", "vbkeydecimal",
    "vbkeydelete", "vbkeydivide", "vbkeydown", "vbkeye", "vbkeyend", "vbkeyescape", "vbkeyexecute",
    "vbkeyf", "vbkeyf1", "vbkeyf10", "vbkeyf11", "vbkeyf12", "vbkeyf13", "vbkeyf14", "vbkeyf15",
    "vbkeyf16", "vbkeyf2", "vbkeyf3", "vbkeyf4", "vbkeyf5", "vbkeyf6", "vbkeyf7", "vbkeyf8",
    "vbkeyf9", "vbkeyg", "vbkeyh", "vbkeyhelp", "vbkeyhome", "vbkeyi", "vbkeyinsert", "vbkeyj",
    "vbkeyk", "vbkeyl", "vbkeylbutton", "vbkeyleft", "vbkeym", "vbkeymbutton", "vbkeymenu",
    "vbkeymultiply", "vbkeyn", "vbkeynumlock", "vbkeynumpad0", "vbkeynumpad1", "vbkeynumpad2",
    "vbkeynumpad3", "vbkeynumpad4", "vbkeynumpad5", "vbkeynumpad6", "vbkeynumpad7", "vbkeynumpad8",
    "vbkeynumpad9", "vbkeyo", "vbkeyp", "vbkeypagedown", "vbkeypageup", "vbkeypause", "vbkeyprint",
    "vbkeyq", "vbkeyr", "vbkeyrbutton", "vbkeyreturn", "vbkeyright", "vbkeys", "vbkeyselect",
    "vbkeyseparator", "vbkeyshift", "vbkeysnapshot", "vbkeyspace", "vbkeysubtract", "vbkeyt",
    "vbkeytab", "vbkeyu", "vbkeyup", "vbkeyv", "vbkeyw", "vbkeyx", "vbkeyy", "vbkeyz", "vblet",
    "vblf", "vblong", "vblongdate", "vblonglong", "vblongtime", "vblowercase", "vbmagenta",
    "vbmaximizedfocus", "vbmenubar", "vbmenutext", "vbmethod", "vbminimizedfocus",
    "vbminimizednofocus", "vbmodal", "vbmodeless", "vbmonday", "vbmsgbox", "vbmsgboxhelpbutton",
    "vbmsgboxresult", "vbmsgboxright", "vbmsgboxrtlreading", "vbmsgboxsetforeground",
    "vbmsgboxstyle", "vbmsgboxtext", "vbnarrow", "vbnewline", "vbno", "vbnormal", "vbnormalfocus",
    "vbnormalnofocus", "vbnull", "vbnullchar", "vbnullstring", "vbobject", "vbobjecterror", "vbok",
    "vbokcancel", "vbokonly", "vbpropercase", "vbqueryclose", "vbquestion", "vbreadonly", "vbred",
    "vbretry", "vbretrycancel", "vbsaturday", "vbscrollbars", "vbset", "vbshortdate", "vbshorttime",
    "vbsingle", "vbstrconv", "vbstring", "vbsunday", "vbsystem", "vbsystemmodal", "vbtab",
    "vbtextcompare", "vbthursday", "vbtitlebartext", "vbtristate", "vbtrue", "vbtuesday",
    "vbunicode", "vbuppercase", "vbusedefault", "vbuserdefinedtype", "vbusesystem",
    "vbusesystemdayofweek", "vbvariant", "vbvartype", "vbverticaltab", "vbvolume", "vbwednesday",
    "vbwhite", "vbwide", "vbwindowbackground", "vbwindowframe", "vbwindowtext", "vbyellow", "vbyes",
    "vbyesno", "vbyesnocancel", "weekday", "weekdayname", "weekdayname$", "width", "year",
})

# The members of each VBA module and enum, and of the hidden Global class.
VBA_LIBRARY_CONTAINERS: Mapping[str, frozenset[str]] = {
    "_hiddenmodule": frozenset({
        "array", "input", "input$", "inputb", "inputb$", "objptr", "strptr", "varptr", "width",
    }),
    "colorconstants": frozenset({
        "vbblack", "vbblue", "vbcyan", "vbgreen", "vbmagenta", "vbred", "vbwhite", "vbyellow",
    }),
    "constants": frozenset({
        "vbback", "vbcr", "vbcrlf", "vbformfeed", "vblf", "vbnewline", "vbnullchar", "vbnullstring",
        "vbobjecterror", "vbtab", "vbverticaltab",
    }),
    "conversion": frozenset({
        "cbool", "cbyte", "ccur", "cdate", "cdbl", "cdec", "cint", "clng", "clnglng", "clngptr",
        "csng", "cstr", "cstr$", "cvar", "cvdate", "cverr", "error", "error$", "fix", "hex", "hex$",
        "int", "macid", "oct", "oct$", "str", "str$", "val",
    }),
    "datetime": frozenset({
        "calendar", "date", "date$", "dateadd", "datediff", "datepart", "dateserial", "datevalue",
        "day", "hour", "minute", "month", "now", "second", "time", "time$", "timer", "timeserial",
        "timevalue", "weekday", "year",
    }),
    "filesystem": frozenset({
        "chdir", "chdrive", "curdir", "curdir$", "dir", "dir$", "eof", "fileattr", "filecopy",
        "filedatetime", "filelen", "freefile", "getattr", "kill", "loc", "lof", "mkdir", "reset",
        "rmdir", "seek", "setattr",
    }),
    "financial": frozenset({
        "ddb", "fv", "ipmt", "irr", "mirr", "nper", "npv", "pmt", "ppmt", "pv", "rate", "sln",
        "syd",
    }),
    "formshowconstants": frozenset({
        "vbmodal", "vbmodeless",
    }),
    "global": frozenset({
        "load", "unload", "userforms",
    }),
    "information": frozenset({
        "erl", "err", "imestatus", "isarray", "isdate", "isempty", "iserror", "ismissing", "isnull",
        "isnumeric", "isobject", "qbcolor", "rgb", "typename", "typename$", "vartype",
    }),
    "interaction": frozenset({
        "appactivate", "beep", "callbyname", "choose", "command", "command$", "createobject",
        "deletesetting", "doevents", "environ", "environ$", "getallsettings", "getobject",
        "getsetting", "getsetting$", "iif", "inputbox", "inputbox$", "macscript", "macscript$",
        "msgbox", "partition", "savesetting", "sendkeys", "shell", "switch",
    }),
    "keycodeconstants": frozenset({
        "vbkey0", "vbkey1", "vbkey2", "vbkey3", "vbkey4", "vbkey5", "vbkey6", "vbkey7", "vbkey8",
        "vbkey9", "vbkeya", "vbkeyadd", "vbkeyb", "vbkeyback", "vbkeyc", "vbkeycancel",
        "vbkeycapital", "vbkeyclear", "vbkeycontrol", "vbkeyd", "vbkeydecimal", "vbkeydelete",
        "vbkeydivide", "vbkeydown", "vbkeye", "vbkeyend", "vbkeyescape", "vbkeyexecute", "vbkeyf",
        "vbkeyf1", "vbkeyf10", "vbkeyf11", "vbkeyf12", "vbkeyf13", "vbkeyf14", "vbkeyf15",
        "vbkeyf16", "vbkeyf2", "vbkeyf3", "vbkeyf4", "vbkeyf5", "vbkeyf6", "vbkeyf7", "vbkeyf8",
        "vbkeyf9", "vbkeyg", "vbkeyh", "vbkeyhelp", "vbkeyhome", "vbkeyi", "vbkeyinsert", "vbkeyj",
        "vbkeyk", "vbkeyl", "vbkeylbutton", "vbkeyleft", "vbkeym", "vbkeymbutton", "vbkeymenu",
        "vbkeymultiply", "vbkeyn", "vbkeynumlock", "vbkeynumpad0", "vbkeynumpad1", "vbkeynumpad2",
        "vbkeynumpad3", "vbkeynumpad4", "vbkeynumpad5", "vbkeynumpad6", "vbkeynumpad7",
        "vbkeynumpad8", "vbkeynumpad9", "vbkeyo", "vbkeyp", "vbkeypagedown", "vbkeypageup",
        "vbkeypause", "vbkeyprint", "vbkeyq", "vbkeyr", "vbkeyrbutton", "vbkeyreturn", "vbkeyright",
        "vbkeys", "vbkeyselect", "vbkeyseparator", "vbkeyshift", "vbkeysnapshot", "vbkeyspace",
        "vbkeysubtract", "vbkeyt", "vbkeytab", "vbkeyu", "vbkeyup", "vbkeyv", "vbkeyw", "vbkeyx",
        "vbkeyy", "vbkeyz",
    }),
    "math": frozenset({
        "abs", "atn", "cos", "exp", "log", "randomize", "rnd", "round", "sgn", "sin", "sqr", "tan",
    }),
    "strings": frozenset({
        "asc", "ascb", "ascw", "chr", "chr$", "chrb", "chrb$", "chrw", "chrw$", "filter", "format",
        "format$", "formatcurrency", "formatcurrency$", "formatdatetime", "formatdatetime$",
        "formatnumber", "formatnumber$", "formatpercent", "formatpercent$", "instr", "instrb",
        "instrrev", "join", "join$", "lcase", "lcase$", "left", "left$", "leftb", "leftb$", "len",
        "lenb", "ltrim", "ltrim$", "mid", "mid$", "midb", "midb$", "monthname", "monthname$",
        "replace", "replace$", "right", "right$", "rightb", "rightb$", "rtrim", "rtrim$", "space",
        "space$", "split", "strcomp", "strconv", "string", "string$", "strreverse", "strreverse$",
        "trim", "trim$", "ucase", "ucase$", "weekdayname", "weekdayname$",
    }),
    "systemcolorconstants": frozenset({
        "vb3ddkshadow", "vb3dface", "vb3dhighlight", "vb3dlight", "vb3dshadow", "vbactiveborder",
        "vbactivetitlebar", "vbapplicationworkspace", "vbbuttonface", "vbbuttonshadow",
        "vbbuttontext", "vbdesktop", "vbgraytext", "vbhighlight", "vbhighlighttext",
        "vbinactiveborder", "vbinactivecaptiontext", "vbinactivetitlebar", "vbinfobackground",
        "vbinfotext", "vbmenubar", "vbmenutext", "vbmsgbox", "vbmsgboxtext", "vbscrollbars",
        "vbtitlebartext", "vbwindowbackground", "vbwindowframe", "vbwindowtext",
    }),
    "vbappwinstyle": frozenset({
        "vbhide", "vbmaximizedfocus", "vbminimizedfocus", "vbminimizednofocus", "vbnormalfocus",
        "vbnormalnofocus",
    }),
    "vbcalendar": frozenset({
        "vbcalgreg", "vbcalhijri",
    }),
    "vbcalltype": frozenset({
        "vbget", "vblet", "vbmethod", "vbset",
    }),
    "vbcomparemethod": frozenset({
        "vbbinarycompare", "vbdatabasecompare", "vbtextcompare",
    }),
    "vbdatetimeformat": frozenset({
        "vbgeneraldate", "vblongdate", "vblongtime", "vbshortdate", "vbshorttime",
    }),
    "vbdayofweek": frozenset({
        "vbfriday", "vbmonday", "vbsaturday", "vbsunday", "vbthursday", "vbtuesday",
        "vbusesystemdayofweek", "vbwednesday",
    }),
    "vbfileattribute": frozenset({
        "vbalias", "vbarchive", "vbdirectory", "vbhidden", "vbnormal", "vbreadonly", "vbsystem",
        "vbvolume",
    }),
    "vbfirstweekofyear": frozenset({
        "vbfirstfourdays", "vbfirstfullweek", "vbfirstjan1", "vbusesystem",
    }),
    "vbimestatus": frozenset({
        "vbimealphadbl", "vbimealphasng", "vbimedisable", "vbimehiragana", "vbimekatakanadbl",
        "vbimekatakanasng", "vbimemodealpha", "vbimemodealphafull", "vbimemodedisable",
        "vbimemodehangul", "vbimemodehangulfull", "vbimemodehiragana", "vbimemodekatakana",
        "vbimemodekatakanahalf", "vbimemodenocontrol", "vbimemodeoff", "vbimemodeon", "vbimenoop",
        "vbimeoff", "vbimeon",
    }),
    "vbmsgboxresult": frozenset({
        "vbabort", "vbcancel", "vbignore", "vbno", "vbok", "vbretry", "vbyes",
    }),
    "vbmsgboxstyle": frozenset({
        "vbabortretryignore", "vbapplicationmodal", "vbcritical", "vbdefaultbutton1",
        "vbdefaultbutton2", "vbdefaultbutton3", "vbdefaultbutton4", "vbexclamation",
        "vbinformation", "vbmsgboxhelpbutton", "vbmsgboxright", "vbmsgboxrtlreading",
        "vbmsgboxsetforeground", "vbokcancel", "vbokonly", "vbquestion", "vbretrycancel",
        "vbsystemmodal", "vbyesno", "vbyesnocancel",
    }),
    "vbqueryclose": frozenset({
        "vbapptaskmanager", "vbappwindows", "vbformcode", "vbformcontrolmenu", "vbformmdiform",
    }),
    "vbstrconv": frozenset({
        "vbfromunicode", "vbhiragana", "vbkatakana", "vblowercase", "vbnarrow", "vbpropercase",
        "vbunicode", "vbuppercase", "vbwide",
    }),
    "vbtristate": frozenset({
        "vbfalse", "vbtrue", "vbusedefault",
    }),
    "vbvartype": frozenset({
        "vbarray", "vbboolean", "vbbyte", "vbcurrency", "vbdataobject", "vbdate", "vbdecimal",
        "vbdouble", "vbempty", "vberror", "vbinteger", "vblong", "vblonglong", "vbnull", "vbobject",
        "vbsingle", "vbstring", "vbuserdefinedtype", "vbvariant",
    }),
}

# ErrObject's properties with a Get and no Let: `VBA.Err.LastDllError = 5` does not compile.
VBA_ERR_READ_ONLY: frozenset[str] = frozenset({
    "lastdllerror",
})
