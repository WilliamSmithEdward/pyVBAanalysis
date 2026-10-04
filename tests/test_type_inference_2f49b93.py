"""The typeInference.ts / memberAccess.ts behavior XLIDE added between 67b844d and
2f49b93: DefType environments, String Consts, the known-value views and walk
starts, function results, ByRef element exactness, object/value arguments, and the
member-signature helpers. Ports of upstream's analyzerFactCacheLifetimes and
analyzerPerformance cases where they test these exports directly."""

from __future__ import annotations

from pyvbaanalysis.completion.member_access import (
    is_late_bound_type_key,
    member_takes_own_arguments,
    preceded_by_member_access_dot,
    signature_declares_parameters,
)
from pyvbaanalysis.conditional import create_conditional_activity_tracker
from pyvbaanalysis.diagnostics.argument_inference import incompatibility_reason, infer_argument_type
from pyvbaanalysis.diagnostics.call_extraction import CallableParamType, InferredArgumentType
from pyvbaanalysis.diagnostics.known_locals import known_local_literal_values
from pyvbaanalysis.lexer.token_kinds import TokenKind, VbaToken
from pyvbaanalysis.lexer.tokenize import tokenize
from pyvbaanalysis.parser.nodes import AssignmentNode, ProcedureNode, Span
from pyvbaanalysis.parser.parse_module import parse_module
from pyvbaanalysis.symbols.build_module_symbols import BuildModuleSymbolsOptions, build_module_symbols
from pyvbaanalysis.symbols.symbol_model import ModuleSymbolKind, ModuleSymbols
from pyvbaanalysis.types.type_inference import (
    SourceDeclaredType,
    by_ref_variable_type_mismatch,
    def_type_of,
    defaulted_straight_line,
    function_result_for,
    known_local_literal_values_at,
    parse_runtime_param_type,
    procedure_symbol_for,
    source_binding_type_resolvers,
    split_signature_top_level,
    string_constants_in_scope,
    type_environment_for,
    unreachable_statements_in,
)


def _symbols(source: str) -> ModuleSymbols:
    return build_module_symbols("M", ModuleSymbolKind.STANDARD, source)


def _procedures(source: str) -> list[ProcedureNode]:
    return [m for m in parse_module(source).members if isinstance(m, ProcedureNode)]


def _tokens(text: str) -> list[VbaToken]:
    return [t for t in tokenize(text) if t.kind is not TokenKind.NEWLINE]


def test_def_type_types_untyped_locals_parameters_and_results() -> None:
    source = "DefLng A-Z\nFunction F(p)\nDim i\nDim v As Variant\nF = p\nEnd Function\n"
    symbols = _symbols(source)
    proc = _procedures(source)[0]
    env = type_environment_for(symbols, proc)
    assert def_type_of(symbols, "i") == "Long"
    assert env["i"] == "Long"
    assert env["p"] == "Long"
    assert env["f"] == "Long"
    assert env["v"] == "Variant"


def test_def_type_variant_and_decimal_give_no_type() -> None:
    symbols = _symbols("DefVar A-Z\nSub S()\nEnd Sub\n")
    assert def_type_of(symbols, "x") is None


def test_string_consts_in_scope_and_their_conversion() -> None:
    source = 'Const K = "abc"\nConst N = 5\nSub S()\nConst L = "x"\nDim y As Long\ny = K\nEnd Sub\n'
    symbols = _symbols(source)
    proc = _procedures(source)[0]
    assert string_constants_in_scope(symbols, proc) == {"k": "abc", "l": "x"}
    resolvers = source_binding_type_resolvers(symbols, procedure_symbol_for(symbols, proc), None)
    declared = resolvers.resolve_expression_type("K")
    assert declared.string_value == "abc"
    actual = infer_argument_type(
        _tokens("K"), 0, type_environment_for(symbols, proc), {}, None,
        resolvers.resolve_expression_type, resolvers.resolve_qualified_expression_type,
    )
    assert actual is not None
    assert actual.label == "constant 'K' (\"abc\")"
    assert incompatibility_reason("Long", actual) == (
        "This string literal cannot be converted to a numeric value. "
        "This will raise Run-time error '13': Type mismatch."
    )


def test_string_overflowing_a_numeric_parameter_reports_its_value() -> None:
    actual = InferredArgumentType(type_="String", label='String literal "&H10000"', span=Span(0, 1), string_value="&H10000")
    assert incompatibility_reason("Integer", actual) == (
        'The string "&H10000" converts to 65536, outside the Integer range -32768 to 32767. '
        "This will raise Run-time error '6': Overflow."
    )


def test_float_literal_overflow_is_rounded_half_to_even() -> None:
    actual = infer_argument_type(_tokens("3000000000#"), 0, {}, {})
    assert actual is not None and actual.float_value == 3000000000.0
    assert incompatibility_reason("Long", actual) == (
        "The numeric literal 3000000000# is outside the Long range -2147483648 to 2147483647. "
        "This will raise Run-time error '6': Overflow."
    )
    rounded = infer_argument_type(_tokens("32767.5"), 0, {}, {})
    assert rounded is not None
    assert incompatibility_reason("Integer", rounded) == (
        "The numeric literal 32767.5, which VBA rounds to 32768, is outside the Integer range "
        "-32768 to 32767. This will raise Run-time error '6': Overflow."
    )


def test_byref_array_element_must_match_the_parameter() -> None:
    param = CallableParamType(name="n", type_="Long", by_ref=True)

    def resolve(name: str) -> SourceDeclaredType:
        from pyvbaanalysis.symbols.symbol_model import VbaSymbolKind

        return SourceDeclaredType(resolved=True, as_type="Variant", kind=VbaSymbolKind.LOCAL_VARIABLE, is_array=True)

    mismatch = by_ref_variable_type_mismatch(param, _tokens("a(1)"), 0, {}, resolve)
    assert mismatch is not None
    assert (mismatch.name, mismatch.actual) == ("a(...)", "Variant")


def test_runtime_param_type_reads_unicode_names() -> None:
    parsed = parse_runtime_param_type("[ByVal Größe As Прибор]")
    assert parsed is not None
    assert (parsed.name, parsed.type_, parsed.optional) == ("Größe", "Прибор", True)
    assert split_signature_top_level('[s As String = ")"], [n As Long]') == ['[s As String = ")"]', " [n As Long]"]


def test_known_locals_variant_empty_byte_true_and_whole_rounding() -> None:
    source = (
        "Sub S()\nDim v\nDim b As Byte\nDim l As Long\nDim d As Long\n"
        "b = True\nl = 4.5\nd = d + 0\nEnd Sub\n"
    )
    symbols = _symbols(source)
    proc = _procedures(source)[0]
    known = known_local_literal_values(source, proc, symbols, None)
    assert (known["v"].kind, known["v"].value) == ("empty", 0)
    assert known["b"].value == 255
    assert known["l"].value == 4
    assert (known["d"].value, known["d"].origin) == (0, "default")


def test_values_at_each_statement_are_shared_and_invalidated() -> None:
    source = (
        "Option Explicit\n#If True Then\nSub Run()\nDim x As Long\nx = 2\nDebug.Print x\n"
        "x = 4\nDebug.Print x\nEnd Sub\n#End If"
    )
    mod = parse_module(source)
    proc = next(m for m in mod.members if isinstance(m, ProcedureNode))
    symbols = _symbols(source)
    activity = create_conditional_activity_tracker(mod)
    at = known_local_literal_values_at(source, proc, symbols, activity)
    assert known_local_literal_values_at(source, proc, symbols, activity) is at
    assert at(proc.body[2])["x"].value == 2
    assert at(proc.body[4])["x"].value == 4
    assert at(proc.body[2])["x"].value == 2
    assert known_local_literal_values_at(source, proc, symbols, create_conditional_activity_tracker(mod)) is not at
    assert known_local_literal_values_at(source, proc, _symbols(source), activity) is not at
    assert known_local_literal_values_at(source + "\n", proc, symbols, activity) is not at


def test_walk_starts_reachability_and_calls_follow_their_symbols() -> None:
    source = (
        "Const Limit As Long = 0\nFunction F(ByVal n As Long) As Long\nIf Limit = 0 Then Exit Function\n"
        "F = Limit + n\nEnd Function\nFunction G(ByVal n As Long) As Long\nG = Limit + n\nEnd Function"
    )
    module = parse_module(source)
    proc, called = _procedures(source)
    base = build_module_symbols("M", ModuleSymbolKind.STANDARD, source, BuildModuleSymbolsOptions(parsed_module=module))
    from dataclasses import replace

    changed = replace(
        base,
        root=replace(
            base.root,
            children=[
                replace(symbol, default_raw="1") if symbol.kind.value == "constant" else symbol
                for symbol in base.root.children or []
            ],
        ),
    )
    result_node = next(node for node in proc.body if isinstance(node, AssignmentNode))
    for symbols, expected in ((base, "0"), (changed, "2"), (base, "0")):
        dead = unreachable_statements_in(source, proc, symbols, None)
        assert (id(result_node) in dead) is (expected == "0")
        assert unreachable_statements_in(source, proc, symbols, None) is dead
        call = function_result_for(source, called, symbols, None, [_tokens("1")])
        assert call is not None
        assert "".join(tok.raw_text for tok in call) == ("1" if expected == "0" else "2")
        assert function_result_for(source, called, symbols, None, [_tokens("1")]) is call
        walk = defaulted_straight_line(source, proc, symbols, None)
        held = (walk.get(id(result_node)) or {}).get("limit")
        assert ("".join(tok.raw_text for tok in held) if held is not None else None) == (
            None if expected == "0" else "1"
        )


def test_function_result_runs_its_branches_for_the_arguments() -> None:
    source = (
        "Function Sign1(ByVal n As Long) As Long\nIf n > 0 Then Sign1 = 1 Else Sign1 = 0\n"
        "End Function\nFunction Twice(ByVal n As Long) As Long\nTwice = n * 2\nEnd Function\n"
    )
    symbols = _symbols(source)
    sign1, twice = _procedures(source)
    negative = function_result_for(source, sign1, symbols, None, [_tokens("-1")])
    assert negative is not None and [tok.raw_text for tok in negative] == ["0"]
    doubled = function_result_for(source, twice, symbols, None, [_tokens("21")])
    assert doubled is not None and [tok.raw_text for tok in doubled] == ["42"]


def test_member_signature_helpers() -> None:
    assert member_takes_own_arguments("Range(Index As Variant) As ShapeRange")
    assert not member_takes_own_arguments("Range(Index As Variant) As Object")
    assert not member_takes_own_arguments("Placeholders() As Shapes")
    assert signature_declares_parameters("Item([Index]) As Shape")
    assert not signature_declares_parameters("Count() As Long")
    assert is_late_bound_type_key("union:Excel.Worksheet")
    assert not is_late_bound_type_key("Excel.Worksheet")
    assert preceded_by_member_access_dot("x = ws. _\n  Name", len("x = ws. _\n  "))
    assert not preceded_by_member_access_dot("x = Name", 4)


def _validate_call_statement(source: str, line: str) -> list[tuple[str, str]]:
    """Run validate_argument_types on the call statement `line` of the module's last
    procedure, as the argument-types rule does."""
    from pyvbaanalysis.completion.member_access import MemberCompletionContext
    from pyvbaanalysis.diagnostics.argument_inference import validate_argument_types
    from pyvbaanalysis.diagnostics.call_extraction import extract_call
    from pyvbaanalysis.diagnostics.callable_signatures import callable_type_signatures_for, source_name_scope_for

    symbols = _symbols(source)
    proc = _procedures(source)[-1]
    start = source.index(line)
    call = extract_call(source, Span(start, start + len(line)))
    assert call is not None
    resolvers = source_binding_type_resolvers(symbols, procedure_symbol_for(symbols, proc), None)
    pushed: list[tuple[str, str]] = []
    validate_argument_types(
        call,
        type_environment_for(symbols, proc),
        callable_type_signatures_for(symbols, None),
        source_name_scope_for(symbols, proc, None),
        lambda rule, message, span, data=None: pushed.append((rule, message)),
        resolvers.resolve_expression_type,
        resolvers.resolve_qualified_expression_type,
        source=source,
        member_ctx=MemberCompletionContext(),
    )
    return pushed


def test_cverr_into_a_long_parameter() -> None:
    source = "Public Sub NeedsLong(ByVal value As Long)\nEnd Sub\n\nPublic Sub Main()\n    NeedsLong CVErr(2015)\nEnd Sub\n"
    assert _validate_call_statement(source, "NeedsLong CVErr(2015)") == [
        (
            "argumentTypeMismatch",
            "Argument 'value' of 'NeedsLong' expects Long, but got CVErr(...) Error Variant. An Error "
            "Variant cannot be coerced to this scalar type. This will raise Run-time error '13': "
            "Type mismatch.",
        )
    ]


def test_an_operator_on_null_into_a_typed_parameter() -> None:
    source = "Sub TakeL(ByVal n As Long)\nEnd Sub\nSub Main()\n    TakeL 1 + Null\nEnd Sub\n"
    assert _validate_call_statement(source, "TakeL 1 + Null") == [
        (
            "argumentTypeMismatch",
            "Argument 'n' of 'TakeL' expects Long, but '1 + Null' is Null: an operator on Null gives "
            "Null. Null cannot be coerced to this scalar type. This will raise Run-time error '94': "
            "Invalid use of Null.",
        )
    ]


def test_objects_and_values_in_the_wrong_parameter() -> None:
    source = (
        "Sub TakeL(ByVal n As Long)\nEnd Sub\nSub TakeC(c As Collection)\nEnd Sub\n"
        "Sub Main()\n    TakeL Nothing\n    TakeC 5\nEnd Sub\n"
    )
    assert _validate_call_statement(source, "TakeL Nothing") == [
        (
            "argumentObjectTypeMismatch",
            "Argument 'n' of 'TakeL' expects Long, but got Nothing. This is a VBE compile error: "
            "Invalid use of object.",
        )
    ]
    assert _validate_call_statement(source, "TakeC 5") == [
        (
            "argumentObjectTypeMismatch",
            "Argument 'c' of 'TakeC' expects Collection, but got numeric literal 5. An object "
            "parameter takes an object. This is a VBE compile error: Type mismatch.",
        )
    ]
