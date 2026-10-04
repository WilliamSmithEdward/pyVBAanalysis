"""Block headers, loop counters, Open numbers and the block-entering walks (XLIDE 2f49b93).

Covers the foundations the 2f49b93 sync ported: blockHeaders.ts,
loopCounters.ts, openedFileNumbers.ts, statementWalk.ts, the new parts of
walker.ts, dataflow.ts's walkEnteringBlocks, procedureLabels.ts's label
declarations and jump targets, and callExtraction.ts's new arity checks.
"""

from __future__ import annotations

from pyvbaanalysis.diagnostics.block_headers import (
    block_header_leaves,
    block_header_statements,
    is_loop_block,
    select_arms,
)
from pyvbaanalysis.diagnostics.call_extraction import (
    CallableParamType,
    CallableTypeSignature,
    call_then_index,
    extract_call,
    validate_arity,
)
from pyvbaanalysis.diagnostics.dataflow import BlockEnteringState, walk_entering_blocks
from pyvbaanalysis.diagnostics.loop_counters import (
    CounterValue,
    check_each_counter_pass,
    counter_text,
    counter_value,
    loop_counters_at,
    numeric_counter_passes,
)
from pyvbaanalysis.diagnostics.opened_file_numbers import (
    merge_opened_file_numbers,
    opened_file_numbers_in,
)
from pyvbaanalysis.diagnostics.walker import (
    bare_assignment_target,
    block_header_line_span,
    for_each_statement_with_headers,
    raw_expression_tokens,
    walk_procedure_statements,
)
from pyvbaanalysis.flow.procedure_labels import (
    collect_procedure_label_declarations,
    jump_target_label_declaration,
)
from pyvbaanalysis.flow.procedure_unstructured import procedure_has_unstructured_flow
from pyvbaanalysis.lexer.token_kinds import TokenKind
from pyvbaanalysis.parser.nodes import (
    BodyNode,
    DoBlockNode,
    ForBlockNode,
    IfBlockNode,
    LeafStatementNode,
    ProcedureNode,
    SelectBlockNode,
    Span,
    WithBlockNode,
    is_leaf_statement,
    iter_body_nodes,
)
from pyvbaanalysis.parser.parse_module import parse_module
from pyvbaanalysis.parser.statement_walk import for_each_statement


def _procedure(source: str, index: int = 0) -> ProcedureNode:
    procs = [m for m in parse_module(source).members if isinstance(m, ProcedureNode)]
    return procs[index]


def _text(source: str, span: Span) -> str:
    return source[span.start : span.end]


def _first(body: list[BodyNode], kind: type) -> BodyNode:
    return next(node for node in iter_body_nodes(body) if isinstance(node, kind))


# --- blockHeaders ---------------------------------------------------------


def test_block_header_statements_give_the_header_and_a_do_footer() -> None:
    source = (
        "Sub S()\n    For i = 1 To n: x = 1: Next\n    Do\n        y = 1\n"
        "    Loop While y < 3 ' done\nEnd Sub\n"
    )
    body = _procedure(source).body
    for_before, for_after = block_header_statements(source, _first(body, ForBlockNode))
    assert for_before is not None and for_before.raw == "For i = 1 To n"
    assert for_after is None
    do_headers = block_header_statements(source, _first(body, DoBlockNode))
    assert do_headers.before is not None and do_headers.before.raw == "Do"
    assert do_headers.after is not None and do_headers.after.raw == "Loop While y < 3"


def test_block_header_leaves_of_an_if_are_its_condition_lines() -> None:
    source = "Sub S()\n    If a Then\n    ElseIf b Then\n    Else\n    End If\nEnd Sub\n"
    node = _first(_procedure(source).body, IfBlockNode)
    assert [stmt.raw for stmt in block_header_leaves(source, node)] == [
        "If a Then",
        "ElseIf b Then",
    ]
    assert block_header_statements(source, node) == (None, None)


def test_select_arms_and_loop_blocks() -> None:
    source = (
        "Sub S()\n    Select Case n\n    Case 1\n        a = 1\n    Case Else\n"
        "        b = 2\n    End Select\n    With c\n    End With\nEnd Sub\n"
    )
    body = _procedure(source).body
    select = _first(body, SelectBlockNode)
    assert isinstance(select, SelectBlockNode)
    arms = select_arms(source, select.body)
    assert [[_text(source, node.span) for node in arm] for arm in arms] == [
        ["Case 1", "a = 1"],
        ["Case Else", "b = 2"],
    ]
    assert not is_loop_block(select)
    assert not is_loop_block(_first(body, WithBlockNode))


# --- openedFileNumbers ----------------------------------------------------


def test_opened_file_numbers() -> None:
    source = (
        "Sub S()\n    Open f For Input As #1\n10  Open g For Output As 2\n"
        "    If a Then Open h For Append As #3\n    x = Open\nEnd Sub\n"
    )
    opened = opened_file_numbers_in(source)
    assert opened.numbers == {1, 2, 3}
    assert not opened.any
    free = opened_file_numbers_in("Sub T()\n    Open f For Input As #n\nEnd Sub\n")
    assert free.any and free.numbers == frozenset()
    merged = merge_opened_file_numbers([opened, free])
    assert merged.any and merged.numbers == {1, 2, 3}


# --- loopCounters ---------------------------------------------------------


def _leaf_containing(proc: ProcedureNode, source: str, text: str) -> LeafStatementNode:
    for node in iter_body_nodes(proc.body):
        if is_leaf_statement(node) and text in _text(source, node.span):
            return node
    raise AssertionError(text)


def test_for_counter_bounds_and_passes() -> None:
    source = (
        "Sub S()\n    Dim s As String\n    For i = 0 To Len(s) - 1\n"
        "        Debug.Print Mid$(s, i, 1)\n    Next\nEnd Sub\n"
    )
    proc = _procedure(source)
    counters = loop_counters_at(source, proc.body, None)
    at = counters.get(_leaf_containing(proc, source, "Mid$"))
    assert at is not None
    counter = at["i"]
    assert counter.loop == "For" and counter.step == 1
    assert counter.first == CounterValue(offset=0)
    assert counter.last is not None and counter_text(counter.last) == "Len(s) - 1"
    passes = numeric_counter_passes(counter, lambda atom, _counter: 3)
    assert [(p.pass_, p.value) for p in passes] == [("first", 0), ("last", 2)]
    # The numbers show the loop never runs: no pass.
    assert numeric_counter_passes(counter, lambda atom, _counter: 0) == []


def test_for_counter_with_a_step_ends_on_its_last_pass() -> None:
    source = "Sub S()\n    For i = 10 To 1 Step -2\n        a(i) = 0\n    Next\nEnd Sub\n"
    proc = _procedure(source)
    at = loop_counters_at(source, proc.body, None).get(_leaf_containing(proc, source, "a(i)"))
    assert at is not None
    assert at["i"].last == CounterValue(offset=2)


def test_a_loop_that_writes_its_counter_has_none() -> None:
    source = "Sub S()\n    For i = 1 To 3\n        i = i + 1\n    Next\nEnd Sub\n"
    assert len(loop_counters_at(source, _procedure(source).body, None)) == 0


def test_stepped_do_counter() -> None:
    source = (
        "Sub S()\n    i = 1\n    Do While i <= 4\n        a(i) = 0\n        i = i + 1\n"
        "    Loop\nEnd Sub\n"
    )
    proc = _procedure(source)
    counters = loop_counters_at(source, proc.body, None)
    at = counters.get(_leaf_containing(proc, source, "a(i)"))
    assert at is not None
    counter = at["i"]
    assert counter.loop == "Do"
    assert (counter.first, counter.last) == (CounterValue(offset=1), CounterValue(offset=4))
    # The increment is not a statement of the counted body.
    assert counters.get(_leaf_containing(proc, source, "i = i + 1")) is None


def test_counter_value_atoms() -> None:
    def value(text: str) -> CounterValue | None:
        return counter_value(raw_expression_tokens(text))

    ubound = value("UBound(a, 2) + 1")
    assert ubound is not None and ubound.atom is not None
    assert (ubound.atom.kind, ubound.atom.dimension, ubound.offset) == ("ubound", 2, 1)
    assert counter_text(ubound) == "UBound(a, 2) + 1"
    count = value("c.Count - 1")
    assert count is not None and count.atom is not None and count.atom.kind == "count"
    assert value('Len("abc")') == CounterValue(offset=3)
    assert value("x.End(xlUp).Row") is None


def test_check_each_counter_pass_names_the_pass() -> None:
    source = "Sub S()\n    For i = 0 To 2\n        Debug.Print a(i)\n    Next\nEnd Sub\n"
    proc = _procedure(source)
    leaf = _leaf_containing(proc, source, "Debug.Print")
    counters = loop_counters_at(source, proc.body, None).get(leaf)
    found: list[str] = []

    def check(values, push) -> None:  # type: ignore[no-untyped-def]
        if values.get("i") == 2:
            push("rule", "boom", leaf.span)

    check_each_counter_pass(
        source,
        leaf.span,
        counters,
        lambda atom, counter: None,
        check,
        lambda rule, message, span, data=None: found.append(message),
    )
    assert found == ["On the last pass of the For loop, where 'i' is 2: boom"]


# --- procedureLabels ------------------------------------------------------


def test_a_line_carries_a_number_and_a_name_label() -> None:
    source = "Sub S()\n10 L1: x = 1\n    GoTo L1\nEnd Sub\n"
    labels = collect_procedure_label_declarations(source, _procedure(source))
    assert [label.key for label in labels] == ["line:10", "name:l1"]


def test_jump_target_labels_are_the_ones_a_jump_names() -> None:
    source = "Sub S()\n    GoTo L1\nL1:\n    x = 1\nL2:\n    y = 1\nEnd Sub\n"
    proc = _procedure(source)
    leaves = [node for node in iter_body_nodes(proc.body)]
    by_text = {_text(source, node.span).rstrip(":"): node for node in leaves}
    target = jump_target_label_declaration(source, by_text["L1"].span)
    assert target is not None and target.key == "name:l1"
    assert jump_target_label_declaration(source, by_text["L2"].span) is None
    assert procedure_has_unstructured_flow(source, proc)


def test_on_error_resume_next_alone_is_unstructured() -> None:
    source = "Sub S()\n    On Error Resume Next\n    x = 1\nEnd Sub\n"
    assert procedure_has_unstructured_flow(source, _procedure(source))
    plain = "Sub S()\n    If a Then\n        x = 1\n    End If\nEnd Sub\n"
    assert not procedure_has_unstructured_flow(plain, _procedure(plain))


# --- callExtraction -------------------------------------------------------


def _arity_messages(sig: CallableTypeSignature, statement: str) -> list[str]:
    source = f"Sub T()\n    {statement}\nEnd Sub\n"
    start = source.index(statement)
    call = extract_call(source, Span(start, start + len(statement)))
    assert call is not None
    found: list[str] = []
    validate_arity(source, sig, call, lambda rule, message, span, data=None: found.append(message))
    return found


def test_a_call_may_not_end_in_an_empty_argument() -> None:
    sig = CallableTypeSignature(
        name="Opt", params=[CallableParamType(name="a"), CallableParamType(name="b", optional=True)]
    )
    assert _arity_messages(sig, "Opt 1,") == [
        "The call to 'Opt' ends in an empty argument. This is a VBE compile error: Syntax error."
    ]


def test_named_arguments_and_a_param_array() -> None:
    many = CallableTypeSignature(
        name="Many",
        params=[CallableParamType(name="p0"), CallableParamType(name="rest", param_array=True)],
    )
    assert _arity_messages(many, "Many p0:=1") == [
        "'Many' has a ParamArray, 'rest', so no argument to it may be named. This is a VBE "
        "compile error: Argument in ParamArray may not be named."
    ]
    assert _arity_messages(many, "Many rest:=1") == [
        "'rest' is the ParamArray of 'Many', which may not be named. This is a VBE compile "
        "error: Argument in ParamArray may not be named."
    ]
    two = CallableTypeSignature(
        name="TTwo", params=[CallableParamType(name="a"), CallableParamType(name="b")]
    )
    assert _arity_messages(two, "TTwo a:=1") == [
        "Parameter 'b' of 'TTwo' is not Optional, and the call gives it no argument, by "
        "position or by name. This is a VBE compile error: Argument not optional."
    ]
    assert _arity_messages(two, "TTwo 1, b:=2") == []


def test_a_parameterless_function_indexes_its_result() -> None:
    arr = CallableTypeSignature(name="Arr", params=[], return_type="Variant", valued=True)
    source = "Sub T()\n    Arr 1\nEnd Sub\n"
    start = source.index("Arr 1")
    call = extract_call(source, Span(start, start + 5))
    assert call is not None and call_then_index(arr, call)
    typed = CallableTypeSignature(name="Arr", params=[], return_type="Long()", valued=True)
    assert not call_then_index(typed, call)
    assert _arity_messages(arr, "Arr 1") == []
    left = CallableTypeSignature(name="Left", params=[], return_type="String", valued=True)
    assert _arity_messages(left, "Left x, 1") == []
    assert len(_arity_messages(left, "Left x")) == 1


# --- walker ---------------------------------------------------------------


def test_raw_expression_tokens_reads_a_leading_date_literal() -> None:
    toks = raw_expression_tokens("#12/31/9999# + 1")
    assert toks[0].kind is TokenKind.DATE_LITERAL
    assert (toks[0].start, toks[0].end) == (0, 12)
    assert [tok.raw_text for tok in toks[1:]] == ["+", "1"]


def test_block_header_line_span_follows_continuations() -> None:
    source = "If a = 1 _\n    And b Then\n    x = 1\nEnd If\n"
    span = block_header_line_span(source, Span(0, len(source)))
    assert _text(source, span) == "If a = 1 _\n    And b Then"


def test_for_each_statement_with_headers_visits_block_lines() -> None:
    source = (
        "Sub S()\n    With c\n        .Add 1\n    End With\n    Do\n        n = n + 1\n"
        "    Loop Until n > 3\nEnd Sub\n"
    )
    seen: list[str] = []
    for_each_statement_with_headers(
        source, _procedure(source).body, lambda stmt: seen.append(_text(source, stmt.span))
    )
    assert seen == ["With c", ".Add 1", "Do", "n = n + 1", "Loop Until n > 3"]
    plain: list[str] = []
    for_each_statement(_procedure(source).body, lambda stmt: plain.append(_text(source, stmt.span)))
    assert plain == [".Add 1", "n = n + 1"]


def test_walk_procedure_statements_gives_headers_to_the_visitors_that_take_them() -> None:
    source = "Sub S()\n    If 10 / d > 1 Then\n        x = 1\n    End If\nEnd Sub\n"
    mod = parse_module(source)
    with_headers: list[str] = []
    without: list[str] = []
    walk_procedure_statements(
        mod,
        None,
        [
            lambda proc: lambda stmt: with_headers.append(_text(source, stmt.span)),
            lambda proc: lambda stmt: without.append(_text(source, stmt.span)),
        ],
        None,
        (source, [True, False]),
    )
    assert with_headers == ["If 10 / d > 1 Then", "x = 1"]
    assert without == ["x = 1"]


def test_walk_procedure_statements_skips_a_body_before_building_its_visitors() -> None:
    from pyvbaanalysis.diagnostics.walker import ProcedureWalkHooks

    source = "Sub A()\n    x = 1\nEnd Sub\nSub B()\n    y = 1\nEnd Sub\n"
    built: list[str] = []

    def visitor(proc: ProcedureNode):  # type: ignore[no-untyped-def]
        built.append(proc.name)
        return lambda stmt: None

    hooks = ProcedureWalkHooks(skip_body=lambda proc: proc.name == "A")
    walk_procedure_statements(parse_module(source), None, [visitor], hooks)
    assert built == ["B"]


# --- walkEnteringBlocks ---------------------------------------------------


def _known_at_reads(source: str, body: list[BodyNode]) -> dict[str, bool]:
    """Walk entering blocks with the set of names assigned a value; for each
    statement `r = x`, whether x is known there."""
    known: set[str] = set()
    out: dict[str, bool] = {}

    def target(stmt: LeafStatementNode) -> str | None:
        hit = bare_assignment_target(source, stmt.span)
        return hit[0].lower() if hit is not None else None

    def visit(node: BodyNode) -> None:
        hit = bare_assignment_target(source, node.span)
        if hit is None:
            return
        name, _span, value = hit
        if len(value) == 1 and value[0].kind is TokenKind.IDENTIFIER:
            out[name.lower()] = value[0].raw_text.lower() in known
        known.add(name.lower())

    def restore(state: set[str]) -> None:
        known.clear()
        known.update(state)

    def forget(names) -> None:  # type: ignore[no-untyped-def]
        known.difference_update(names)

    state: BlockEnteringState[set[str]] = BlockEnteringState(
        snapshot=lambda: set(known),
        restore=restore,
        forget=forget,
        touches=lambda stmt: [name] if (name := target(stmt)) else [],
    )
    walk_entering_blocks(source, body, lambda node: False, visit, state)
    return out


def test_walk_entering_blocks_carries_state_into_blocks() -> None:
    source = (
        "Sub S()\n    x = 1\n    If c Then\n        a = x\n    End If\n"
        "    For i = 1 To 2\n        If c Then\n            b = x\n        End If\n"
        "        x = 2\n    Next\n    z = x\n    w = q\nEnd Sub\n"
    )
    reads = _known_at_reads(source, _procedure(source).body)
    # Entered with x known; a block nested in a loop forgets what the loop
    # changes; after the loop, what it touched is unknown.
    assert reads == {"a": True, "b": False, "z": False, "w": False}


# --- explicit stacks --------------------------------------------------------

DEPTH = 1100


def test_block_walks_run_past_the_recursion_limit() -> None:
    source = (
        "Sub S()\n    x = 1\n"
        + "    If c Then\n" * DEPTH
        + "    y = x\n"
        + "    End If\n" * DEPTH
        + "    For i = 1 To 3\n" * DEPTH
        + "    a(i) = 0\n"
        + "    Next\n" * DEPTH
        + "End Sub\n"
    )
    proc = _procedure(source)
    assert _known_at_reads(source, proc.body) == {"y": True}
    counters = loop_counters_at(source, proc.body, None)
    at = counters.get(_leaf_containing(proc, source, "a(i)"))
    assert at is not None and at["i"].last == CounterValue(offset=3)
    seen: list[str] = []
    for_each_statement_with_headers(
        source, proc.body, lambda stmt: seen.append(_text(source, stmt.span))
    )
    # The For headers, then the statements, each once.
    assert seen.count("For i = 1 To 3") == DEPTH
    assert seen[-1] == "a(i) = 0"
