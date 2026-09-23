"""Doc comments that do not match their declaration.

Ported from xlide_vscode/src/analyzer/diagnostics/rules/docComments.ts.

A `'''` block written in XML, one holding any tag of the documented vocabulary,
documents the Sub, Function, Property, Declare or Event below it, and hovers and
call tips show it as the truth about that declaration. So once a procedure has one,
it has to describe the whole surface a caller sees:

- every parameter, by a <param> naming it, once, with something in it;
- the value a Function returns, by a <returns>, which a Sub, a Property Let or Set
  and an Event must not have, since they return nothing;
- every tag closed, and no tag but <param> given twice.

A property is its value, and its <summary> describes it: neither the value a
Property Get returns nor the one a Property Let or Set receives in its last
parameter needs a tag of its own, though either may have one. A block of plain text
is a note, not XML, and is left alone, and so is a row of apostrophes drawn above a
procedure.

A finding about a tag points at the tag; one about something the block leaves out
points at the declaration.
"""

from __future__ import annotations

import re
from collections.abc import Sequence
from dataclasses import dataclass
from typing import Literal, Union

from ...conditional import ConditionalActivityTracker
from ...docs.doc_comment import (
    DocBlockLine,
    DocTagOccurrence,
    detect_eol,
    leading_doc_lines,
    leading_whitespace,
    line_start_at,
    scan_doc_tags,
)
from ...parser.nodes import (
    DeclareNode,
    EventNode,
    ModuleNode,
    ParameterNode,
    ProcedureNode,
    ProcKind,
    Span,
)
from ..context import PushFn
from ..model import VbaDiagnosticData, VbaDocCommentFix, VbaEdit
from ..walker import active_module_members

_DocumentedMember = Union[ProcedureNode, DeclareNode, EventNode]

_PROCEDURE_LABELS = {
    ProcKind.SUB: "Sub",
    ProcKind.FUNCTION: "Function",
    ProcKind.PROPERTY_GET: "Property Get",
    ProcKind.PROPERTY_LET: "Property Let",
    ProcKind.PROPERTY_SET: "Property Set",
}

# A `disable-next-line` directive, which is about the line below it only.
_NEXT_LINE_DIRECTIVE_RE = re.compile(r"^\s*'+\s*@xlide-analysis-disable-next-line\b", re.IGNORECASE)

# Tags a doc comment has at most one of; the parser reads the first.
_SINGLE_TAGS = frozenset({"summary", "returns", "remarks", "example", "signature"})


@dataclass(slots=True)
class _Surface:
    """What a declaration shows its callers, which its doc comment describes."""

    # Sub, Function, Property Get, Property Let, Property Set or Event.
    label: str
    params: list[ParameterNode]
    returns: Literal["required", "optional", "none"]
    # A Property Let or Set's last parameter: the property's value.
    value_param: ParameterNode | None = None


def check_doc_comments(
    source: str, mod: ModuleNode, activity: ConditionalActivityTracker | None, push: PushFn
) -> None:
    """A procedure's XML doc comment describes every parameter and return value it
    has, and nothing it does not have."""
    for member in active_module_members(mod, activity):
        if not isinstance(member, (ProcedureNode, DeclareNode, EventNode)):
            continue
        lines = leading_doc_lines(source, member.span.start)
        tags = scan_doc_tags(lines) if lines else None
        if tags is not None:
            _check_member(_DocBlock(source, lines, tags), member, _surface_of(member), push)


def _surface_of(member: _DocumentedMember) -> _Surface:
    params = [param for param in member.params if param.name]
    if isinstance(member, ProcedureNode):
        setter = member.proc_kind in (ProcKind.PROPERTY_LET, ProcKind.PROPERTY_SET)
        last = member.params[-1] if member.params else None
        if member.proc_kind is ProcKind.FUNCTION:
            returns: Literal["required", "optional", "none"] = "required"
        elif member.proc_kind is ProcKind.PROPERTY_GET:
            returns = "optional"
        else:
            returns = "none"
        return _Surface(
            label=_PROCEDURE_LABELS[member.proc_kind],
            params=params,
            returns=returns,
            value_param=last if setter and last is not None and last.name else None,
        )
    if isinstance(member, DeclareNode):
        return _Surface(
            label="Function" if member.is_function else "Sub",
            params=params,
            returns="required" if member.is_function else "none",
        )
    return _Surface(label="Event", params=params, returns="none")


def _at(offset: int) -> Span:
    return Span(offset, offset)


def _check_member(block: _DocBlock, member: _DocumentedMember, surface: _Surface, push: PushFn) -> None:
    def report(rule: str, message: str, span: Span, fixes: Sequence[VbaDocCommentFix] = ()) -> None:
        push(rule, message, span, VbaDiagnosticData(doc_comment_fixes=tuple(fixes)) if fixes else None)

    # The tag each parameter is described by: the first naming it, as the call tip
    # reads it.
    param_tags: dict[str, DocTagOccurrence] = {}
    for tag in block.tags:
        if tag.tag == "param" and tag.name:
            param_tags.setdefault(tag.name.lower(), tag)
    by_name = {param.name.lower(): param for param in surface.params}
    undescribed = [param for param in surface.params if param.name.lower() not in param_tags]

    seen: set[str] = set()
    returns: DocTagOccurrence | None = None
    for tag in block.tags:
        if tag.tag == "param":
            first = not tag.name or param_tags.get(tag.name.lower()) is tag
        else:
            first = tag.tag not in seen
        seen.add(tag.tag)
        if tag.tag == "returns" and returns is None:
            returns = tag
        if tag.end is None:
            report("docTagUnclosed", f"This <{tag.tag}> is not closed; end it with </{tag.tag}>.", tag.open)
            continue
        empty = tag.text == "" and not tag.has_hints
        if not first and (tag.tag == "param" or tag.tag in _SINGLE_TAGS):
            report(
                "docTagDuplicate",
                f"The doc comment already describes parameter '{tag.name}'."
                if tag.tag == "param"
                else f"The doc comment already has a <{tag.tag}>, and only the first is shown.",
                tag.name_span if tag.name_span is not None else tag.open,
                [_remove_fix(block, tag, f"Remove the repeated <{tag.tag}>")],
            )
            continue
        if tag.tag == "param":
            if not tag.name or tag.name_span is None:
                report(
                    "docParamUnknown",
                    'This <param> has no name="..." to say which parameter it describes.',
                    tag.open,
                    [
                        *(
                            VbaDocCommentFix(
                                title=f"Name the <param> '{param.name}'",
                                is_preferred=len(undescribed) == 1,
                                edits=(
                                    VbaEdit(_at(tag.open.start + len("<param")), f' name="{param.name}"'),
                                ),
                            )
                            for param in undescribed
                        ),
                        _remove_fix(block, tag, "Remove the <param>"),
                    ],
                )
            elif tag.name.lower() not in by_name:
                name_span = tag.name_span
                report(
                    "docParamUnknown",
                    f"'{member.name}' has no parameter named '{tag.name}'.",
                    name_span,
                    [
                        *(
                            VbaDocCommentFix(
                                title=f"Rename the <param> to '{param.name}'",
                                is_preferred=len(undescribed) == 1,
                                edits=(VbaEdit(name_span, param.name),),
                            )
                            for param in undescribed
                        ),
                        _remove_fix(block, tag, f"Remove the <param> for '{tag.name}'"),
                    ],
                )
            elif empty:
                report(
                    "docParamMissing",
                    f"The <param> for '{by_name[tag.name.lower()].name}' is empty.",
                    tag.open,
                )
        elif tag.tag == "returns":
            if surface.returns == "none":
                report(
                    "docReturnsUnexpected",
                    f"{surface.label} '{member.name}' returns no value, but its doc comment describes one.",
                    tag.open,
                    [_remove_fix(block, tag, "Remove the <returns>")],
                )
            elif empty:
                report("docReturnsMissing", "The <returns> is empty.", tag.open)

    missing = [param for param in undescribed if param is not surface.value_param]
    add_all = (
        VbaDocCommentFix(
            title=f"Add the {len(missing)} missing <param> tags",
            edits=tuple(
                _merge_insertions(
                    [_add_param_edit(block, surface, param, param_tags) for param in missing]
                )
            ),
        )
        if len(missing) > 1
        else None
    )
    for param in missing:
        report(
            "docParamMissing",
            f"The doc comment does not describe parameter '{param.name}'.",
            param.name_span if param.name_span is not None else param.span,
            [
                VbaDocCommentFix(
                    title=f"Add a <param> for '{param.name}'",
                    is_preferred=True,
                    edits=(_add_param_edit(block, surface, param, param_tags),),
                ),
                *((add_all,) if add_all is not None else ()),
            ],
        )
    if surface.returns == "required" and returns is None:
        report(
            "docReturnsMissing",
            f"The doc comment does not describe what '{member.name}' returns.",
            member.name_span if member.name_span is not None else member.span,
            [VbaDocCommentFix(title="Add a <returns>", is_preferred=True, edits=(_add_returns_edit(block),))],
        )


def _remove_fix(block: _DocBlock, tag: DocTagOccurrence, title: str) -> VbaDocCommentFix:
    return VbaDocCommentFix(title=title, edits=(block.remove_tag(tag),))


def _add_param_edit(
    block: _DocBlock,
    surface: _Surface,
    param: ParameterNode,
    param_tags: dict[str, DocTagOccurrence],
) -> VbaEdit:
    """A new `<param>` goes among the others in signature order; with none to go by,
    after the summary, or before whatever follows the parameters."""
    content = f'<param name="{param.name}"></param>'
    index = next(i for i, candidate in enumerate(surface.params) if candidate is param)
    before_end: int | None = None
    after_start: int | None = None
    for i, candidate in enumerate(surface.params):
        tag = param_tags.get(candidate.name.lower())
        if tag is None or tag.end is None:
            continue
        if i < index:
            before_end = tag.end
        elif i > index and after_start is None:
            after_start = tag.open.start
    if before_end is not None:
        return block.insert_line_after(before_end, content)
    if after_start is not None:
        return block.insert_line_before(after_start, content)
    return block.insert_before_trailing_tags(content)


def _add_returns_edit(block: _DocBlock) -> VbaEdit:
    """A new `<returns>` goes after the last `<param>`, or where one would go."""
    params = [tag for tag in block.tags if tag.tag == "param" and tag.end is not None]
    last = params[-1] if params else None
    if last is not None and last.end is not None:
        return block.insert_line_after(last.end, "<returns></returns>")
    return block.insert_before_trailing_tags("<returns></returns>")


def _merge_insertions(edits: list[VbaEdit]) -> list[VbaEdit]:
    """Insertions at one offset become one edit, in the order given."""
    by_offset: dict[int, VbaEdit] = {}
    for edit in edits:
        merged = by_offset.get(edit.span.start)
        by_offset[edit.span.start] = (
            VbaEdit(merged.span, merged.new_text + edit.new_text) if merged is not None else edit
        )
    return sorted(by_offset.values(), key=lambda edit: edit.span.start)


class _DocBlock:
    """The `'''` block above one declaration, and the edits that change it."""

    __slots__ = ("_eol", "_source", "lines", "tags")

    def __init__(self, source: str, lines: list[DocBlockLine], tags: list[DocTagOccurrence]) -> None:
        self._source = source
        self.lines = lines
        self.tags = tags
        self._eol = detect_eol(source)

    def insert_line_before(self, offset: int, content: str) -> VbaEdit:
        line = self._line_at(offset)
        return VbaEdit(_at(line.directives_start), self._new_line(line, content))

    def insert_line_after(self, offset: int, content: str) -> VbaEdit:
        line = self._line_at(offset)
        return VbaEdit(_at(self._next_line_start(line)), self._new_line(line, content))

    def insert_before_trailing_tags(self, content: str) -> VbaEdit:
        """After the summary, else before the returns, remarks or example, else last."""
        summary = next((tag for tag in self.tags if tag.tag == "summary"), None)
        if summary is not None and summary.end is not None:
            return self.insert_line_after(summary.end, content)
        trailing = next((tag for tag in self.tags if tag.tag in ("returns", "remarks", "example")), None)
        if trailing is not None:
            return self.insert_line_before(trailing.open.start, content)
        written = [line for line in self.lines if line.text.strip() != ""]
        return self.insert_line_after(written[-1].start, content)

    def remove_tag(self, tag: DocTagOccurrence) -> VbaEdit:
        """The lines a tag has to itself go with it, and so does a `disable-next-line`
        right above them, which would otherwise move on to the line after. Otherwise
        just the tag goes."""
        end = tag.end if tag.end is not None else tag.open.end
        first = self._line_at(tag.open.start)
        last = self._line_at(end)
        before = self._source[first.text_start : tag.open.start]
        after = self._source[end : self._line_end(last)]
        if before.strip() != "" or after.strip() != "":
            return VbaEdit(Span(tag.open.start, end), "")
        start = first.start
        if first.directives_start < first.start:
            above = line_start_at(self._source, first.start - 1)
            if _NEXT_LINE_DIRECTIVE_RE.match(self._source[above : first.start]):
                start = above
        return VbaEdit(Span(start, self._next_line_start(last)), "")

    def _new_line(self, beside: DocBlockLine, content: str) -> str:
        """A `'''` line indented the way the line it goes next to is."""
        prefix = self._source[beside.start : beside.text_start] + leading_whitespace(beside.text)
        return f"{prefix}{content}{self._eol}"

    def _line_at(self, offset: int) -> DocBlockLine:
        found = self.lines[0]
        for line in self.lines:
            if line.start <= offset:
                found = line
        return found

    def _line_end(self, line: DocBlockLine) -> int:
        return line.text_start + len(line.text)

    def _next_line_start(self, line: DocBlockLine) -> int:
        newline = self._source.find("\n", self._line_end(line))
        return len(self._source) if newline < 0 else newline + 1
