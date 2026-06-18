"""VBA abstract syntax tree node definitions.

Ported from xlide_vscode/src/analyzer/parser/nodes.ts. Verified against MS-VBAL
v20250520 (sections 4.2, 5.1-5.4, 5.6, 3.4).

Design notes (the once-up-front decisions, agent.md Risk 2):
- Each node is a mutable @dataclass(slots=True), not frozen. The parser is
  error-tolerant and builds nodes incrementally (appends to `body`, sets `closed`,
  attaches `attributes`) exactly as the TypeScript object-literal-then-mutate code
  does; mutability keeps the port faithful.
- Dispatch is by isinstance (the Python idiom, UM-02). Each concrete node carries
  its discriminant as a KIND / EXPR_KIND ClassVar so the value still round-trips to
  the TS `kind` / `exprKind` string for the differential harness and debugging.
- Every node carries an absolute source span [start, end) in code-point offsets.
  Malformed input still yields a best-effort tree (never throws).
"""

from __future__ import annotations

import enum
from dataclasses import dataclass, field
from typing import ClassVar, TypeGuard, Union


@dataclass(frozen=True, slots=True)
class Span:
    """Absolute source offsets [start, end) (code points)."""

    start: int
    end: int


class ParseSeverity(str, enum.Enum):
    """Severity of a parse-time diagnostic."""

    ERROR = "error"
    WARNING = "warning"


@dataclass(slots=True)
class ParseDiagnostic:
    """A diagnostic produced while parsing (block mismatches, etc.)."""

    span: Span
    message: str
    severity: ParseSeverity
    # MS-VBAL section that justifies the rule, when applicable.
    spec_ref: str | None = None


class NodeKind(str, enum.Enum):
    """Discriminant tag for every statement/declaration node."""

    MODULE = "Module"
    ATTRIBUTE = "Attribute"
    OPTION = "Option"
    CONDITIONAL_DIRECTIVE = "ConditionalDirective"
    DECLARE = "Declare"
    EVENT = "Event"
    VARIABLE_GROUP = "VariableGroup"
    VARIABLE_DECL = "VariableDecl"
    TYPE = "Type"
    TYPE_FIELD = "TypeField"
    ENUM = "Enum"
    ENUM_MEMBER = "EnumMember"
    PROCEDURE = "Procedure"
    PARAMETER = "Parameter"
    ASSIGNMENT = "Assignment"
    CALL = "Call"
    IF_BLOCK = "IfBlock"
    FOR_BLOCK = "ForBlock"
    DO_BLOCK = "DoBlock"
    WHILE_BLOCK = "WhileBlock"
    WITH_BLOCK = "WithBlock"
    SELECT_BLOCK = "SelectBlock"
    STATEMENT = "Statement"


class ModuleKind(str, enum.Enum):
    """What kind of module this is (MS-VBAL 4.2)."""

    PROCEDURAL = "procedural"
    CLASS = "class"
    UNKNOWN = "unknown"


class ConditionalDirectiveKind(str, enum.Enum):
    """Conditional-compilation directive kind (MS-VBAL 3.4)."""

    CONST = "Const"
    IF = "If"
    ELSE_IF = "ElseIf"
    ELSE = "Else"
    END_IF = "EndIf"
    UNKNOWN = "Unknown"


class ProcKind(str, enum.Enum):
    """The kind of procedure (MS-VBAL 5.3)."""

    SUB = "Sub"
    FUNCTION = "Function"
    PROPERTY_GET = "PropertyGet"
    PROPERTY_LET = "PropertyLet"
    PROPERTY_SET = "PropertySet"


class ExprKind(str, enum.Enum):
    """Discriminant tag for expression nodes."""

    LITERAL = "LiteralExpr"
    IDENTIFIER = "IdentifierExpr"
    MEMBER_ACCESS = "MemberAccessExpr"
    INDEX = "IndexExpr"
    PAREN = "ParenExpr"
    UNARY = "UnaryExpr"
    BINARY = "BinaryExpr"
    NEW = "NewExpr"
    ADDRESS_OF = "AddressOfExpr"
    TYPE_OF_IS = "TypeOfIsExpr"


class LiteralKind(str, enum.Enum):
    """Literal value kind (MS-VBAL 5.6 literal-expression)."""

    INTEGER = "integer"
    FLOAT = "float"
    STRING = "string"
    DATE = "date"
    BOOLEAN = "boolean"
    NOTHING = "nothing"
    NULL = "null"
    EMPTY = "empty"


class MemberAccessKind(str, enum.Enum):
    """Member-access flavor: ordinary dot, or the receiver!name bang form."""

    DOT = "dot"
    BANG = "bang"


class IfBranchKind(str, enum.Enum):
    """Which arm of an If block a branch represents."""

    IF = "if"
    ELSE_IF = "elseif"
    ELSE = "else"


# Operator vocabularies (MS-VBAL 5.6). Kept as plain strings, matching the raw
# operator text the expression parser produces.
UNARY_OPERATORS: frozenset[str] = frozenset({"-", "+", "Not"})
BINARY_OPERATORS: frozenset[str] = frozenset(
    {
        "+", "-", "*", "/", "\\", "Mod", "^", "&",
        "=", "<>", "<", ">", "<=", ">=",
        "And", "Or", "Xor", "Eqv", "Imp", "Like", "Is",
    }
)


class Node:
    """Marker base for statement/declaration nodes (isinstance grouping)."""

    __slots__ = ()
    KIND: ClassVar[NodeKind]


class Expr:
    """Marker base for expression nodes (isinstance grouping)."""

    __slots__ = ()
    EXPR_KIND: ClassVar[ExprKind]


# ---------------------------------------------------------------------------
# Expressions (MS-VBAL 5.6)
# ---------------------------------------------------------------------------


@dataclass(slots=True)
class LiteralExpr(Expr):
    """A literal value: number, string, date, boolean, or Nothing/Null/Empty."""

    EXPR_KIND: ClassVar[ExprKind] = ExprKind.LITERAL
    span: Span
    literal_kind: LiteralKind
    raw: str


@dataclass(slots=True)
class IdentifierExpr(Expr):
    """A simple identifier reference (local, module, or project-visible name)."""

    EXPR_KIND: ClassVar[ExprKind] = ExprKind.IDENTIFIER
    span: Span
    name: str
    # Type-declaration character (%, &, #, !, @, $) if present.
    type_suffix: str | None = None


@dataclass(slots=True)
class MemberAccessExpr(Expr):
    """object.member access, or leading .member inside a With block.

    When `object_` is None the member resolves against the innermost With receiver.
    `access_kind` is BANG for the receiver!name form (sugar for the receiver's
    default member indexed by the string "name"); DOT/None for ordinary access.
    Bang names are not literal members of the receiver type, so member-existence
    rules must treat them differently from dot access.
    """

    EXPR_KIND: ClassVar[ExprKind] = ExprKind.MEMBER_ACCESS
    span: Span
    object_: ExprNode | None
    member: str
    member_span: Span
    access_kind: MemberAccessKind | None = None


@dataclass(slots=True)
class Argument:
    """One entry in a call / index argument list (MS-VBAL 5.6.9 / 5.4.2).

    Arguments may be positional, named (name:=value), or omitted (an empty slot
    that tells the callee to use the parameter default).
    """

    # The argument value expression, or None for an omitted slot.
    value: ExprNode | None
    # Span of the whole argument entry (name + value, or just the value).
    span: Span
    # Named-argument name (without the :=), or None for a positional arg.
    name: str | None = None
    # Span of the name token, when `name` is present.
    name_span: Span | None = None


@dataclass(slots=True)
class IndexExpr(Expr):
    """callee(arg, ...) covering function calls and array indexing."""

    EXPR_KIND: ClassVar[ExprKind] = ExprKind.INDEX
    span: Span
    callee: ExprNode
    args: list[Argument] = field(default_factory=list)


@dataclass(slots=True)
class ParenExpr(Expr):
    """(inner) parenthesized expression.

    Wrapping a ByRef argument in parens coerces it to ByVal at the call site.
    """

    EXPR_KIND: ClassVar[ExprKind] = ExprKind.PAREN
    span: Span
    inner: ExprNode


@dataclass(slots=True)
class UnaryExpr(Expr):
    """A prefix unary expression: -, +, or Not."""

    EXPR_KIND: ClassVar[ExprKind] = ExprKind.UNARY
    span: Span
    operator: str
    operand: ExprNode


@dataclass(slots=True)
class BinaryExpr(Expr):
    """A binary expression (arithmetic, comparison, logical, Like, Is)."""

    EXPR_KIND: ClassVar[ExprKind] = ExprKind.BINARY
    span: Span
    operator: str
    left: ExprNode
    right: ExprNode


@dataclass(slots=True)
class NewExpr(Expr):
    """New TypeName expression."""

    EXPR_KIND: ClassVar[ExprKind] = ExprKind.NEW
    span: Span
    type_name: str
    type_name_span: Span


@dataclass(slots=True)
class AddressOfExpr(Expr):
    """AddressOf procedureName expression."""

    EXPR_KIND: ClassVar[ExprKind] = ExprKind.ADDRESS_OF
    span: Span
    target: IdentifierExpr


@dataclass(slots=True)
class TypeOfIsExpr(Expr):
    """TypeOf expr Is TypeName expression."""

    EXPR_KIND: ClassVar[ExprKind] = ExprKind.TYPE_OF_IS
    span: Span
    operand: ExprNode
    type_name: str
    type_name_span: Span


ExprNode = Union[
    LiteralExpr,
    IdentifierExpr,
    MemberAccessExpr,
    IndexExpr,
    ParenExpr,
    UnaryExpr,
    BinaryExpr,
    NewExpr,
    AddressOfExpr,
    TypeOfIsExpr,
]


# ---------------------------------------------------------------------------
# Declarations and module-level nodes
# ---------------------------------------------------------------------------


@dataclass(slots=True)
class AttributeNode(Node):
    """Attribute name/value line, e.g. Attribute VB_Name = "Module1" (MS-VBAL 4.2)."""

    KIND: ClassVar[NodeKind] = NodeKind.ATTRIBUTE
    span: Span
    # Raw attribute target/name text after `Attribute`, e.g. VB_Name or
    # Value.VB_UserMemId.
    name: str
    name_span: Span
    # Raw value text (right of '='), unparsed.
    value_raw: str


@dataclass(slots=True)
class OptionNode(Node):
    """Option directive (MS-VBAL 5.2.1): Explicit / Base n / Compare m / Private Module."""

    KIND: ClassVar[NodeKind] = NodeKind.OPTION
    span: Span
    # Canonical text after "Option", e.g. "Explicit", "Base 1", "Compare Text".
    option_text: str


@dataclass(slots=True)
class ConditionalDirectiveNode(Node):
    """#Const, #If, #ElseIf, #Else, or #End If directive (MS-VBAL 3.4)."""

    KIND: ClassVar[NodeKind] = NodeKind.CONDITIONAL_DIRECTIVE
    span: Span
    directive_kind: ConditionalDirectiveKind
    # #Const compiler constant name.
    name: str | None = None
    name_span: Span | None = None
    # Raw #Const value expression, preserving source spacing.
    value_raw: str | None = None
    value_span: Span | None = None
    # Raw #If / #ElseIf condition expression, without the trailing Then.
    condition_raw: str | None = None
    condition_span: Span | None = None


@dataclass(slots=True)
class ParameterNode(Node):
    """A formal parameter (MS-VBAL 5.3.1)."""

    KIND: ClassVar[NodeKind] = NodeKind.PARAMETER
    span: Span
    name: str
    optional: bool = False
    by_val: bool = False
    by_ref: bool = False
    param_array: bool = False
    is_array: bool = False
    name_span: Span | None = None
    type_suffix: str | None = None
    type_suffix_span: Span | None = None
    has_as_clause: bool | None = None
    as_type: str | None = None
    default_raw: str | None = None


@dataclass(slots=True)
class DeclareNode(Node):
    """Declare statement for an external procedure (MS-VBAL 5.2.3.5)."""

    KIND: ClassVar[NodeKind] = NodeKind.DECLARE
    span: Span
    name: str
    is_function: bool
    ptr_safe: bool
    name_span: Span | None = None
    type_suffix: str | None = None
    type_suffix_span: Span | None = None
    has_as_clause: bool | None = None
    visibility: str | None = None
    lib_name: str | None = None
    alias_name: str | None = None
    params: list[ParameterNode] = field(default_factory=list)
    return_type: str | None = None


@dataclass(slots=True)
class EventNode(Node):
    """Event declaration in a class/document/UserForm module (MS-VBAL 5.2.5)."""

    KIND: ClassVar[NodeKind] = NodeKind.EVENT
    span: Span
    name: str
    name_span: Span | None = None
    visibility: str | None = None
    params: list[ParameterNode] = field(default_factory=list)


@dataclass(slots=True)
class VariableDeclNode(Node):
    """A single declared name within a VariableGroup."""

    KIND: ClassVar[NodeKind] = NodeKind.VARIABLE_DECL
    span: Span
    name: str
    is_array: bool = False
    is_new: bool = False
    name_span: Span | None = None
    type_suffix: str | None = None
    type_suffix_span: Span | None = None
    has_as_clause: bool | None = None
    as_type: str | None = None
    fixed_length: str | None = None
    default_raw: str | None = None
    array_bounds: str | None = None


@dataclass(slots=True)
class VariableGroupNode(Node):
    """A group of variable declarations sharing one modifier (MS-VBAL 5.2.3 / 5.2.4)."""

    KIND: ClassVar[NodeKind] = NodeKind.VARIABLE_GROUP
    span: Span
    modifier: str
    is_const: bool
    with_events: bool
    declarations: list[VariableDeclNode] = field(default_factory=list)


@dataclass(slots=True)
class TypeFieldNode(Node):
    """A field within a user-defined Type (MS-VBAL 5.2.3.3)."""

    KIND: ClassVar[NodeKind] = NodeKind.TYPE_FIELD
    span: Span
    name: str
    is_array: bool = False
    name_span: Span | None = None
    type_suffix: str | None = None
    type_suffix_span: Span | None = None
    has_as_clause: bool | None = None
    as_type: str | None = None
    fixed_length: str | None = None


@dataclass(slots=True)
class TypeNode(Node):
    """User-defined Type ... End Type (MS-VBAL 5.2.3.3)."""

    KIND: ClassVar[NodeKind] = NodeKind.TYPE
    span: Span
    name: str
    closed: bool
    name_span: Span | None = None
    visibility: str | None = None
    fields: list[TypeFieldNode] = field(default_factory=list)
    # Conditional-compilation directives (#If etc.) inside the Type body.
    directives: list[ConditionalDirectiveNode] | None = None


@dataclass(slots=True)
class EnumMemberNode(Node):
    """A member within an Enum (MS-VBAL 5.2.3.4)."""

    KIND: ClassVar[NodeKind] = NodeKind.ENUM_MEMBER
    span: Span
    name: str
    name_span: Span | None = None
    # Raw member value expression after `=`, when present.
    value_raw: str | None = None


@dataclass(slots=True)
class EnumNode(Node):
    """Enum ... End Enum (MS-VBAL 5.2.3.4)."""

    KIND: ClassVar[NodeKind] = NodeKind.ENUM
    span: Span
    name: str
    closed: bool
    name_span: Span | None = None
    visibility: str | None = None
    members: list[EnumMemberNode] = field(default_factory=list)
    # Conditional-compilation directives (#If etc.) inside the Enum body.
    directives: list[ConditionalDirectiveNode] | None = None


# ---------------------------------------------------------------------------
# Structured statement nodes (MS-VBAL 5.4)
# ---------------------------------------------------------------------------


@dataclass(slots=True)
class AssignmentNode(Node):
    """[Let] lhs = rhs or Set lhs = rhs (MS-VBAL 5.4.3).

    `is_set` distinguishes object-reference assignment from value assignment.
    """

    KIND: ClassVar[NodeKind] = NodeKind.ASSIGNMENT
    span: Span
    is_set: bool
    is_let: bool
    lhs: ExprNode
    rhs: ExprNode


@dataclass(slots=True)
class CallNode(Node):
    """[Call] callee [(args)] or implicit call `callee arg, ...` (MS-VBAL 5.4.2).

    When `has_call_keyword` is true the argument list must be parenthesised.
    """

    KIND: ClassVar[NodeKind] = NodeKind.CALL
    span: Span
    has_call_keyword: bool
    callee: ExprNode
    args: list[Argument] = field(default_factory=list)


@dataclass(slots=True)
class StatementNode(Node):
    """Generic catch-all statement (Exit, GoTo, label, Return, and anything not
    yet parsed into a structured node)."""

    KIND: ClassVar[NodeKind] = NodeKind.STATEMENT
    span: Span
    # Raw source text of the statement (without separators).
    raw: str


@dataclass(slots=True)
class IfBranchNode:
    """One arm of an If block (MS-VBAL 5.4.2.1).

    Branch modeling makes flow-sensitive analysis possible (each arm's statements
    and entry condition are explicit) without disturbing IfBlockNode.body, which
    stays a flat list for the generic body walkers. Not a NodeKind member.
    """

    branch_kind: IfBranchKind
    # Entry condition for if/elseif arms; None for else (or unparsable condition).
    condition: ExprNode | None
    # Statements in this arm only (excludes the arm's own header line).
    body: list[BodyNode]
    # Span of the arm header line (If ... Then / ElseIf ... Then / Else).
    header_span: Span
    span: Span
    # Raw condition text between the keyword and Then (absent for else).
    condition_raw: str | None = None
    condition_span: Span | None = None


@dataclass(slots=True)
class IfBlockNode(Node):
    """An If ... End If block (MS-VBAL 5.4.2.1)."""

    KIND: ClassVar[NodeKind] = NodeKind.IF_BLOCK
    span: Span
    closed: bool
    # Structured arms: always begins with the `if` arm, then any elseif/else.
    branches: list[IfBranchNode] = field(default_factory=list)
    # Flat list of every arm's statements (and the ElseIf/Else header lines as raw
    # Statements), in source order, for generic body walkers and block balance.
    body: list[BodyNode] = field(default_factory=list)


@dataclass(slots=True)
class ForBlockNode(Node):
    """A For ... Next or For Each ... Next block (MS-VBAL 5.4.2.4)."""

    KIND: ClassVar[NodeKind] = NodeKind.FOR_BLOCK
    span: Span
    each: bool
    closed: bool
    control_variable: str | None = None
    control_variable_span: Span | None = None
    source_expression: str | None = None
    source_expression_span: Span | None = None
    next_variable: str | None = None
    next_variable_span: Span | None = None
    body: list[BodyNode] = field(default_factory=list)


@dataclass(slots=True)
class DoBlockNode(Node):
    """A Do ... Loop block (MS-VBAL 5.4.2.3)."""

    KIND: ClassVar[NodeKind] = NodeKind.DO_BLOCK
    span: Span
    closed: bool
    body: list[BodyNode] = field(default_factory=list)


@dataclass(slots=True)
class WhileBlockNode(Node):
    """A While ... Wend block (MS-VBAL 5.4.2.2)."""

    KIND: ClassVar[NodeKind] = NodeKind.WHILE_BLOCK
    span: Span
    closed: bool
    body: list[BodyNode] = field(default_factory=list)


@dataclass(slots=True)
class WithBlockNode(Node):
    """A With ... End With block (MS-VBAL 5.4.2.6)."""

    KIND: ClassVar[NodeKind] = NodeKind.WITH_BLOCK
    span: Span
    closed: bool
    body: list[BodyNode] = field(default_factory=list)


@dataclass(slots=True)
class SelectBlockNode(Node):
    """A Select Case ... End Select block (MS-VBAL 5.4.2.5)."""

    KIND: ClassVar[NodeKind] = NodeKind.SELECT_BLOCK
    span: Span
    closed: bool
    body: list[BodyNode] = field(default_factory=list)


@dataclass(slots=True)
class ProcedureNode(Node):
    """A Sub / Function / Property procedure (MS-VBAL 5.3)."""

    KIND: ClassVar[NodeKind] = NodeKind.PROCEDURE
    span: Span
    proc_kind: ProcKind
    name: str
    closed: bool
    name_span: Span | None = None
    type_suffix: str | None = None
    type_suffix_span: Span | None = None
    has_as_clause: bool | None = None
    modifiers: list[str] = field(default_factory=list)
    params: list[ParameterNode] = field(default_factory=list)
    return_type: str | None = None
    # Exported member metadata lines such as Attribute Value.VB_UserMemId = 0.
    attributes: list[AttributeNode] | None = None
    body: list[BodyNode] = field(default_factory=list)


@dataclass(slots=True)
class ModuleNode(Node):
    """A whole VBA module (MS-VBAL 4.2)."""

    KIND: ClassVar[NodeKind] = NodeKind.MODULE
    span: Span
    module_kind: ModuleKind
    members: list[ModuleMember] = field(default_factory=list)
    diagnostics: list[ParseDiagnostic] = field(default_factory=list)


# Body nodes (inside a procedure body).
BodyNode = Union[
    AssignmentNode,
    CallNode,
    StatementNode,
    ConditionalDirectiveNode,
    VariableGroupNode,
    IfBlockNode,
    ForBlockNode,
    DoBlockNode,
    WhileBlockNode,
    WithBlockNode,
    SelectBlockNode,
]

# Leaf executable statement: a structured assignment or call, or the raw
# Statement catch-all. All three carry a single statement span and no nested body.
LeafStatementNode = Union[AssignmentNode, CallNode, StatementNode]

# Module-level members.
ModuleMember = Union[
    AttributeNode,
    OptionNode,
    ConditionalDirectiveNode,
    DeclareNode,
    EventNode,
    VariableGroupNode,
    TypeNode,
    EnumNode,
    ProcedureNode,
    StatementNode,
]


def is_leaf_statement(node: object) -> TypeGuard[LeafStatementNode]:
    """True for the leaf statement nodes (Assignment / Call / raw Statement)."""
    return isinstance(node, (AssignmentNode, CallNode, StatementNode))
