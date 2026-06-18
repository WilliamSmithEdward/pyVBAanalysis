"""VBA parser: logical statements, the AST node model, and parse entry points.

Concrete node classes live in pyvbaanalysis.parser.nodes; the entry points
parse_module and parse_expression are re-exported here.
"""

from .nodes import ModuleNode
from .parse_expression import ExprParseResult, parse_expression, parse_parenless_arguments
from .parse_module import parse_module
from .parser_state import LogicalStatement, StatementCursor, code_tokens, split_logical_statements

__all__ = [
    "parse_module",
    "parse_expression",
    "parse_parenless_arguments",
    "ExprParseResult",
    "ModuleNode",
    "LogicalStatement",
    "StatementCursor",
    "code_tokens",
    "split_logical_statements",
]
