"""pyVBAanalysis: pure-Python static analysis for VBA.

Ported from the XLIDE analyzer
(https://github.com/WilliamSmithEdward/xlide_vscode). See agent.md for the build
plan, parity inventory, and port order. The public API (analyze_module,
parse_module, tokenize, ProjectIndex) is added incrementally during the port,
starting at milestone M0.
"""

__version__ = "0.0.0"
