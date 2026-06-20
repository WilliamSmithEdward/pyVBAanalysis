# API reference

The package root re-exports the headline entry points; the rest of the public
surface lives in the submodules listed below. This page is an import map, not a
full autodoc dump. For task-oriented examples see [usage.md](usage.md).

## Headline API (`pyvbaanalysis`)

These are re-exported at the package root for convenience:

| Symbol | Role |
| --- | --- |
| `analyze_module(source, opts=None)` | Analyze one module's source; returns `list[VbaDiagnostic]`. Never raises. |
| `analyze_project(modules, *, only=None, severity_overrides=None, conditional_compilation=None)` | Analyze a set of `ModuleInput`s with cross-module context; returns `dict[str, list[VbaDiagnostic]]`. |
| `analyze_loose_file(path, *, severity_overrides=None, conditional_compilation=None)` | Analyze one `.bas`/`.cls`/`.frm` file. |
| `analyze_loose_files(paths, *, only=None, severity_overrides=None, conditional_compilation=None)` | Analyze several loose files as one project. |
| `analyze_workbook(path, *, only=None, severity_overrides=None, conditional_compilation=None)` | Analyze the VBA in an Excel workbook (via pyOpenVBA). |
| `analyze_module_options_for(index, name, kind, *, severity_overrides=None, conditional_compilation=None)` | Build per-module `AnalyzeModuleOptions` from a populated `ProjectIndex`. |
| `build_project_index(modules)` | A `ProjectIndex` with every module registered. |
| `AnalyzeModuleOptions` | Inputs for `analyze_module` (name, kind, project context, overrides). |
| `VbaDiagnostic` | A single diagnostic (`code`, `message`, `severity`, `span`, `spec_reference`). |
| `ModuleInput` | One module fed into a project (`module_name`, `module_kind`, `source`). |
| `ModuleSymbolKind` | `STANDARD`, `CLASS`, `DOCUMENT`, `USERFORM`. |
| `ProjectIndex` | The cross-module symbol/type index. |
| `__version__` | The package version string. |

## Engine and result model (`pyvbaanalysis.diagnostics`)

* Entry point: `analyze_module`, `AnalyzeModuleOptions`, `RulePassContext`, `PushFn`,
  `is_object_module_kind`.
* Result model: `VbaDiagnostic`, `VbaDiagnosticData`, `VbaEdit`, and the enums
  `DiagnosticSeverity`, `DiagnosticCategory`, `DiagnosticEvidenceKind`,
  `DiagnosticSuppressionScope`.
* Registry and metadata: `DIAGNOSTIC_RULE_REGISTRY`, `DiagnosticRuleEntry`,
  `STRUCTURAL_DIAGNOSTIC_RULES`, `DIAGNOSTIC_RULES`, `DiagnosticRuleMetadata`,
  `diagnostic_metadata_for_code`, `rule_metadata_by_code`, `load_rule_metadata`,
  `normalize_diagnostic_severity_override`.

## Readers (`pyvbaanalysis.reader`)

* High level: `analyze_loose_file`, `analyze_loose_files`, `analyze_workbook`,
  `load_loose_module`, `read_workbook_modules`.
* Building blocks: `LoadedModule`, `strip_export_header`, `classify_module_kind`,
  `module_name_from_text`, `loaded_module_from_text`, `WorkbookReadError`,
  `LOOSE_EXTENSIONS`, `EXCEL_EXTENSIONS`.

## Project model (`pyvbaanalysis.symbols`)

* Index: `ProjectIndex`, `ProjectIndexOptions`, `ModuleInput`, `ReferenceScope`,
  `ShadowedSpan`.
* Module symbols: `build_module_symbols`, `BuildModuleSymbolsOptions`,
  `ModuleSymbols`, `ModuleSymbolKind`, `SymbolVisibility`, `VbaSymbol`,
  `VbaSymbolKind`, `VbaSymbolAttribute`, `VbaProcedureSignature`,
  `VbaProcedureParam`, and the predicates `is_procedure_kind`,
  `is_bare_callable_kind`, `procedure_signature_from_symbol`,
  `qualified_procedure_key`.
* Name resolution: `resolve_bare_identifier_binding`, `source_identifier_names`,
  `BareIdentifierContext`, `BareIdentifierResolution`.

## Lower-level building blocks

These are the analysis layers the engine is built on. Most consumers do not need
them directly, but they are public and stable.

| Package | What it provides |
| --- | --- |
| `pyvbaanalysis.lexer` | `tokenize`, `tokenize_cached`, `VbaToken`, `TokenKind`, and token helpers. |
| `pyvbaanalysis.parser` | `parse_module`, `parse_expression`, `ModuleNode`, and the AST nodes. |
| `pyvbaanalysis.conditional` | Conditional-compilation indexing and activity (`index_conditional_compilation`, `evaluate_conditional_expression`, ...). |
| `pyvbaanalysis.types` | Type-name helpers: `normalize_type`, `is_known_scalar_type`, `is_numeric_type`, `numeric_literal_bounds`. |
| `pyvbaanalysis.completion` | Member-completion surface and project-type resolution (`resolve_member_surface_at`, `resolve_type_name`, ...). |
| `pyvbaanalysis.host` | The Excel host object model (`get_excel_object_model`, `resolve_host_alias`, ...). |
| `pyvbaanalysis.runtime` | VBA runtime functions, constants, and objects (`resolve_runtime_function`, ...). |
| `pyvbaanalysis.call` | Call-statement shape helpers. |
| `pyvbaanalysis.flow` | Procedure labels and unstructured-flow detection. |
| `pyvbaanalysis.constants` | Integer constant-expression evaluation. |

Many helper free functions inside these packages are intentionally internal even
when importable. The symbols listed above are the supported surface.

## A note on the package root

The package root exports only the headline API plus `__version__`. Everything
else is imported from its submodule, e.g.
`from pyvbaanalysis.diagnostics import DiagnosticSeverity` or
`from pyvbaanalysis.symbols import ModuleSymbols`.
