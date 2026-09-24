# API reference

The package root re-exports the headline entry points; the rest of the public
surface lives in the submodules listed below. This page is an import map, not a
full autodoc dump. For task-oriented examples see [usage.md](usage.md).

## Headline API (`pyvbaanalysis`)

These are re-exported at the package root for convenience:

| Symbol | Role |
| --- | --- |
| `analyze_module(source, opts=None)` | Analyze one module's source; returns `list[VbaDiagnostic]`. Never raises. |
| `analyze_project(modules, *, only=None, severity_overrides=None, conditional_compilation=None, whole_project=True, inline_suppression=True, host=None, referenced_hosts=None)` | Analyze a set of `ModuleInput`s with cross-module context; returns `dict[str, list[VbaDiagnostic]]`. `whole_project=False` for a partial set. `host` names the Office host (absent means Excel). `referenced_hosts` names the other Office libraries the project references, in declaration order; `None` means the reference list is unknown, `[]` that it is known to name nothing else. |
| `analyze_loose_file(path, *, severity_overrides=None, conditional_compilation=None, whole_project=False, inline_suppression=True)` | Analyze one `.bas`/`.cls`/`.frm` file (partial by default; the whole-project checks are skipped). |
| `analyze_loose_files(paths, *, only=None, severity_overrides=None, conditional_compilation=None, whole_project=True, inline_suppression=True)` | Analyze several loose files as one project. |
| `analyze_workbook(path, *, only=None, severity_overrides=None, conditional_compilation=None, inline_suppression=True)` | The Excel-only form of `analyze_office_file`: refuses any other extension, and otherwise returns exactly what it does. |
| `analyze_office_file(path, *, only=None, severity_overrides=None, conditional_compilation=None, inline_suppression=True)` | Analyze the VBA in any readable Office container (Excel, Word, PowerPoint, Access, via pyOpenVBA); the extension selects the host model, and the project's reference list is read with its modules. |
| `host_token_for_file_name(name)` | The host a container implies (`excel`, `word`, `powerpoint`, `access`), or `None`. |
| `host_object_model_for_token(host)` | The model a host token selects: `None` for Excel (the default), the host's own model (`word`, `powerpoint`, `access`, `vb6`), or the empty model when a named host has none. |
| `analyze_module_options_for(index, name, kind, *, severity_overrides=None, conditional_compilation=None, whole_project=True, inline_suppression=True, host=None, referenced_hosts=None)` | Build per-module `AnalyzeModuleOptions` from a populated `ProjectIndex`. |
| `build_project_index(modules)` | A `ProjectIndex` with every module registered. |
| `AnalyzeModuleOptions` | Inputs for `analyze_module` (name, kind, project context, overrides). |
| `VbaDiagnostic` | A single diagnostic (`code`, `message`, `severity`, `span`, `spec_reference`). |
| `ModuleInput` | One module fed into a project (`module_name`, `module_kind`, `source`), plus what a caller that read the module's designer or attribute header knows: `implicit_members` (a UserForm's controls, as `ImplicitMember`s), `predeclared_id` and `designer_class`. Left `None`, the index reads what the source itself carries. |
| `ModuleSymbolKind` | `STANDARD`, `CLASS`, `DOCUMENT`, `USERFORM`. |
| `ProjectIndex` | The cross-module symbol/type index. |
| `__version__` | The package version string. |

## Engine and result model (`pyvbaanalysis.diagnostics`)

* Entry point: `analyze_module`, `AnalyzeModuleOptions`, `RulePassContext`, `PushFn`,
  `is_object_module_kind`.
* Result model: `VbaDiagnostic`, `VbaDiagnosticData`, `VbaEdit`, and the enums
  `DiagnosticSeverity`, `DiagnosticCategory`, `DiagnosticEvidenceKind`,
  `DiagnosticSuppressionScope`.
* Code-action payloads, carried on `VbaDiagnosticData`:
  `VbaMissingRequiredArgumentPlaceholderData`, `VbaCreateProcedureStubData`,
  `VbaAddLibraryReferenceData` (the library a `missing-library-reference` finding
  needs), `VbaRemoveDeclarationData` and `VbaRemoveUnreachableCodeData` (the edit
  that deletes an unused declaration or an unreachable run), and `VbaDocCommentFix`
  (each way to bring a doc comment in line with its declaration).
* Registry and metadata: `DIAGNOSTIC_RULE_REGISTRY`, `DiagnosticRuleEntry`,
  `STRUCTURAL_DIAGNOSTIC_RULES`, `DIAGNOSTIC_RULES`, `DiagnosticRuleMetadata`,
  `diagnostic_metadata_for_code`, `rule_metadata_by_code`, `load_rule_metadata`,
  `normalize_diagnostic_severity_override`.

## Readers (`pyvbaanalysis.reader`)

* High level: `analyze_loose_file`, `analyze_loose_files`, `analyze_workbook`,
  `analyze_office_file`, `load_loose_module`, `read_workbook_modules`,
  `read_office_modules`, `read_office_project`.
* `read_office_project(path)` returns an `OfficeProject`: the container's
  `modules`, the `host` its extension implies, and `referenced_hosts`, the other
  Office libraries its project references (`None` when the reference list could
  not be read).
* A `LoadedModule` carries a module's `name`, `kind`, and `source` (the code body
  the analyzer reads), plus `designer_block`, the export header the reader stripped
  ahead of `source` (empty when there was none).
* Building blocks: `LoadedModule`, `strip_export_header`, `classify_module_kind`,
  `module_name_from_text`, `loaded_module_from_text`, `WorkbookReadError`,
  `LooseFileReadError`, and the extension sets `LOOSE_EXTENSIONS`,
  `EXCEL_EXTENSIONS`, `WORD_EXTENSIONS`, `POWERPOINT_EXTENSIONS`,
  `ACCESS_EXTENSIONS`, `OFFICE_EXTENSIONS`.

## Project model (`pyvbaanalysis.symbols`)

* Index: `ProjectIndex`, `ProjectIndexOptions`, `ModuleInput`, `ImplicitMember`,
  `ReferenceScope`, `ShadowedSpan`.
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
| `pyvbaanalysis.host` | The Excel, Word, PowerPoint, Access and VB6 object models and the host registry (`get_excel_object_model`, `get_vb6_object_model`, `host_object_model_for_token`, `resolve_host_alias`, ...). `host_object_model_for_tokens` merges the models for a project and its references, the first token winning a shared name; `host_token_for_libid`, `host_tokens_for_project` and `referenced_host_tokens` map a project's reference libids to host tokens. `pyvbaanalysis.host.msforms` carries the Microsoft Forms members of a UserForm and its controls (`msforms_control_members`, `resolve_msforms_type_name`). |
| `pyvbaanalysis.references` | `classify_reference_kinds`: whether each mention of a name reads it, writes it, or modifies it in place. |
| `pyvbaanalysis.docs` | The `'''` doc-comment grammar: `leading_doc_lines`, `scan_doc_tags`, `attached_comments_start`. |
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
