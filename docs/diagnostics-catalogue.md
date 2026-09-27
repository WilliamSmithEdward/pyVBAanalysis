# Diagnostic catalogue

The diagnostic codes pyVBAanalysis can emit, generated from the rule metadata (`tools/generate_diagnostics_catalogue.py`). This table lists the 165 rule-metadata codes across 6 categories. A further 3 structural block-balance codes (`mismatched-end-keyword`, `missing-block-closer`, `unmatched-block-closer`) are emitted by the parser pass and are not in the metadata table, for a full set of 168 codes.

Each code is reported only when it is provably correct; anything unknown or ambiguous stays quiet (the no-false-positive discipline). The **kind** column says what a code means: a *compile error* is rejected by the VBE compiler, a *runtime error* is a deterministic Run-time error, a *runtime risk* is a likely fault, and *style* is advisory.

Override a code's severity with `AnalyzeModuleOptions.severity_overrides` (or the `severity_overrides` argument of `analyze_project` / the reader functions), keyed by code. Use `"off"`, `"information"`, `"warning"`, or `"error"`; the allowed values per code are constrained by policy (some codes can be downgraded but not disabled). See [docs/usage.md](usage.md).

## Declaration (46)

| Code | Title | Default | Kind | Spec reference |
| --- | --- | --- | --- | --- |
| `array-parameter-form` | Array parameter cannot be ByVal or Optional | error | compile error | MS-VBAL 5.3.1.5; VBE "Array argument must be ByRef" / "Optional argument must be Variant or intrinsic type with a default value" (issue #124, Excel 16.0) |
| `byval-udt-parameter` | User-defined type parameter cannot be ByVal | error | compile error | MS-VBAL 5.3.1 (parameter passing) / VBE compiler |
| `const-overflow` | Const value overflows | error | compile error | MS-VBAL 5.6.9.3; VBE compile error: Overflow (issue #116, Excel 16.0) |
| `const-value-not-constant` | Const value must be a constant expression | error | compile error | MS-VBAL 5.2.4 (Const declaration value is a constant-expression) |
| `declare-missing-ptrsafe` | Declare statement missing PtrSafe for 64-bit Office | error | compile error | VBA 7 Declare statement PtrSafe requirement for 64-bit Office |
| `dim-initializer` | Declaration cannot include an initializer (VB.NET syntax) | error | compile error | MS-VBAL 5.2.3.1 |
| `duplicate-const-directive` | Duplicate #Const | error | compile error | MS-VBAL 3.4.1 #Const; VBE "Duplicate definition" (issue #130, Excel 16.0) |
| `duplicate-declaration` | Duplicate declaration in the current scope | error | compile error | MS-VBAL 5.2 / 5.3 |
| `duplicate-deftype` | Duplicate Deftype statement | error | compile error | MS-VBAL 5.2.3.1.3; VBE "Duplicate Deftype statement" (issue #124, Excel 16.0) |
| `duplicate-enum-member` | Duplicate Enum member | error | compile error | MS-VBAL 5.2.3.4 |
| `duplicate-module-variable` | Duplicate module-level declaration | error | compile error | MS-VBAL 5.2.3 |
| `duplicate-option` | Duplicate Option statement | error | compile error | MS-VBAL 5.2.1 (module options) |
| `duplicate-procedure` | Ambiguous (duplicate) procedure name in module | error | compile error | MS-VBAL 5.3 |
| `duplicate-type-field` | Duplicate Type field | error | compile error | MS-VBAL 5.2.3.3 |
| `empty-type` | User-defined Type has no members | error | compile error | MS-VBAL 5.2.3.3 (UDT declaration) |
| `enum-member-not-constant` | Enum member value must be a constant expression | error | compile error | MS-VBAL 5.2.3.4 (Enum member value is a constant-expression) |
| `fixed-length-string-size` | Invalid fixed-length String size | error | compile error | MS-VBAL fixed-length String bounds / VBE compiler: Invalid length for fixed-length string |
| `identifier-too-long` | Identifier exceeds 255 characters | error | compile error | MS-VBAL 3.3.5.1 (identifier length) / VBE compiler |
| `invalid-as-type-name` | Invalid type name | error | compile error | MS-VBAL 3.3.5.2 / 5.2.3.1 |
| `invalid-declaration-name` | Reserved keyword used as a declaration name | error | compile error | MS-VBAL 3.3.5.2 |
| `invalid-identifier-character` | Invalid character in identifier | error | compile error | MS-VBAL 3.3.5 (identifier) / VBE compiler |
| `invalid-identifier-start` | Invalid identifier start | error | compile error | MS-VBAL 3.3.5 |
| `invalid-new-type-name` | Type cannot be created with New | error | compile error | MS-VBAL 5.2.3.1 / 5.6.9 |
| `invalid-option-statement` | Malformed Option statement | error | compile error | MS-VBAL 5.2.1 (module options) / VBE compiler |
| `invalid-proc-header` | Invalid procedure declaration | error | compile error | MS-VBAL 5.3.1 |
| `module-declaration-after-procedure` | Module-level declaration after procedure | error | compile error | MS-VBAL 5.2 / 5.3 |
| `module-declaration-in-procedure` | Module-level declaration inside procedure | error | compile error | MS-VBAL 5.2 / 5.3 |
| `object-module-public-member` | Invalid public member in object module | error | compile error | VBE compiler: public object-module member restrictions |
| `option-after-declaration` | Option statement after a procedure | error | compile error | MS-VBAL 5.2.1 |
| `optional-udt-parameter` | Optional parameter cannot be a user-defined type | error | compile error | MS-VBAL 5.3.1 (Optional parameter) / VBE compiler |
| `paramarray-non-variant` | ParamArray elements must be Variant | error | compile error | MS-VBAL 5.3.1.6 |
| `paramarray-not-last` | ParamArray must be the final parameter | error | compile error | MS-VBAL 5.3.1.6 |
| `paramarray-with-optional` | ParamArray cannot be combined with Optional parameters | error | compile error | MS-VBAL 5.3.1.5 / 5.3.1.6 |
| `parameter-array-as-type-syntax` | Array parameter parentheses must follow the parameter name | error | compile error | MS-VBAL 5.3.1.5 |
| `parameter-default-not-constant` | Optional default must be a constant expression | error | compile error | MS-VBAL 5.3.1.5 (Optional default-value is a constant-expression) |
| `parameter-default-type-mismatch` | Parameter default type mismatch | error | compile error | MS-VBAL 5.3.1 / VBE compiler: Type mismatch |
| `property-accessor-signature-mismatch` | Property accessors have incompatible signatures | error | compile error | MS-VBAL 5.3.1.4 |
| `property-let-object-value` | Property Let value parameter must not be object reference | error | compile error | MS-VBAL 5.3.1.4 |
| `property-set-scalar-value` | Property Set value parameter must be object reference | error | compile error | MS-VBAL 5.3.1.4 |
| `property-setter-missing-value` | Property setter is missing value parameter | error | compile error | MS-VBAL 5.3.1.4 |
| `property-setter-return-type` | Property setter cannot declare a return type | error | compile error | MS-VBAL 5.3.1.4 |
| `required-param-after-optional` | A required parameter cannot follow an Optional parameter | error | compile error | MS-VBAL 5.3.1.5 |
| `too-many-array-dimensions` | Array has more than 60 dimensions | error | compile error | MS-VBAL 5.2.3.1 (array declaration) / VBE compiler |
| `too-many-parameters` | Procedure has more than 60 parameters | error | compile error | MS-VBAL 5.3.1 (procedure parameters) / VBE compiler |
| `type-declaration-character-as-clause` | Invalid type-declaration character with As | error | compile error | MS-VBAL 5.2.3.1 / 5.3.1 type-declaration characters |
| `unexpected-declaration-token` | Unexpected token after declaration type | error | compile error | MS-VBAL 5.2.3.1 / VBE compiler: Syntax error |

## Module-kind (5)

| Code | Title | Default | Kind | Spec reference |
| --- | --- | --- | --- | --- |
| `event-declaration-module-kind` | Event declaration is not valid in this module | error | compile error | MS-VBAL 5.2.5: Event declarations belong to object modules |
| `event-handler-module-scope` | Event handler is not wired in this module | information | style-policy | Office document-module event binding |
| `friend-declaration` | Invalid Friend declaration | error | compile error | MS-VBAL Friend procedure visibility: object-module procedures only |
| `implements-statement-placement` | Invalid Implements statement | error | compile error | MS-VBAL Implements statement: module-level object-module declaration |
| `withevents-declaration` | Invalid WithEvents declaration | error | compile error | MS-VBAL 5.2.3: WithEvents object variable declarations |

## Project-symbol (2)

| Code | Title | Default | Kind | Spec reference |
| --- | --- | --- | --- | --- |
| `undeclared-variable` | Variable not defined | error | compile error | MS-VBAL 5.2.4.1.1 |
| `unknown-call` | Sub or Function not defined | error | compile error | MS-VBAL 5.4.2.1 |

## Semantic (75)

| Code | Title | Default | Kind | Spec reference |
| --- | --- | --- | --- | --- |
| `ambiguous-enum-member` | Ambiguous Enum member reference | error | compile error | VBE compiler: Ambiguous name detected |
| `ambiguous-project-procedure` | Ambiguous unqualified call to a name two modules export | error | compile error | VBE compiler: Ambiguous name detected |
| `argument-count` | Wrong number of arguments | error | compile error | MS-VBAL 5.4.2.1 |
| `argument-object-type-mismatch` | Object argument type mismatch | error | compile error | MS-VBAL 5.3.1 |
| `argument-shape-mismatch` | Argument shape (array/Type vs scalar) mismatch | error | compile error | MS-VBAL 5.3.1 (argument passing) / VBE compiler: ByRef argument type mismatch; array or user-defined type expected |
| `argument-type-mismatch` | Argument type mismatch | error | runtime error | MS-VBAL 5.3.1 / runtime type coercion and numeric overflow |
| `arithmetic-overflow` | Arithmetic overflow | error | runtime error | MS-VBAL 5.6.9.3 operator result types; VBE runtime error 6: Overflow (issue #116, Excel 16.0) |
| `array-assignment-to-scalar` | Array cannot be assigned to scalar | error | compile error | MS-VBAL 5.4.3 / VBE compiler: Type mismatch |
| `array-bound-requires-array` | Array bound function requires array | error | compile error | MS-VBAL LBound/UBound / VBE compiler: Expected array |
| `array-declaration-impossible-bounds` | Array declaration lower bound is greater than upper bound | error | compile error | MS-VBAL 5.2.3.1 (array declaration bounds) |
| `array-subscript-out-of-bounds` | Array subscript out of range | error | runtime error | VBE runtime error 9: Subscript out of range |
| `assignment-object-type-mismatch` | Object assignment type mismatch | error | compile error | MS-VBAL 5.4.3 / Set statement |
| `assignment-type-mismatch` | Assignment type mismatch | error | runtime error | MS-VBAL 5.4.3 / runtime type coercion and numeric overflow |
| `byref-argument-type-mismatch` | ByRef argument type mismatch | error | compile error | MS-VBAL 5.3.1 / VBE compiler: ByRef argument type mismatch |
| `case-outside-select` | Case statement outside Select Case | error | compile error | MS-VBAL 5.4.2.4 |
| `collection-index-out-of-range` | Collection index out of range | error | runtime error | VBE runtime errors 5 and 9 on VBA.Collection (issue #121, Excel 16.0) |
| `collection-key-in-use` | Collection key already in use | error | runtime error | VBE runtime error 457 on VBA.Collection (issue #121, Excel 16.0) |
| `collection-key-not-found` | Collection key not found | error | runtime error | VBE runtime error 5 on VBA.Collection (issue #121, Excel 16.0) |
| `collection-operand` | Collection used as an operand | error | compile error | VBE "Argument not optional" on a Collection whose default member Item takes an index (issue #125, Excel 16.0) |
| `const-assignment` | Assignment to a constant | error | compile error | MS-VBAL 5.4.3.1 |
| `division-by-zero` | Division by zero | error | runtime error | MS-VBAL 5.6 / runtime division by zero |
| `duplicate-case-else` | Duplicate Case Else in Select Case | error | compile error | MS-VBAL 5.4.2.10 (Select Case) |
| `duplicate-label` | Duplicate procedure label | error | compile error | MS-VBAL 5.4.1 |
| `else-without-if` | 'Else'/'ElseIf' outside an If block | error | compile error | MS-VBAL 5.4.2.1 (If block) / VBE compiler |
| `erase-requires-array` | Erase target must be array or Variant | error | compile error | MS-VBAL Erase statement / VBE compiler: Expected array |
| `exit-outside-block` | Loop exit statement outside matching loop | error | compile error | MS-VBAL 5.4.1.3 |
| `exit-wrong-proc` | Exit statement does not match the enclosing procedure | error | compile error | MS-VBAL 5.4.1.3 |
| `file-already-open` | File number opened twice | error | runtime error | VBE runtime error 55: File already open (issue #123, Excel 16.0) |
| `file-mode-mismatch` | File statement the open mode forbids | error | runtime error | VBE runtime error 54: Bad file mode (issue #123, Excel 16.0) |
| `file-number-zero` | File number 0 | error | runtime error | VBE runtime error 52: Bad file name or number (issue #123, Excel 16.0) |
| `file-record-zero` | Record number 0 | error | runtime error | VBE runtime error 63: Bad record number (issue #123, Excel 16.0) |
| `file-used-after-close` | File statement on a closed file number | error | runtime error | VBE runtime error 52: Bad file name or number (issue #123, Excel 16.0) |
| `fixed-array-redim` | Fixed-size array cannot be ReDimmed | error | compile error | MS-VBAL ReDim statement |
| `for-counter-overflow` | For counter overflows after its last pass | error | runtime error | MS-VBAL 5.4.2.3; VBE runtime error 6: Overflow (issue #116, Excel 16.0) |
| `for-each-control-variable-type` | For Each control variable must be Variant or Object | error | compile error | MS-VBAL 5.4.2.5 |
| `for-each-source-type` | For Each source must be collection or array | error | compile error | MS-VBAL 5.4.2.5 / VBE compiler: For Each may only iterate over a collection object or an array |
| `handler-fall-through` | Execution falls into an error handler that re-raises | error | runtime error | VBE runtime error 5 from Err.Raise Err.Number with Err.Number 0 (issue #117, Excel 16.0) |
| `host-argument-out-of-range` | Host argument out of range | error | runtime error | Office object model: 1-based collections, cell rows and columns from 1 (issue #122, Excel/Word/PowerPoint 16.0) |
| `implements-member-missing` | Interface member not implemented | error | compile error | MS-VBAL 5.2.4.1; VBE "Object module needs to implement ... for interface ..." (issue #125, Excel 16.0) |
| `implements-member-signature` | Interface member implemented with another signature | error | compile error | MS-VBAL 5.2.4.1; VBE "Procedure declaration does not match description of event or procedure having the same name" (issue #125, Excel 16.0) |
| `invalid-assignment-target` | Cannot assign to a literal value | error | compile error | MS-VBAL 5.4.3 (assignment) / VBE compiler |
| `is-operator-non-object` | 'Is' operator requires object operands | error | compile error | MS-VBAL 5.6 (Is operator) |
| `late-bound-friend-member` | Friend member reached through a late-bound receiver | error | runtime error | VBE runtime error 438: Object does not support this property or method |
| `me-outside-object-module` | 'Me' is only valid in an object module | error | compile error | MS-VBAL 5.6.2.2 (Me) / VBE compiler |
| `member-access-outside-with` | Leading member access outside With block | error | compile error | MS-VBAL 5.4.2.6 |
| `member-not-found` | Object member not found | error | compile error | VBE compiler: Method or data member not found |
| `mid-statement-literal-target` | Mid statement target must be a writable String variable | error | compile error | MS-VBAL 5.4.3.4 (Mid/MidB statement) |
| `missing-library-reference` | Type library not referenced | error | compile error | VBE compiler: User-defined type not defined |
| `missing-return-assignment` | Function has no return assignment | warning | runtime risk | VBA Function return variable semantics |
| `multi-cell-range-as-scalar` | Multi-cell range read as a scalar | error | runtime error | VBE runtime error 13: Type mismatch on a Range value that is an array (issue #122, Excel 16.0) |
| `next-variable-mismatch` | Next variable does not match active For loop | error | compile error | MS-VBAL 5.4.2.5 |
| `non-callable-call` | Identifier is not callable | error | compile error | MS-VBAL 5.4.2.1 |
| `non-scalar-binary-operand` | Operator requires a scalar operand | error | compile error | MS-VBAL 5.6 (binary operators) / VBE compiler: array operand Type mismatch |
| `object-variable-not-set` | Object variable not set | error | runtime error | VBE runtime error 91: Object variable or With block variable not set |
| `raiseevent-undeclared-event` | RaiseEvent target is not declared | error | compile error | MS-VBAL RaiseEvent statement: event name resolution |
| `readonly-member-assignment` | Assignment to a read-only member | error | compile error | VBE compiler: Can't assign to read-only property |
| `recursive-property-accessor` | Property procedure calls itself | error | runtime error | VBE runtime error 28: Out of stack space (issue #117, Excel 16.0) |
| `redim-impossible-bounds` | ReDim lower bound is greater than upper bound | error | runtime error | MS-VBAL ReDim statement / VBE compiler runtime error 9 |
| `redim-preserve-dimension-change` | ReDim Preserve can only resize the last dimension | error | runtime error | MS-VBAL ReDim Preserve statement |
| `resume-without-error` | Resume with no error handler installed | error | runtime error | VBE runtime error 20: Resume without error (issue #117, Excel 16.0) |
| `return-without-gosub` | Execution falls into a GoSub target | error | runtime error | VBE runtime error 3: Return without GoSub (issue #117, Excel 16.0) |
| `runtime-argument-value` | Invalid runtime argument value | error | runtime error | MS-VBAL 5.6 / VBA runtime argument bounds and VBE compiler runtime error 5 |
| `runtime-conversion-value` | Invalid runtime conversion value | error | runtime error | MS-VBAL 5.6 / VBA runtime conversion and VBE compiler runtime error 13 |
| `runtime-member-not-found` | Member not found at run time | error | runtime error | VBE runtime error 438: Object doesn't support this property or method (issue #121, Excel 16.0) |
| `scalar-member-access` | Member access on scalar variable | error | compile error | VBE compiler: Invalid qualifier / Syntax error |
| `scalar-redim` | Scalar variable cannot be ReDimmed | error | compile error | MS-VBAL ReDim statement / VBE compiler: Expected array |
| `set-required` | Object assignment requires Set | error | runtime error | VBE runtime error 91: Object variable or With block variable not set (MS-VBAL 5.4.3 / Set statement) |
| `set-requires-object` | Set assignment requires an object variable | error | compile error | MS-VBAL 5.4.3 |
| `sheet-name-invalid` | Sheet name Excel refuses | error | runtime error | Excel runtime error 1004: invalid name for a sheet or chart (issue #122, Excel 16.0) |
| `string-arithmetic-coercion` | Nonnumeric string in numeric expression | error | runtime error | MS-VBAL 5.6 / runtime type coercion |
| `sub-used-as-value` | Sub used as a value | error | compile error | VBE "Expected Function or variable" (issue #125, Excel 16.0) |
| `typeof-is-always-false` | 'TypeOf ... Is' is always False | warning | runtime risk | MS-VBAL 5.6 (TypeOf...Is expression) |
| `unallocated-dynamic-array-access` | Dynamic array is not allocated | error | runtime error | VBE runtime error 9: Subscript out of range |
| `undefined-label` | Label not defined | error | compile error | MS-VBAL 5.4.1 / VBE compiler: Label not defined |
| `variant-value-misuse` | Variant value used as another kind | error | runtime error | VBE runtime errors 424 and 13 on a Variant holding a scalar or an array (issue #121, Excel 16.0) |

## Style (13)

| Code | Title | Default | Kind | Spec reference |
| --- | --- | --- | --- | --- |
| `analysis-suppression-directive` | Invalid analysis suppression directive | warning | style-policy | Analysis suppression directive comment syntax |
| `doc-param-missing` | Doc comment does not describe a parameter | warning | style-policy | Doc comment syntax |
| `doc-param-unknown` | Doc comment describes a parameter the declaration does not have | warning | style-policy | Doc comment syntax |
| `doc-returns-missing` | Doc comment does not describe the return value | warning | style-policy | Doc comment syntax |
| `doc-returns-unexpected` | Doc comment describes a return value the declaration does not have | warning | style-policy | Doc comment syntax |
| `doc-tag-duplicate` | Doc comment repeats a tag | warning | style-policy | Doc comment syntax |
| `doc-tag-unclosed` | Doc comment tag is not closed | warning | style-policy | Doc comment syntax |
| `option-explicit-missing` | Option Explicit is not specified | warning | style-policy | MS-VBAL 5.2.4.1.1 |
| `unreachable-code` | Code is never reached | information | style-policy | MS-VBAL 5.4.1.3 (Exit), 5.4.1.4 (GoTo), 5.4.4 (Resume) plus dead-code policy |
| `unused-procedure` | Private procedure is never called | information | style-policy | MS-VBAL 5.3.1.1 (Private procedure visibility) plus dead-code policy |
| `unused-variable` | Variable or constant is never used | information | style-policy | MS-VBAL 5.2.3 / 5.4.3.1 (declaration scope) plus dead-code policy |
| `variable-never-read` | Variable is assigned but never read | information | style-policy | MS-VBAL 5.4.3 (assignment) plus dead-code policy |
| `vba-test-directive` | Invalid VBA test directive | warning | style-policy | VBA test directive comment syntax |

## Syntax (24)

| Code | Title | Default | Kind | Spec reference |
| --- | --- | --- | --- | --- |
| `bracketed-variable-name` | Bracketed name in a variable declaration | error | compile error | MS-VBAL 3.3.5.3 foreign-name; VBE "Syntax error" on Dim [name] (issue #124, Excel 16.0) |
| `call-requires-parens` | Call statement requires parentheses around arguments | error | compile error | MS-VBAL 5.4.2.1 |
| `call-statement-forbids-parens` | Standalone zero-argument call cannot use empty parentheses | error | compile error | MS-VBAL 5.4.2.1 |
| `call-statement-multi-arg-parens` | Standalone call cannot parenthesize multiple arguments | error | compile error | MS-VBAL 5.4.2.1 |
| `date-literal-invalid` | Date literal the VBE refuses | error | compile error | MS-VBAL 3.3.3 date tokens; VBE "Syntax error" (issue #133, Excel 16.0) |
| `directive-trailing-statement` | Code after a compiler directive on its line | error | compile error | MS-VBAL 3.4 conditional compilation; VBE "An # ElseIf, # Else, or # EndIf must be preceded by an # If clause" (issue #130, Excel 16.0) |
| `else-branch-order` | Else branch must be final in conditional block | error | compile error | MS-VBAL 3.4 / 5.4.2.1 |
| `expression-call-requires-parens` | Function call in an expression requires parentheses around arguments | error | compile error | MS-VBAL 5.6.9 |
| `float-literal-overflow` | Floating-point literal overflows | error | compile error | MS-VBAL 3.3.2 FLOAT; VBE "Syntax error" on 1E400 (issue #125, Excel 16.0) |
| `if-missing-then` | If statement is missing Then | error | compile error | MS-VBAL 5.4.2.1 |
| `if-reserved-keyword-in-condition` | Reserved keyword in If condition | error | compile error | MS-VBAL 5.4.2.1 (If block) / 3.3.5.2 (reserved identifiers) / VBE compiler: Syntax error |
| `invalid-erase-target` | Erase target must be a variable or array name | error | compile error | MS-VBAL Erase statement |
| `invalid-explicit-call-target` | Invalid explicit Call target | error | compile error | VBE compiler: Syntax error |
| `invalid-expression-syntax` | Invalid expression syntax | error | compile error | MS-VBAL 5.6 / VBE compiler: Syntax error |
| `invalid-line-continuation` | Invalid line continuation | error | compile error | MS-VBAL 3.2.2 |
| `line-too-long` | Physical line over 1023 characters | error | compile error | VBE line limit: 1023 characters compile, 1024 are refused (issue #133, Excel 16.0) |
| `open-missing-for` | 'Open' statement mode without 'For', or no 'As' clause | error | compile error | MS-VBAL 5.4.5.1 (Open statement) / VBE compiler |
| `rem-after-then` | Rem after Then | error | compile error | MS-VBAL 3.3.1 comment / 5.4.2.9 single-line If; VBE "Syntax error" (issue #125, Excel 16.0) |
| `statement-outside-procedure` | Statement outside procedure | error | compile error | MS-VBAL 5.2 / 5.4 |
| `stray-character` | Character VBA does not use | error | compile error | MS-VBAL 3.3.1 special-token; VBE "Syntax error" on ; ` { } @ ~ \| and a non-breaking space (issue #132, Excel 16.0) |
| `suffixed-literal-overflow` | Type-suffixed literal out of range | error | compile error | MS-VBAL 3.3.2 (number tokens / type suffixes) / VBE compiler |
| `typeof-missing-operand` | 'TypeOf' requires an object expression | error | compile error | MS-VBAL 5.6 (TypeOf...Is) / VBE compiler |
| `unbalanced-parens` | Unbalanced parentheses | error | compile error | MS-VBAL 3.3.1 |
| `unterminated-string` | Unterminated string literal | error | compile error | MS-VBAL 3.3.4 |
