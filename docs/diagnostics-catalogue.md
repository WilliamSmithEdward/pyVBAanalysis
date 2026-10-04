# Diagnostic catalogue

The diagnostic codes pyVBAanalysis can emit, generated from the rule metadata (`tools/generate_diagnostics_catalogue.py`). This table lists the 219 rule-metadata codes across 6 categories. A further 3 structural block-balance codes (`mismatched-end-keyword`, `missing-block-closer`, `unmatched-block-closer`) are emitted by the parser pass and are not in the metadata table, for a full set of 222 codes.

Each code is reported only when it is provably correct; anything unknown or ambiguous stays quiet (the no-false-positive discipline). The **kind** column says what a code means: a *compile error* is rejected by the VBE compiler, a *runtime error* is a deterministic Run-time error, a *runtime risk* is a likely fault, and *style* is advisory.

Override a code's severity with `AnalyzeModuleOptions.severity_overrides` (or the `severity_overrides` argument of `analyze_project` / the reader functions), keyed by code. Use `"off"`, `"information"`, `"warning"`, or `"error"`; the allowed values per code are constrained by policy (some codes can be downgraded but not disabled). See [docs/usage.md](usage.md).

## Declaration (55)

| Code | Title | Default | Kind | Spec reference |
| --- | --- | --- | --- | --- |
| `array-bound-not-constant` | Dim bound or String length names a variable | error | compile error | MS-VBAL 5.2.3.1 array bounds; VBE "Constant expression required" (issue #216, Excel 16.0) |
| `array-parameter-form` | Array parameter cannot be ByVal or Optional | error | compile error | MS-VBAL 5.3.1.5; VBE "Array argument must be ByRef" / "Optional argument must be Variant or intrinsic type with a default value" (issue #124, Excel 16.0) |
| `byval-udt-parameter` | User-defined type parameter cannot be ByVal | error | compile error | MS-VBAL 5.3.1 (parameter passing) / VBE compiler |
| `circular-declaration-dependency` | Declaration depends on itself through a cycle | error | compile error | VBE "Circular dependencies between modules" (issue #211, Excel 16.0) |
| `const-evaluation-error` | Const value that does not evaluate | error | compile error | MS-VBAL 5.2.3.1 constant expressions; VBE compile errors Division by zero and Type mismatch (issue #235, Excel 16.0) |
| `const-invalid-type` | Const declared As a type a constant cannot have | error | compile error | MS-VBAL 5.2.3.2 constant declarations; VBE "Invalid data type for constant" and "Expected: type name" (issue #212, Excel 16.0) |
| `const-overflow` | Const value overflows | error | compile error | MS-VBAL 5.6.9.3; VBE compile error: Overflow (issue #116, Excel 16.0) |
| `const-value-not-constant` | Const value must be a constant expression | error | compile error | MS-VBAL 5.2.4 (Const declaration value is a constant-expression) |
| `declaration-forward-reference` | Declaration refers to one further down the module | error | compile error | MS-VBAL 5.2.3.2 constant declarations; VBE "Constant expression required" and "Forward reference to user-defined type" (issue #211, Excel 16.0) |
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
| `empty-enum` | Empty Enum | error | compile error | MS-VBAL 5.2.3.4 Enum declarations; VBE "Enum without members not allowed" (issue #212, Excel 16.0) |
| `empty-type` | User-defined Type has no members | error | compile error | MS-VBAL 5.2.3.3 (UDT declaration) |
| `enum-member-not-constant` | Enum member value must be a constant expression | error | compile error | MS-VBAL 5.2.3.4 (Enum member value is a constant-expression) |
| `enum-member-type-mismatch` | Enum member value is a string | error | compile error | MS-VBAL 5.2.3.4 (Enum member values are Long); VBE "Type mismatch" (issue #210, Excel 16.0) |
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
| `private-type-in-public-signature` | Private Enum or Type in a public signature | error | compile error | VBE "Private Enum and user defined types cannot be used as parameters or return types for public procedures, public data members, or fields of public user defined types" (issue #212, Excel 16.0) |
| `property-accessor-signature-mismatch` | Property accessors have incompatible signatures | error | compile error | MS-VBAL 5.3.1.4 |
| `property-let-object-value` | Property Let value parameter must not be object reference | error | compile error | MS-VBAL 5.3.1.4 |
| `property-set-scalar-value` | Property Set value parameter must be object reference | error | compile error | MS-VBAL 5.3.1.4 |
| `property-setter-missing-value` | Property setter is missing value parameter | error | compile error | MS-VBAL 5.3.1.4 |
| `property-setter-return-type` | Property setter cannot declare a return type | error | compile error | MS-VBAL 5.3.1.4 |
| `required-param-after-optional` | A required parameter cannot follow an Optional parameter | error | compile error | MS-VBAL 5.3.1.5 |
| `too-many-array-dimensions` | Array has more than 60 dimensions | error | compile error | MS-VBAL 5.2.3.1 (array declaration) / VBE compiler |
| `too-many-parameters` | Procedure has more than 60 parameters | error | compile error | MS-VBAL 5.3.1 (procedure parameters) / VBE compiler |
| `type-declaration-character-as-clause` | Invalid type-declaration character with As | error | compile error | MS-VBAL 5.2.3.1 / 5.3.1 type-declaration characters |
| `type-enum-name-conflict` | Type and Enum share a name | error | compile error | VBE "Ambiguous name detected" (issue #212, Excel 16.0) |
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

## Semantic (108)

| Code | Title | Default | Kind | Spec reference |
| --- | --- | --- | --- | --- |
| `addressof-misuse` | AddressOf where the VBE refuses it | error | compile error | MS-VBAL 5.6.16.8 (AddressOf expressions) / VBE compile errors |
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
| `array-target-assignment` | Can't assign to array | error | compile error | VBE "Can't assign to array": a fixed-size array, a scalar into a dynamic array, or an array of another element type (issue #194, Excel 16.0) |
| `array-temporarily-locked` | Array is locked while it is resized or erased | error | runtime error | VBE runtime error 10: This array is fixed or temporarily locked |
| `assignment-object-type-mismatch` | Object assignment type mismatch | error | runtime error | MS-VBAL 5.4.3 / Set statement; VBE runtime error 13: Type mismatch |
| `assignment-to-procedure-name` | Assignment to a Sub or typed Function from outside it | error | compile error | VBE "Expected Function or variable" and "Function call on left-hand side of assignment must return Variant or Object" (issue #213, Excel 16.0) |
| `assignment-type-mismatch` | Assignment type mismatch | error | runtime error | MS-VBAL 5.4.3 / runtime type coercion and numeric overflow |
| `byref-argument-type-mismatch` | ByRef argument type mismatch | error | compile error | MS-VBAL 5.3.1 / VBE compiler: ByRef argument type mismatch |
| `case-outside-select` | Case statement outside Select Case | error | compile error | MS-VBAL 5.4.2.4 |
| `collection-add-argument` | Collection Add argument refused | error | runtime error | VBE runtime errors 13 and 5 from VBA.Collection.Add: a key that is not a string, Before with After (issue #219, Excel 16.0) |
| `collection-index-out-of-range` | Collection index out of range | error | runtime error | VBE runtime errors 5 and 9 on VBA.Collection (issue #121, Excel 16.0) |
| `collection-key-in-use` | Collection key already in use | error | runtime error | VBE runtime error 457 on VBA.Collection (issue #121, Excel 16.0) |
| `collection-key-not-found` | Collection key not found | error | runtime error | VBE runtime error 5 on VBA.Collection (issue #121, Excel 16.0) |
| `collection-operand` | Collection or object without a value used as an operand | error | compile error | VBE "Argument not optional" on a Collection whose default member Item takes an index (issue #125, Excel 16.0); "Type mismatch" on a Word Paragraph, whose default member Range holds an object (issue #462, Word 16.0) |
| `const-assignment` | Assignment to a constant | error | compile error | MS-VBAL 5.4.3.1 |
| `division-by-zero` | Division by zero | error | runtime error | MS-VBAL 5.6 / runtime division by zero |
| `duplicate-case-else` | Duplicate Case Else in Select Case | error | compile error | MS-VBAL 5.4.2.10 (Select Case) |
| `duplicate-label` | Duplicate procedure label | error | compile error | MS-VBAL 5.4.1 |
| `else-without-if` | 'Else'/'ElseIf' outside an If block | error | compile error | MS-VBAL 5.4.2.1 (If block) / VBE compiler |
| `empty-file-path` | Empty file path | error | runtime error | VBE runtime errors 53, 75 and 76 on an empty path (issue #262, Excel 16.0) |
| `erase-requires-array` | Erase target must be array or Variant | error | compile error | MS-VBAL Erase statement / VBE compiler: Expected array |
| `event-handler-signature` | Event handler declared unlike its event | error | compile error | VBE "Procedure declaration does not match description of event or procedure having the same name"; events from the Office and MSForms type libraries (issue #195, Excel 16.0) |
| `exit-outside-block` | Loop exit statement outside matching loop | error | compile error | MS-VBAL 5.4.1.3 |
| `exit-wrong-proc` | Exit statement does not match the enclosing procedure | error | compile error | MS-VBAL 5.4.1.3 |
| `file-already-open` | File number opened twice | error | runtime error | VBE runtime error 55: File already open (issue #123, Excel 16.0) |
| `file-mode-mismatch` | File statement the open mode forbids | error | runtime error | VBE runtime error 54: Bad file mode (issue #123, Excel 16.0) |
| `file-number-zero` | File number outside 1 to 512 | error | runtime error | VBE runtime error 52: Bad file name or number (issues #123 and #262, Excel 16.0) |
| `file-read-past-end` | Reading a file the procedure created empty | error | runtime error | VBE runtime error 62: Input past end of file (issue #262, Excel 16.0) |
| `file-record-zero` | Record number 0 | error | runtime error | VBE runtime error 63: Bad record number (issue #123, Excel 16.0) |
| `file-used-after-close` | File statement on a closed file number | error | runtime error | VBE runtime error 52: Bad file name or number (issue #123, Excel 16.0) |
| `fixed-array-redim` | Fixed-size array cannot be ReDimmed | error | compile error | MS-VBAL ReDim statement |
| `for-counter-overflow` | For counter overflows after its last pass | error | runtime error | MS-VBAL 5.4.2.3; VBE runtime error 6: Overflow (issue #116, Excel 16.0) |
| `for-counter-type` | For counter of a type that cannot count | error | compile error | MS-VBAL 5.4.2.3; VBE "Type mismatch" (issue #213, Excel 16.0) |
| `for-each-control-variable-type` | For Each control variable must be Variant or Object | error | compile error | MS-VBAL 5.4.2.5 |
| `for-each-source-type` | For Each source must be collection or array | error | compile error | MS-VBAL 5.4.2.5 / VBE compiler: For Each may only iterate over a collection object or an array |
| `for-variable-in-use` | For control variable already in use | error | compile error | MS-VBAL 5.4.2.3; VBE "For control variable already in use" (issue #213, Excel 16.0) |
| `formula-string-unparsed` | Formula String Excel cannot parse | warning | runtime risk | Excel 1004 on a formula String with a parenthesis or string left open, or an operator last; a cell formatted as Text takes it (issue #276, Excel 16.0) |
| `handler-fall-through` | Execution falls into an error handler that re-raises | error | runtime error | VBE runtime error 5 from Err.Raise Err.Number with Err.Number 0 (issue #117, Excel 16.0) |
| `host-argument-out-of-range` | Host argument out of range | error | runtime error | Office object model: 1-based collections, cell rows and columns from 1 (issue #122, Excel/Word/PowerPoint 16.0) |
| `host-property-value-out-of-range` | Host property value out of range | error | runtime error | Office object model: property ranges the host refuses, measured in Excel/Word/PowerPoint 16.0 (issue #204) |
| `implements-member-missing` | Interface member not implemented | error | compile error | MS-VBAL 5.2.4.1; VBE "Object module needs to implement ... for interface ..." (issue #125, Excel 16.0) |
| `implements-member-signature` | Interface member implemented with another signature | error | compile error | MS-VBAL 5.2.4.1; VBE "Procedure declaration does not match description of event or procedure having the same name" (issue #125, Excel 16.0) |
| `invalid-assignment-target` | Cannot assign to a literal value | error | compile error | MS-VBAL 5.4.3 (assignment) / VBE compiler |
| `invalid-paramarray-use` | A ParamArray resized, erased or passed ByRef | error | compile error | VBE "Invalid ParamArray use": ReDim, Erase, or a ByRef argument of a ParamArray, and one declared with no parentheses (issue #445, Excel 16.0) |
| `invalid-property-use` | Invalid use of property | error | compile error | VBE "Invalid use of property" (issue #266, Excel 16.0) |
| `is-operator-non-object` | 'Is' operator requires object operands | error | compile error | MS-VBAL 5.6 (Is operator) |
| `late-bound-friend-member` | Friend member reached through a late-bound receiver | error | runtime error | VBE runtime error 438: Object does not support this property or method |
| `late-bound-object-state` | Late-bound object used before it is opened | error | runtime error | ADODB error 3704: Operation is not allowed when the object is closed (issue #477, Excel 16.0) |
| `longlong-narrowing` | LongLong narrowed implicitly in 64-bit VBA | error | compile error | VBA 64-bit: a LongLong or LongPtr converts to no narrower whole-number type implicitly; VBE compile error Type mismatch |
| `lset-type-mismatch` | LSet or RSet that cannot copy into its target | error | compile error | VBE "Type mismatch" on LSet between two user-defined types (issue #253, Excel 16.0); "LSet allowed only on strings and user-defined types" and "RSet allowed only on strings" (issue #451, Excel 16.0) |
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
| `null-directive-condition` | #If condition is Null | error | compile error | MS-VBAL 3.4 conditional compilation; VBE "Invalid use of Null" (issue #208, Excel 16.0) |
| `object-default-value` | Object read as a value has no default member | error | runtime error | VBE runtime errors 438 and 450 reading an object with no default member, or one that needs an argument (issue #183, Excel 16.0) |
| `object-used-after-delete` | Object used after it is deleted or closed | error | runtime error | Excel object model: a deleted sheet, shape or name, a closed workbook or an unlisted table raises 424, 1004 or -2147221080 |
| `object-variable-not-set` | Object variable not set | error | runtime error | VBE runtime error 91: Object variable or With block variable not set |
| `paste-with-nothing-copied` | PasteSpecial after the clipboard was emptied | error | runtime error | Excel runtime error 1004: PasteSpecial method of Range class failed, after Application.CutCopyMode = False with no Copy or Cut since (issue #308, Excel 16.0) |
| `raiseevent-argument-count` | RaiseEvent passes the wrong number of arguments | error | compile error | VBE "Wrong number of arguments or invalid property assignment" and "Argument not optional" (issue #213, Excel 16.0) |
| `raiseevent-undeclared-event` | RaiseEvent target is not declared | error | compile error | MS-VBAL RaiseEvent statement: event name resolution |
| `readonly-member-assignment` | Assignment to a read-only member | error | compile error | VBE compiler: Can't assign to read-only property; Wrong number of arguments or invalid property assignment; Assignment to constant not permitted |
| `recursive-property-accessor` | Property procedure calls itself | error | runtime error | VBE runtime error 28: Out of stack space (issue #117, Excel 16.0) |
| `redim-impossible-bounds` | ReDim lower bound is greater than upper bound | error | runtime error | MS-VBAL ReDim statement / VBE compiler runtime error 9 |
| `redim-preserve-dimension-change` | ReDim Preserve can only resize the last dimension | error | runtime error | MS-VBAL ReDim Preserve statement |
| `redim-type-change` | ReDim changes an array's element type | error | compile error | MS-VBAL 5.4.3.3 ReDim; VBE "Can't change data types of array elements" (issue #212, Excel 16.0) |
| `resume-without-error` | Resume with no error handler installed | error | runtime error | VBE runtime error 20: Resume without error (issue #117, Excel 16.0) |
| `return-without-gosub` | Execution falls into a GoSub target | error | runtime error | VBE runtime error 3: Return without GoSub (issue #117, Excel 16.0) |
| `runtime-argument-value` | Invalid runtime argument value | error | runtime error | MS-VBAL 5.6 / VBA runtime argument bounds and VBE compiler runtime error 5 |
| `runtime-conversion-value` | Invalid runtime conversion value | error | runtime error | MS-VBAL 5.6 / VBA runtime conversion and VBE compiler runtime error 13 |
| `runtime-member-not-found` | Member not found at run time | error | runtime error | VBE runtime error 438: Object doesn't support this property or method (issue #121, Excel 16.0) |
| `scalar-indexed` | A value that is no array, indexed | error | compile error | VBE "Expected array" on a number, string or user-defined type local or Type field given a subscript (issue #417, Excel 16.0) |
| `scalar-member-access` | Member access on scalar variable | error | compile error | VBE compiler: Invalid qualifier / Syntax error |
| `scalar-redim` | Scalar variable cannot be ReDimmed | error | compile error | MS-VBAL ReDim statement / VBE compiler: Expected array |
| `set-required` | Object assignment requires Set | error | runtime error | VBE runtime error 91: Object variable or With block variable not set (MS-VBAL 5.4.3 / Set statement) |
| `set-requires-object` | Set assignment requires an object variable | error | compile error | MS-VBAL 5.4.3 |
| `sheet-name-invalid` | Sheet name Excel refuses | error | runtime error | Excel runtime error 1004: invalid name for a sheet or chart (issue #122, Excel 16.0) |
| `sheet-not-in-workbook` | Sheet the workbook does not have | error | runtime error | Excel runtime error 9: Subscript out of range from ThisWorkbook.Sheets, Worksheets or Charts given a name or index the workbook lacks (issue #229, Excel 16.0) |
| `string-arithmetic-coercion` | Nonnumeric string in numeric expression | error | runtime error | MS-VBAL 5.6 / runtime type coercion |
| `sub-used-as-value` | Sub used as a value | error | compile error | VBE "Expected Function or variable" (issue #125, Excel 16.0) |
| `type-suffix-mismatch` | Type-declaration character does not match the declared type | error | compile error | MS-VBAL 3.3.5.2; VBE "Type-declaration character does not match declared data type" (issue #216, Excel 16.0) |
| `typeof-is-always-false` | 'TypeOf ... Is' is always False | warning | runtime risk | MS-VBAL 5.6 (TypeOf...Is expression) |
| `udt-value-mismatch` | User-defined type used as a single value | error | compile error | VBE "Type mismatch" (issue #213, Excel 16.0) |
| `udt-variant-coercion` | User-defined type handed to a Variant | error | compile error | VBE "Only user-defined types defined in public object modules can be coerced to or from a variant or passed to late-bound functions" (issue #213, Excel 16.0) |
| `unallocated-dynamic-array-access` | Dynamic array is not allocated | error | runtime error | VBE runtime error 9: Subscript out of range; 92 for For Each |
| `unbounded-recursion` | Sub or Function calls itself on every run | error | runtime error | VBE runtime error 28: Out of stack space (issue #240, Excel 16.0) |
| `undefined-label` | Label not defined | error | compile error | MS-VBAL 5.4.1 / VBE compiler: Label not defined |
| `unusable-declare` | Declare that fails on every call | error | runtime error | VBE runtime errors 48, 49, 452 and 453 on a Declare whose calling convention, Lib or Alias no call can use (issue #254, Excel 16.0) |
| `variable-required` | Constant where a variable is required | error | compile error | VBE "Variable required - can't assign to this expression" (issue #216, Excel 16.0) |
| `variant-value-misuse` | Variant value used as another kind | error | runtime error | VBE runtime errors 424 and 13 on a Variant holding a scalar or an array (issue #121, Excel 16.0) |
| `with-scalar-target` | With on a scalar | error | compile error | MS-VBAL 5.4.2.21 With; VBE "With object must be user-defined type, Object, or Variant" (issue #213, Excel 16.0) |
| `wrong-number-of-dimensions` | Subscript count differs from the array's dimensions | error | compile error | VBE "Wrong number of dimensions" (issue #248, Excel 16.0) |

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

## Syntax (36)

| Code | Title | Default | Kind | Spec reference |
| --- | --- | --- | --- | --- |
| `bracketed-variable-name` | Bracketed name in a variable declaration | error | compile error | MS-VBAL 3.3.5.3 foreign-name; VBE "Syntax error" on Dim [name] (issue #124, Excel 16.0) |
| `call-requires-parens` | Call statement requires parentheses around arguments | error | compile error | MS-VBAL 5.4.2.1 |
| `call-statement-forbids-parens` | Standalone zero-argument call cannot use empty parentheses | error | compile error | MS-VBAL 5.4.2.1 |
| `call-statement-multi-arg-parens` | Standalone call cannot parenthesize multiple arguments | error | compile error | MS-VBAL 5.4.2.1 |
| `date-literal-invalid` | Date literal the VBE refuses | error | compile error | MS-VBAL 3.3.3 date tokens; VBE "Syntax error" (issue #133, Excel 16.0) |
| `directive-trailing-statement` | Code after a compiler directive on its line | error | compile error | MS-VBAL 3.4 conditional compilation; VBE "An # ElseIf, # Else, or # EndIf must be preceded by an # If clause" (issue #130, Excel 16.0) |
| `else-branch-order` | Else branch must be final in conditional block | error | compile error | MS-VBAL 3.4 / 5.4.2.1 |
| `event-parameter-form` | Event parameter is Optional or a ParamArray | error | compile error | MS-VBAL 5.2.4.3 event declarations; VBE "Syntax error" (issue #212, Excel 16.0) |
| `expression-call-requires-parens` | Function call in an expression requires parentheses around arguments | error | compile error | MS-VBAL 5.6.9 |
| `float-literal-overflow` | Floating-point literal overflows | error | compile error | MS-VBAL 3.3.2 FLOAT; VBE "Syntax error" on 1E400 (issue #125, Excel 16.0) |
| `if-missing-then` | If statement is missing Then | error | compile error | MS-VBAL 5.4.2.1 |
| `if-reserved-keyword-in-condition` | Reserved keyword in If condition | error | compile error | MS-VBAL 5.4.2.1 (If block) / 3.3.5.2 (reserved identifiers) / VBE compiler: Syntax error |
| `invalid-erase-target` | Erase target must be a variable or array name | error | compile error | MS-VBAL Erase statement |
| `invalid-explicit-call-target` | Invalid explicit Call target | error | compile error | VBE compiler: Syntax error |
| `invalid-expression-syntax` | Invalid expression syntax | error | compile error | MS-VBAL 5.6 / VBE compiler: Syntax error |
| `invalid-line-continuation` | Invalid line continuation | error | compile error | MS-VBAL 3.2.2 |
| `invalid-line-number` | Line number the VBE refuses | error | compile error | MS-VBAL 5.4.1.1 line numbers: 0 to 2147483647, at the start of a line; VBE "Syntax error" (issues #210 and #230, Excel 16.0) |
| `line-too-long` | Physical line over 1023 characters | error | compile error | VBE line limit: 1023 characters compile, 1024 are refused (issue #133, Excel 16.0) |
| `malformed-statement` | Line the VBE cannot parse | error | compile error | VBE compile errors on unfinished lines: an unknown #directive, an unclosed [bracket, Sub with no name, a word after a parameter, an Enum line that is no member (issue #234, Excel 16.0) |
| `named-argument-not-allowed` | Named argument to a function that takes none | error | compile error | VBE "Syntax error" (issue #216, Excel 16.0) |
| `open-missing-for` | 'Open' statement mode without 'For', or no 'As' clause | error | compile error | MS-VBAL 5.4.5.1 (Open statement) / VBE compiler |
| `optional-property-value` | Property Let or Set value is Optional | error | compile error | MS-VBAL 5.3.1.7 property parameters; VBE "Syntax error" (issue #212, Excel 16.0) |
| `paramarray-passing-mode` | ByVal or ByRef on a ParamArray | error | compile error | MS-VBAL 5.3.1.5 ParamArray; VBE "Expected: identifier" (issue #213, Excel 16.0) |
| `rem-after-statement` | Rem after a statement | error | compile error | MS-VBAL 3.3.1 comment / 5.4.2.9 single-line If; VBE "Syntax error", or "Expected: end of statement" at module level (issue #231, Excel 16.0) |
| `rem-after-then` | Rem after Then | error | compile error | MS-VBAL 3.3.1 comment / 5.4.2.9 single-line If; VBE "Syntax error" (issue #125, Excel 16.0) |
| `reserved-keyword-in-expression` | Statement keyword where a value goes | error | compile error | MS-VBAL 3.3.5.2 reserved identifiers; VBE "Syntax error", or "Expected: expression" in a Const (issue #234, Excel 16.0) |
| `return-with-value` | Return with a value | error | compile error | MS-VBAL 5.4.2.14 Return; VBE "Syntax error" (issue #213, Excel 16.0) |
| `statement-before-first-case` | Statement between Select Case and the first Case | error | compile error | MS-VBAL 5.4.2.10; VBE "Statements and labels invalid between Select Case and first Case" (issue #213, Excel 16.0) |
| `statement-outside-procedure` | Statement outside procedure | error | compile error | MS-VBAL 5.2 / 5.4 |
| `static-outside-procedure` | Static at module level | error | compile error | MS-VBAL 5.2.3.1; VBE "Invalid outside procedure" (issue #216, Excel 16.0) |
| `stray-character` | Character VBA does not use | error | compile error | MS-VBAL 3.3.1 special-token; VBE "Syntax error" on ; ` { } @ ~ \| and a non-breaking space (issue #132, Excel 16.0) |
| `suffixed-literal-overflow` | Type-suffixed literal out of range | error | compile error | MS-VBAL 3.3.2 (number tokens / type suffixes) / VBE compiler |
| `type-member-without-type` | Type member without an As clause | error | compile error | MS-VBAL 5.2.3.3 UDT declarations; VBE "Statement invalid inside Type block" (issue #212, Excel 16.0) |
| `typeof-missing-operand` | 'TypeOf' requires an object expression | error | compile error | MS-VBAL 5.6 (TypeOf...Is) / VBE compiler |
| `unbalanced-parens` | Unbalanced parentheses | error | compile error | MS-VBAL 3.3.1 |
| `unterminated-string` | Unterminated string literal | error | compile error | MS-VBAL 3.3.4 |
