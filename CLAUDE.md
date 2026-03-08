# OpenVB6toCS — Project Guide for Claude

## What this project is

A VB6-to-C# source code translator targeting **ActiveX DLLs** (business logic, no UI).
Converts `.cls` and `.bas` files to idiomatic C# classes, grouped by VB6 project (`.vbp`).

## Scope

**In scope:**
- `.vbp` project files → the primary input mode; drives multi-file translation
- `.cls` class modules → C# `public class`
- `.bas` standard modules → C# `public static class`
- Fields, properties, methods, enums, constants
- Control flow: `If/ElseIf/Else`, `Select Case`, `For/Next`, `For Each`, `While`, `Do/Loop`
- Error handling: `On Error GoTo`, `On Error Resume Next`, `Err.Raise`
- COM interfaces: `Implements`
- `ByVal` / `ByRef` parameters
- `Optional` parameters, `ParamArray`
- `.csproj` generation for the translated project

**Out of scope:**
- `.frm` forms, `.ctl` UserControls — any visual/UI components
- ADO/DAO data access objects

**Deferred — to be revisited:**
- External COM references: handled with a three-tier strategy (see ROADMAP.md):
  1. Known types → translated to .NET equivalents automatically
  2. Unknown types → .NET COM Interop fallback (code still compiles; usage annotated for review)
  3. `<COMReference>` entries are written into the generated `.csproj` for any COM library
     that has no full mapping, so the interop assembly is available at build time

## Input modes

### Project mode (primary): pass a `.vbp` file
```
vb6tocs MyProject.vbp
```
Reads the `.vbp`, extracts all in-scope source files, translates them, and writes output to a new directory:
```
MyProject/
  ClassName.cs
  ModuleName.cs
  ...
  MyProject.csproj
```

### Single-file mode (diagnostic / testing): pass a `.cls` or `.bas` file
```
vb6tocs Calculator.cls
```
Prints the AST tree to stdout. Used during development only; does not produce C# output.

## VBP file format

A `.vbp` is a plain-text key=value file. Lines we care about:

| Line pattern | Action |
|---|---|
| `Class=Name; path/to/File.cls` | Translate — becomes a C# class |
| `Module=Name; path/to/File.bas` | Translate — becomes a C# static class |
| `Form=Name; path/to/File.frm` | **Skip** — warn user (UI, out of scope) |
| `UserControl=Name; path/to/File.ctl` | **Skip** — warn user (UI, out of scope) |
| `Reference=...` | Parse — extract GUID + version for COM reference handling |
| `Object=...` | **Skip** — ActiveX control reference, out of scope |
| `ExeName32=` or `Name=` | Read project name (used for output directory and `.csproj` name) |

Paths in `.vbp` are relative to the `.vbp` file's directory.

## Output: generated `.csproj`

A modern SDK-style `.csproj` is generated alongside the translated files. SDK projects auto-include all `*.cs` files, so no explicit item list is needed:

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net8.0</TargetFramework>
    <OutputType>Library</OutputType>
    <RootNamespace>ProjectName</RootNamespace>
    <AssemblyName>ProjectName</AssemblyName>
    <Nullable>enable</Nullable>
    <ImplicitUsings>enable</ImplicitUsings>
  </PropertyGroup>
</Project>
```

## File type mapping

| VB6 | C# |
|---|---|
| `.cls` class module | `public class Foo { }` |
| `.bas` standard module | `public static class Foo { }` (members are `static`) |

## Architecture

Pipeline: **Lex → Parse → Semantic Analysis → Transform → Generate → Format**

For project mode, a **VbpReader** stage runs first to discover the file list; a **CsprojWriter** stage runs after Generate to emit the `.csproj`.

All pipeline stages are in `VB6toCS.Core`. See `ROADMAP.md` for full detail.

## Project structure

```
VB6toCS.sln
src/
  VB6toCS.Core/           ← translator engine (class library)
    Lexing/
      TokenKind.cs        ← enum of all token types
      Token.cs            ← token struct (kind, text, line, column)
      Lexer.cs            ← hand-written VB6 lexer
      TokenListValidator.cs ← Option Explicit enforcement
    Parsing/
      ParseException.cs
      ModuleKind.cs
      TokenStream.cs
      Parser.cs           ← recursive descent parser
      AstPrinter.cs       ← debug output (indented tree)
      Nodes/              ← all 59 AST node types (8 files)
    Analysis/
      Symbol.cs           ← Symbol record + SymbolKind enum
      SymbolTable.cs      ← two-level (module + procedure) symbol table
      AnalysisDiagnostic.cs
      Analyser.cs         ← three-pass semantic analysis (Stage 3)
      CollectionTypeInferrer.cs ← cross-module Collection<T> element type inference
    Transformation/
      TransformDiagnostic.cs
      Transformer.cs      ← type normalization (Stage 4); error handling restructuring disabled
    CodeGeneration/
      CodeWriter.cs       ← indentation-aware line writer
      BuiltInMap.cs       ← VB6 built-in functions/constants → C# mapping table
      ComTypeMap.cs       ← static lookup: known COM indexed-member + ByRef parameter info
      CodeGenerator.cs    ← Stage 5: AST → C# source text
    Projects/
      VbProject.cs        ← data model (VbProject, VbSourceFile, VbComReference, VbSkippedFile)
      VbpReader.cs        ← parses .vbp project files
      CsprojWriter.cs     ← generates SDK-style .csproj
  VB6toCS.Cli/            ← command-line entry point
    Program.cs            ← pipeline orchestration + cross-module map builders
    CliOptions.cs         ← CLI argument parsing
samples/
  Calculator.vbp          ← sample VB6 project file
  Calculator.cls          ← sample VB6 class for testing
  SampleClass.cls         ← comprehensive .cls covering all token types
  SampleModule.bas        ← comprehensive .bas covering all token types
  syas/                   ← real-world legacy VB6 codebase (16 projects, 562 files total)
                             Final integration target: D46O003_1080.vbp
                             (136 classes, 1 module, no forms — pure ActiveX DLL)
ROADMAP.md
CLAUDE.md                 ← this file
```

## Current state

- **VBP project reading: complete.** `VbpReader` + `CsprojWriter` in `VB6toCS.Core/Projects/`.
- **Stage 1 (Lexer): complete.**
- **Stage 2 (Parser): complete.** `Parser.cs` produces `ModuleNode` AST; CLI prints indented tree.
  Validated against `D46O003_1080.vbp`: all 137 files parse successfully.
- **Stage 3 (Semantic Analysis): complete.** `Analyser.cs` in `VB6toCS.Core/Analysis/`.
  Three passes: (1) build module symbol table, (2) transform procedure bodies
  (resolve `CallOrIndexNode`, detect function returns, normalize identifier casing),
  (3) group `Property Get/Let/Set` into `CsPropertyNode`.
  Validated against `D46O003_1080.vbp`: all 137 files pass Stage 3 with 0 diagnostics.
- **Stage 4 (IR Transformation): complete.** `Transformer.cs` in `VB6toCS.Core/Transformation/`.
  Single pass: type normalization (VB6 type names → C# types in all TypeRefNodes).
  Error handling restructuring (`On Error GoTo` → `TryCatchNode`) has been **disabled** — VB6
  error handling semantics differ too much from C# try/catch to be mechanically reliable.
  `OnErrorNode`, `ResumeNode`, and error-handler `LabelNode`s are left in the body so the code
  generator emits them as comments for manual developer review.
  Validated against `D46O003_1080.vbp`: all 137 files pass Stage 4.
- **Stage 5 (C# Code Generation): complete.** `CodeGenerator.cs` + `BuiltInMap.cs` + `ComTypeMap.cs` in `VB6toCS.Core/CodeGeneration/`.
  Full C# output for all VB6 constructs. Five cross-module maps built in `Program.cs` between
  Stages 4 and 5: `collectionTypeMap`, `enumMemberMap`, `enumTypedFieldNames`, `methodParamMap`, `globalVarMap`.
  Validated against `D46O003_1080.vbp`: all 137 files translated, `.cs` files written.
  Post-release enhancements applied (all validated against D46O003_1080.vbp):
  - `ref` keyword emission at call sites for `ByRef` parameters (cross-module `methodParamMap`).
  - COM indexed-member disambiguation: `pRs.Fields(i)` → `pRs.Fields[i]` via `ComTypeMap`.
  - Type-aware string coercion for `Left`/`Right`/`Mid`/`InStr`/`Len` (`.ToString()` on non-string args).
  - VB6 date arithmetic: `Date - 60` → `DateTime.Now.Date.AddDays(-60)`.
  - `IsMissing(x)` → `(x == null)` or `(x == default)` based on parameter type.
  - `ReDim arr(n) As T` → `arr = new T[n]` (dimension expression now captured by parser).
  - Function return-type inference: functions declared `As Variant` (→ `object`) whose body
    consistently returns a single known type are emitted with that inferred type instead.
  - `CollectionTypeInferrer` extended to infer element types through local `Collection`
    intermediates (e.g. `registros` correctly typed as `Collection<Collection<clsYasField>>`).
  Post-compilation-error sprint (items 1–8 from TODO.md, all validated against D46O003_1080.vbp):
  - All string literals now emitted as `@"..."` verbatim strings — fixes `""` escape sequences
    and backslash handling simultaneously. ~1,110 errors eliminated.
  - Missing optional args: trailing args omitted; middle args → `default /* missing optional */`.
  - VB6 `Static` locals: `static` keyword removed (illegal in method bodies), comment added.
  - `ref` on literals dropped (C# forbids it); `/* was ByRef */` comment added.
  - `IndexNode` with 0 arguments now emits target only (no empty `[]`).
  - Lexer: type-suffix characters (`%&!#@$`) consumed but NOT appended for identifiers and
    number literals. Hex format: `&H1F` → `0x1F`.
  - Additional fixes: numeric labels (`10:` → `_L10:`), `Imp` operator (`(!a || b)`),
    `DateAdd`/`DateDiff`/`Choose`/`Switch` → `default /* TODO */`, nested `/* */` in block
    comments fixed, VB6 `Or` in case patterns split into two labels, `Collection` in function
    return types now resolved, `static const` → `const`, `const object` type inferred from
    literal, missing VB6 constants added, `ByRef` optional param defaults stripped,
    unqualified `Dictionary` type → `Dictionary<string, object>`.
  Result: ~2,500 errors → ~290, of which ~286 are expected (COM/cross-project dependencies).
- **Stage 6 (Roslyn formatting): not yet implemented.**

## How to run

```bash
# Translate a VB6 project — produces output directory with .cs files + .csproj
dotnet run --project src/VB6toCS.Cli -- samples/Calculator.vbp

# Single-file diagnostic mode — prints AST tree to stdout, no .cs output
dotnet run --project src/VB6toCS.Cli -- samples/Calculator.cls

# Run only up to a specific stage (0=list files, 1=lex, 2=parse, 3=analyse, 4=transform, 5=generate)
dotnet run --project src/VB6toCS.Cli -- samples/Calculator.vbp -UpToStage 3

# Specify source file encoding (default: windows-1252, fallback: Latin-1)
dotnet run --project src/VB6toCS.Cli -- samples/Calculator.vbp -Encoding utf-8
```

Note: MSB3492 errors after cleaning `obj/` dirs are transient — retry the build once.

## Stage 5 — Code generation conventions

These are critical decisions encoded in `CodeGenerator.cs`. Do not change them without understanding the VB6 semantics they encode.

### Cross-module maps (built in `Program.cs` between Stages 4 and 5)

Five maps are built from all Stage-3 ASTs and passed to `CodeGenerator.Generate()`:

| Map | Type | Purpose |
|---|---|---|
| `collectionTypeMap` | `IReadOnlyDictionary<(string module, string field), string elemType>` | Element type for `Collection<T>` fields, inferred by `CollectionTypeInferrer` |
| `enumMemberMap` | `IReadOnlyDictionary<string memberName, (string module, string enumName)>` | Qualifies bare enum member identifiers at emit time; ambiguous names excluded |
| `enumTypedFieldNames` | `IReadOnlySet<string>` | Field/property/UDT-field names whose declared type is an enum; suppresses `(int)` casts in same-enum comparisons |
| `methodParamMap` | `IReadOnlyDictionary<string moduleName, IReadOnlyDictionary<string methodName, IReadOnlyList<ParameterMode>>>` | Parameter modes for every Sub/Function/Property in every module; used to emit `ref` at call sites |
| `globalVarMap` | `IReadOnlyDictionary<string fieldName, string typeName>` | Public field names → declared type across all modules; used for cross-module type resolution in `TryResolveProcParamModes` |

### Enum emission rules

- All enums emit `: int`: `public enum e_Foo : int`
- Bare enum member identifiers are qualified via `enumMemberMap` in `MapIdentifier()`:
  - Same module → `EnumName.MemberName`
  - Different module → `ModuleName.EnumName.MemberName`

### Type comparison fixes

VB6 coerces freely between types in comparisons; C# does not. Applied in `MapBinary()` for `==`, `!=`, `<`, `>`, `<=`, `>=`:

1. **Enum vs int**: if one side is an enum member (in `enumMemberMap`) and the other is NOT in `enumTypedFieldNames` → cast enum side to `(int)`.
2. **Enum vs enum**: if the other side IS in `enumTypedFieldNames` (both sides are enum-typed) → no cast.
3. **String vs numeric**: if one side is `string` (from local type scope) and the other is numeric → convert numeric to string:
   - Numeric literals → `"1"` (quoted)
   - Other expressions → `.ToString()`

### Local type scope in `CodeGenerator`

Two dictionaries maintained for type-mismatch detection:
- `_moduleFieldTypes`: populated from `FieldNode` declarations at `GenerateModule()` start.
- `_procTypes`: reset and populated from parameters at each procedure entry (`EnterProcScope()`); `LocalDimNode` adds to it during body generation.

`TryGetSimpleType(ExpressionNode)` checks `_procTypes` first, then `_moduleFieldTypes`. Returns `null` for cross-module member accesses (not resolvable without a full type registry — known limitation).

### VB6 `Collection` → C# `Collection<T>`

- Module-level fields: type resolved via `collectionTypeMap`; falls back to `Collection<object>` with `// REVIEW:` comment.
- Locals and parameters: always `Collection<object>` with `// REVIEW:` comment.
- `New Collection` → `new()` (target-typed new — avoids C# generic invariance errors since `Collection<object>` is not assignable to `Collection<SomeType>`).
- Always adds `using System.Collections.ObjectModel;`.

### Function return pattern

Functions (and Property Get) use a `_result` local:
```csharp
ReturnType _result = default;
// FunctionReturnNode → _result = expr;
return _result;
```
`Exit Function/Sub/Property` inside a getter → `return _result;`; inside a void Sub → `return;`.

### `ref` parameter emission at call sites

`ArgsWithRef(args, modes)` in `CodeGenerator` prepends `ref` to each argument whose matching parameter is declared `ByRef` in `methodParamMap`. `TryResolveProcParamModes(target)` resolves the parameter list for a call target by:
1. Checking the current module's own methods (bare identifier calls).
2. Checking all other modules (unqualified cross-module calls).
3. For `obj.Method(...)` member-access calls: looking up `obj`'s declared type in `_procTypes` → `_moduleFieldTypes` → `globalVarMap`, then finding that type's method in `methodParamMap`.

`_moduleFieldTypes` includes both `FieldNode` declarators **and** `CsPropertyNode` return types, so property-typed variables (e.g. `M46V999` of type `clsM46V999`) resolve correctly.

### COM indexed-member disambiguation

`ComTypeMap.cs` contains a static `HashSet<(ownerType, memberName)>` of COM members that are indexed collections rather than methods. When `MapCall` encounters `CallOrIndexNode(MemberAccessNode(obj, member), [args])`, it calls `ComTypeMap.IsIndexedMember(receiverType, member)` (type from `TryGetSimpleType`). If matched, `[args]` is emitted instead of `(args)`.

Type names are matched case-insensitively, trying both the fully-qualified form (`ADODB.Recordset`) and the unqualified form (`Recordset`). Add new entries to `ComTypeMap.IndexedMembers` as new COM libraries are encountered. See ROADMAP.md Stage 5c for the planned `.tlb` reader that will replace this static table.

### ReDim statement

`ReDimNode` carries a `DimensionLists` init-property (`IReadOnlyList<IReadOnlyList<ExpressionNode>>`), one inner list per declarator. The parser populates it via `ParseReDimDeclarator` + `ParseReDimDimension`:
- `ReDim arr(n)` → dimension `[n]` → `arr = new T[n]`
- `ReDim arr(lb To ub)` → lower bound dropped, dimension `[ub]` → `arr = new T[ub]`  (C# arrays are always 0-based)
- `ReDim arr(n1, n2)` → two dimensions → `arr = new T[n1, n2]`

### Function return-type inference

`InferActualReturnType(parameters, body, declaredType)` in `CodeGenerator` fires only when `declaredType == "object"` (the result of `Variant` → `object` normalization). It pre-scans the full body with `CollectBodyTypeRefs` to build a complete locals map, then collects every `FunctionReturnNode` value with `CollectReturnValues`. If all returned expressions are identifiers whose `TypeRef` resolves to the same C# type via `TypeStr`, that type replaces `object` in both the method signature and the `_result` local declaration.

### Built-in functions and constants

All mappings live in `BuiltInMap.cs`. Key ones to remember:
- `CStr(x)` / `Str(x)` → `x.ToString()`
- `CInt(x)` / `CLng(x)` → `Convert.ToInt32(x)`
- `CDate(x)` / `CVDate(x)` → `Convert.ToDateTime(x)`
- `Val(x)` → `Convert.ToDouble(x)`
- `IIf(c,t,f)` → `(c ? t : f)`
- `Nz(v,d)` → `(v ?? d)`
- `Left/Right/Mid/InStr/Replace/Split` → corresponding `string` methods
- `vbCrLf` / `True` / `False` / `Nothing` / `Empty` / `Null` → C# literal equivalents
- Functions ending in `$` (e.g. `Mid$`) are stripped of the suffix before lookup.

## Key decisions made

- **No ANTLR, no Java.** Parser is hand-written recursive descent in C#.
- **No external parser dependencies** in `VB6toCS.Core` — pure .NET 8.
- **Roslyn** will be added later (Stage 6) for output formatting only.
- **Line continuation** (`_` at end of line) is handled silently in the lexer.
- **Comments are preserved** as tokens so they can be emitted in the C# output.
- **`End If` / `End Sub` etc.** are lexed as two separate tokens (`KwEnd` + `KwIf`) — the parser handles the pairing.
- **`.bas` modules** produce `static class` with all members `static`; module-level mutable state is flagged for human review in the output.
- **`Option Explicit` is required.** Files without it are rejected before parsing with a clear error message. Users must add it to their VB6 source before translating. This ensures every variable has an explicit declaration, which gives us a single canonical casing for each symbol.
- **Project name from `.vbp`**: the `Name=` entry is preferred; `ExeName32=` (without `.dll`) is the fallback; the filename stem is the last resort.
- **Output directory**: named after the project, created next to the `.vbp` file. Existing files are overwritten with a warning.
- **SDK-style `.csproj`**: no explicit `<Compile>` list — SDK auto-includes all `*.cs`.
- **COM Interop in `.csproj`**: for COM libraries with no full type mapping, a `<COMReference>` item is added to the `.csproj` (GUID and version extracted from the `.vbp` `Reference=` line). This lets the translated project compile immediately while the user decides whether to replace interop usages with native .NET code.
- **Skipped files are reported**: when a `.vbp` references `.frm`/`.ctl` files, the tool prints a warning listing them so the user knows what needs manual porting.

## Lexer token notes

These token kinds are defined in `TokenKind.cs` but have special handling that is important to remember:

| Token | Behaviour | User warning needed? |
|---|---|---|
| `Bang` (`!`) | Only emitted when `!` is followed by a letter (default-member access: `col!Key`). When `!` follows an identifier with no letter after it, it is consumed as a VB6 Single type-suffix and becomes part of the `Identifier` token text. | Yes — if we encounter bang-access in parsing, warn the user it maps to a dictionary/collection lookup that may need manual refactoring. |
| `Hash` (`#`) | Only emitted as a bare token when `#` is not part of a valid `#date#` literal. Appears in VB6 file I/O (`Print #1, ...`) and as a conditional compilation directive prefix (`#If`). File I/O statements are skipped with a `// TODO` comment; `#If`/`#Else`/`#End` directive lines are skipped (code in all branches is included). | No — handled gracefully. |
| `At` (`@`) | The `@` character is always consumed as a Currency type-suffix inside the number scanner (e.g. `9.99@` → `DoubleLiteral`). It is never emitted as a standalone token. No special handling needed. | No. |
| `KwEndIf` | Defined but never produced. `End If` in VB6 is always two words and is lexed as `KwEnd` + `KwIf`. The parser handles the pairing. | No. |
| `Underscore` | Defined in `ScanSymbol` but never reached — `_` is always caught first by the identifier scanner. A standalone `_` becomes `Identifier("_")`; a trailing `_` at end of line is consumed silently as line-continuation. | No. |

## Coding conventions

- Nullable reference types enabled.
- Implicit usings enabled.
- One class per file, file name matches class name.
- Lexing code lives in `VB6toCS.Core/Lexing/`, parsing in `VB6toCS.Core/Parsing/`, project-level I/O in `VB6toCS.Core/Projects/`, etc.
