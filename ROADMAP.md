# VB6 to C# Translator — Roadmap

## Scope

Targeting **ActiveX DLLs** — no forms, no controls, no visual components.

Primary input is a **`.vbp` project file**, which drives multi-file translation and produces
a ready-to-build C# project directory. Individual `.cls`/`.bas` files can also be passed
for diagnostic/testing purposes.

### File type mapping

| VB6 file | VB6 concept | C# output |
|---|---|---|
| `.vbp` | VB6 project file | output directory + `.csproj` |
| `.cls` | Class module | `public class Foo { }` |
| `.bas` | Standard module | `public static class Foo { }` |
| `.ctl` / `.frm` | Visual components | **out of scope** — skipped, user warned |

### External COM references

Currently skipped. **To be revisited** in a later stage using a three-tier strategy:

**Tier 1 — Known mapping**: the COM type has a direct .NET equivalent. Translated automatically and silently.

| COM type | C# equivalent |
|---|---|
| `Scripting.Dictionary` | `Dictionary<string, object>` |
| `Scripting.FileSystemObject` | `System.IO` namespace (File, Directory, Path, StreamReader/Writer) |
| `MSXML2.DOMDocument` | `System.Xml.XmlDocument` (partial) |
| *(more to be added as encountered)* | |

**Tier 2 — COM Interop fallback**: no known mapping, but the COM library is available. The generated code keeps the original COM type usage and a `<COMReference>` entry is added to the `.csproj` so the project still compiles. Usage sites get an annotation:
```csharp
// INTEROP: 'ADODB.Recordset' — consider replacing with a native .NET equivalent
ADODB.Recordset rs = new ADODB.Recordset();
```

**Tier 3 — Unresolvable**: the COM library path is not found and there is no mapping. A `// TODO` comment is emitted and the type is left as-is.

The `<COMReference>` item written to the `.csproj` for Tier 2:
```xml
<COMReference Include="ADODB">
  <Guid>{00000200-0000-0010-8000-00AA006D2EA4}</Guid>
  <VersionMajor>2</VersionMajor>
  <VersionMinor>0</VersionMinor>
  <Lcid>0</Lcid>
  <WrapperTool>tlbimp</WrapperTool>
  <Isolated>false</Isolated>
</COMReference>
```

The GUID and version are parsed directly from the `.vbp` `Reference=` line
(`*\G{GUID}#major.minor#lcid#path#description`).

---

## End-to-end workflow (project mode)

```
vb6tocs MyProject.vbp
```

1. **VbpReader** — parse `.vbp`, extract `Class=` and `Module=` entries (skip `Form=`, `UserControl=`, `Reference=`, `Object=`). Report skipped files to the user.
2. **Translate each file** through the full pipeline (Lex → Parse → Analyse → Transform → Generate).
3. **Write output** — create `MyProject/` directory next to the `.vbp`, write one `.cs` per source file.
4. **CsprojWriter** — emit `MyProject/MyProject.csproj` (SDK-style, auto-includes all `*.cs`).

---

## VBP file format

A `.vbp` is a plain-text `key=value` file (no sections). Key entries:

| Line pattern | Meaning |
|---|---|
| `Class=Name; File.cls` | Translate as class |
| `Module=Name; File.bas` | Translate as static class |
| `Form=Name; File.frm` | Skip (UI — out of scope) |
| `UserControl=Name; File.ctl` | Skip (UI — out of scope) |
| `Reference=*\G{GUID}#maj.min#lcid#path#desc` | Parse — extract GUID + version for COM Interop |
| `Object=...` | Skip (ActiveX control) |
| `Name=ProjectName` | Project name (preferred) |
| `ExeName32=MyDLL.dll` | Fallback project name (strip `.dll`) |

Paths in `.vbp` are relative to the `.vbp` file's directory.

---

## Generated `.csproj`

SDK-style project — no explicit `<Compile>` list; SDK auto-includes all `*.cs`:

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

---

## Pipeline Stages

### Stage 0: VBP Project Reading ✅
Parse the `.vbp` file into a `VbProject` record:
- `Name`, `VbpPath`
- `SourceFiles` — `Class=` and `Module=` entries (resolved to absolute paths)
- `ComReferences` — parsed from `Reference=*\G{GUID}#major.minor#lcid#path#desc`
- `SkippedFiles` — `Form=`, `UserControl=`, `PropertyPage=` entries with reason

`CsprojWriter` emits an SDK-style `.csproj`; COM libraries with known full mappings
(OLE Automation, Scripting Runtime) are excluded from `<COMReference>`; all others
get a `<COMReference>` block so the project compiles immediately via interop.

### Stage 1: Lexing (Tokenization) ✅
Break raw VB6 source text into a flat stream of typed tokens: keywords, identifiers, literals, operators, punctuation, and comments. Hand-written in C# — no external tool dependencies.

**Token edge cases to handle in later stages:**
- `Bang` (`!`) — emitted only when followed by a letter (default-member access). Warn the user: this pattern maps to a collection/dictionary lookup requiring manual review.
- `Hash` (`#`) as a bare token — only appears in VB6 file I/O (`Print #1, ...`), which is out of scope. Reject with a clear error if encountered.
- `At` (`@`) — consumed inside number literals as a Currency suffix; never emitted standalone. No action needed.
- `KwEndIf` — dead token; `End If` is always two tokens (`KwEnd` + `KwIf`). Parser handles the pairing.
- `Underscore` — dead token; `_` is always scanned as part of an identifier or consumed as line-continuation.

### Stage 2: Parsing → AST (Abstract Syntax Tree) ✅
A recursive descent parser consumes the token stream and produces a `ModuleNode` root.
56 AST node types in `VB6toCS.Core/Parsing/Nodes/`. `AstPrinter` renders the tree
as an indented text tree for debugging.

**Parser implementation notes:**
- `TokenStream.SingleLineIfMode` must be `true` when parsing single-line If bodies so `Else` is treated as end-of-statement.
- `ParamArray values()` — the `()` after the parameter name is consumed in `ParseParameter` and folded into `TypeRefNode.IsArray`.
- Trailing inline comments are skipped in `ExpectEndOfStatement`.
- Mismatched procedure terminators (`End Function` inside a `Property Get`) are tolerated: `ParseBody` terminates on any `End Sub/Function/Property`, and the closer consumes whichever keyword follows.
- `Declare Sub/Function` for Win32 API — parsed as `DeclareNode`; detected by the `"Declare"` identifier before `Sub`/`Function`.
- `ReDim [Preserve]` — parsed as `ReDimNode`; reuses `VariableDeclaratorNode` for each declarator.
- `As String * n` fixed-length string — `*` and length consumed in `ParseTypeRef`, treated as plain `String`.
- Qualified type names (`ADODB.Connection`) — handled by a loop in `ParseTypeRef`.
- `On Local Error GoTo` — `Local` identifier skipped before `Error` keyword.
- Comment line continuation (`' text _\n  continued`) — handled in `ScanComment`.
- File I/O statements (`Open`, `Close`, `Print #`, `Line Input`, `Get #`, `Put #`) — skipped with a `// TODO` comment (out of scope).
- `ByVal`/`ByRef` in call argument lists — prefix consumed and discarded in `ParseArgument`.
- `Error number` statement — parsed as `ErrorStatementNode`.
- Numeric line labels (`10 Statement`) — parsed as `LabelNode`; the statement on the same line is parsed next.
- Keyword-as-label-name (`ERROR:`, `SAIDA:`) — label detection in `ParseStatement` checks identifier OR keyword followed by `:`.
- `#If`/`#ElseIf`/`#Else`/`#End If` conditional compilation directives — directive line skipped; code in all branches included.

**Integration test result (D46O003_1080.vbp — 137 files):**
- 137 files: parsed OK ✅

### Stage 3: Semantic Analysis ✅
Three-pass transformation over the parsed `ModuleNode`:

**Pass 1 — Module symbol table**: collect all declarations (fields, consts, enums, UDTs, subs, functions, properties, Declare statements) into a two-level `SymbolTable` (module scope + procedure scope).

**Pass 2 — Procedure body transformation**:
- `CallOrIndexNode` resolution: if the callee is a known array symbol → `IndexNode`; otherwise stays `CallOrIndexNode` (resolved further in Stage 4/5 when type info is available for member-access chains)
- Function/Property Get return detection: `FunctionName = expr` inside a Function or Property Get body → `FunctionReturnNode`
- Identifier casing normalization: all usages of a symbol are rewritten to match the declaration-site casing
- Local `Dim` and `ReDim` statements register new symbols into the procedure scope

**Pass 3 — Property grouping**: sibling `Property Get/Let/Set` nodes sharing the same name are merged into a single `CsPropertyNode`, which maps directly to a C# property with optional getter and setter bodies.

New AST nodes introduced by Stage 3:
- `FunctionReturnNode` — return-value assignment (replaces `AssignmentNode` where LHS = function name)
- `IndexNode` — resolved array index (replaces `CallOrIndexNode` where callee is a known array)
- `CsPropertyNode` — grouped C# property (replaces separate `PropertyNode` members)

**Integration test result (D46O003_1080.vbp — 137 files):**
- 137 files: Stage 3 complete, 0 warnings ✅

### Stage 4: IR Transformation / Normalization ✅
Two passes over the Stage-3 AST:

**Pass 1 — Type normalization**: maps VB6 primitive type names to C# equivalents in every `TypeRefNode` in the AST (fields, locals, parameters, return types, UDT fields, Declare statements).

| VB6 type | C# type |
|---|---|
| `String` | `string` |
| `Integer` | `int` |
| `Long` | `int` |
| `Single` | `float` |
| `Double` | `double` |
| `Boolean` | `bool` |
| `Byte` | `byte` |
| `Date` | `DateTime` |
| `Currency` | `decimal` |
| `Variant` | `object` |
| `Object` | `object` |

COM types, enum names, and user-defined type names are left unchanged (handled by code gen / interop).

**Pass 2 — Error handling restructuring**: transforms the VB6 single-handler pattern into a `TryCatchNode`:
- `On Error GoTo label` / `[guarded body]` / `Exit Sub|Function` / `label:` / `[handler body]` → `TryCatchNode(TryBody, "_ex", CatchBody)`
- Trailing `Exit Sub/Function` stripped from TryBody (implicit in try block)
- Trailing `Resume Next` stripped from CatchBody (C# fall-through is equivalent)
- `Resume` (retry) and `Resume label` emit diagnostics — no C# equivalent
- Multiple handlers, label inside nested block, or missing label → diagnostic + no transform

New AST node: `TryCatchNode(TryBody, CatchVariable, CatchBody)`

**Integration test result (D46O003_1080.vbp — 137 files):**
- 137 files: Stage 4 complete, 33 diagnostics (all flagging real patterns needing human review) ✅

Deferred to Stage 5 (handled directly in code generation):
- `Set obj = expr` → emit `obj = expr` (drop `Set`)
- `obj = Nothing` → emit `obj = null`
- `New Foo` → emit `new Foo()`
- `Me` → emit `this`
- Operator mapping (`And`/`Or`/`Not`, `\`, `^`, `&`)
- VB6 string/math intrinsics (`Left`, `Mid`, `Len`, `InStr`, `CStr`, `CInt`, …)
- `Err.Raise` → `throw new Exception(...)`
- Module-level state in `.bas` → `private static` fields (flagged for review)

### Stage 5: C# Code Generation ✅
Walk the Stage-4 AST and emit C# source text. One `.cs` file per source module.

Implemented in `VB6toCS.Core/CodeGeneration/CodeGenerator.cs` + `BuiltInMap.cs` + `CodeWriter.cs`.

Between Stages 4 and 5 (in `Program.cs`), three cross-module maps are built from the full list of Stage-3 ASTs:
- **`collectionTypeMap`** (`CollectionTypeInferrer`) — infers the element type `T` for `Collection<T>` fields by scanning `.Add()` call sites across all modules, then propagating via assignment chains and property return types.
- **`enumMemberMap`** — maps every public enum member name to `(moduleName, enumName)`. Ambiguous names (same member in two different modules) are excluded.
- **`enumTypedFieldNames`** — set of field/property/UDT-field names whose declared VB6 type is an enum. Used to suppress casts in enum-vs-enum comparisons.

All three are passed as optional parameters to `CodeGenerator.Generate()`.

#### Class / module structure

| VB6 | C# output |
|---|---|
| `.cls` class module | `public class Foo { }` |
| `.bas` standard module | `public static class Foo { }` — all members `static` |
| `Implements IFoo` | `: IFoo` on the class declaration |

#### Enums

- All enums emit `: int` (`public enum e_Foo : int`).
- **Cross-module member qualification** — bare enum member identifiers are qualified at emit time using `enumMemberMap`:
  - Same module as the reference site → `EnumName.MemberName`
  - Different module → `ModuleName.EnumName.MemberName`

#### Type conversions in comparisons

VB6 compares string/integer/enum values freely (implicit coercion). C# does not. Three rules are applied in `MapBinary()` for comparison operators (`==`, `!=`, `<`, `>`, `<=`, `>=`):

1. **Enum vs int field** — when one side is a known enum member and the other is NOT an enum-typed field (checked via `enumTypedFieldNames`), the enum side is cast to `(int)`.
2. **Enum vs enum field** — when the other side IS an enum-typed field (both sides are the same enum type), no cast is applied.
3. **String vs numeric** — when one side is known to be `string` (from the local type scope) and the other is a numeric type, the numeric side is converted to string:
   - Integer/double literals → wrapped in quotes: `1` → `"1"`
   - Variables/expressions → `.ToString()`
   - Direction is always numeric→string (parsing a string to int can throw).

The **local type scope** is maintained during generation:
- `_moduleFieldTypes`: module-level field types, populated once at `GenerateModule()`.
- `_procTypes`: procedure parameters + locals, reset at each Sub/Function/Property accessor entry. `LocalDimNode` adds to it during body walk.
- `TryGetSimpleType()` checks `_procTypes` first, then `_moduleFieldTypes`. Returns `null` for unknown expressions (e.g., cross-module member accesses — not resolvable without a full type registry).

#### VB6 Collection → C# Collection\<T\>

- `Collection` fields: type resolved to `Collection<T>` where `T` comes from `collectionTypeMap`; falls back to `Collection<object>` with a `// REVIEW:` comment.
- `Collection` locals and parameters: always `Collection<object>` with `// REVIEW:`.
- `New Collection` → `new()` (C# 9 target-typed new — avoids generic invariance errors).
- Requires `using System.Collections.ObjectModel;` — added automatically.

#### Function return pattern

Functions use a `_result` local variable:
```csharp
ReturnType _result = default;
// body — FunctionReturnNode emits: _result = expr;
return _result;
```
`Exit Function/Sub/Property` inside a getter emits `return _result;`; inside a void Sub emits `return;`.

#### Error handling

| VB6 | C# |
|---|---|
| `TryCatchNode` (from Stage 4) | `try { } catch (Exception _ex) { }` |
| `Err.Raise num, src, desc` | `throw new Exception(desc); // Err.Raise` |
| `Err.Number` | `_ex.HResult` |
| `Err.Description` | `_ex.Message` |
| `Err.Source` | `_ex.Source ?? "" /* Err.Source */` |
| `Err.Clear()` | `/* Err.Clear() */` (no-op comment) |
| `On Error Resume Next` | comment — errors suppressed; review manually |
| `On Error GoTo 0` | comment — reset error handler |
| Unstructured `On Error GoTo` (not restructured) | comment |
| `Resume Next` in catch | stripped by Transformer (fall-through is equivalent) |
| `Resume` / `Resume label` | `goto label;` or `// TODO` comment |

#### Operator mapping

| VB6 | C# |
|---|---|
| `And` | `&&` |
| `Or` | `\|\|` |
| `Not` | `!` |
| `Xor` | `^` |
| `&` (concat) | `+` |
| `\` (integer div) | `((int)(left) / (int)(right))` |
| `^` (exponent) | `Math.Pow(left, right)` |
| `Is` | `==` (reference equality) |
| `Eqv` | `==` (approx) |
| `Like` | `==` with `/* Like — use Regex */` comment |
| `Imp` | `\|\|` with `/* Imp */` comment |
| `Mod` | `%` |

#### Bang access

`obj!Key` → `obj["Key"] /* ! */` — maps to indexed access; emits a comment because this is often a `Scripting.Dictionary` or ADO `Recordset` default member and may need manual review.

#### With blocks

`With expr` → `{ // With expr` block; `.Member` inside → `expr.Member` using a string stack (`_withStack`).

#### Built-in function mapping

All VB6 built-in functions and constants are mapped in `BuiltInMap.cs`. Key translations:

| VB6 | C# |
|---|---|
| `Len(s)` | `s.Length` |
| `Left(s,n)` / `Right(s,n)` | `s.Substring(...)` |
| `Mid(s,start[,len])` | `s.Substring(start-1[,len])` |
| `InStr([start,]s,find)` | `(s.IndexOf(find) + 1)` |
| `Replace(s,old,new)` | `s.Replace(old, new)` |
| `Split(s,delim)` | `s.Split(delim)` |
| `Join(arr,delim)` | `string.Join(delim, arr)` |
| `UCase/LCase` | `.ToUpper()/.ToLower()` |
| `Trim/LTrim/RTrim` | `.Trim()/.TrimStart()/.TrimEnd()` |
| `Chr(n)` | `((char)n).ToString()` |
| `Asc(s)` | `(int)s[0]` |
| `Hex(n)` | `n.ToString("X")` |
| `Format(v,fmt)` | `v.ToString(fmt) /* Format */` |
| `CStr(x)` | `x.ToString()` |
| `CInt/CLng(x)` | `Convert.ToInt32(x)` |
| `CDbl(x)` | `Convert.ToDouble(x)` |
| `CDate(x)` / `CVDate(x)` | `Convert.ToDateTime(x)` |
| `CBool(x)` | `Convert.ToBoolean(x)` |
| `Val(x)` | `Convert.ToDouble(x)` |
| `IIf(cond,t,f)` | `(cond ? t : f)` |
| `Nz(v[,def])` | `(v ?? def)` |
| `IsNull/IsEmpty(x)` | `(x == null)` |
| `IsNumeric(x)` | `double.TryParse(x.ToString(), out _)` |
| `Now` / `Date` | `DateTime.Now` / `DateTime.Today` |
| `DateSerial(y,m,d)` | `new DateTime(y, m, d)` |
| `Year/Month/Day/Hour/Minute/Second(d)` | `d.Year` / `.Month` / etc. |
| `UBound(a[,dim])` | `a.GetUpperBound(dim-1)` or `a.Length - 1` |
| `LBound(a)` | `0` |
| `Array(...)` | `new object[] { ... }` |
| `Debug.Print(x)` | `System.Diagnostics.Debug.WriteLine(x)` |
| `DateAdd/DateDiff/DatePart` | `/* TODO */` comment |

VB6 constants:

| VB6 | C# |
|---|---|
| `vbCrLf` | `"\r\n"` |
| `vbCr` / `vbLf` / `vbTab` / `vbNullChar` | `"\r"` / `"\n"` / `"\t"` / `'\0'` |
| `vbNullString` | `null` |
| `True` / `False` | `true` / `false` |
| `Nothing` / `Empty` / `Null` | `null` |
| `vbObjectError` | `unchecked((int)0x80040000)` |

#### `using` directives

Collected lazily during generation into `_requiredUsings` and prepended to the file:
- `System.Collections.ObjectModel` — added whenever any `Collection<…>` type is emitted.

**Integration test result (D46O003_1080.vbp — 137 files):**
- 137 files: Stage 5 complete, all `.cs` files written ✅

### Stage 5b: CsprojWriter ✅
After all files are generated, emit the SDK-style `.csproj`. Runs once per project.
Implemented in `VB6toCS.Core/Projects/CsprojWriter.cs`.

### Stage 5c: COM Type Library Reader (`.tlb` parser)

> **Status: planned**

#### Motivation

The current `ComTypeMap.cs` is a hand-maintained static table that covers the handful of COM libraries present in the integration test target (`ADODB`, `DAO`, `MSXML`, `Scripting`). It solves two known problems:

1. **Index vs. call disambiguation** — `pRs.Fields(i)` is a VB6 indexed property access, not a function call, so it must emit `pRs.Fields[i]` in C#.
2. **Parameter mode resolution** — cross-module `ref` annotations require knowing which COM parameters are `ByRef`.

Any COM library not in the table falls through to function-call syntax and loses `ref` qualifiers. The fix is to read the actual COM type library (`.tlb`) file, which is the authoritative metadata source for every COM interface — the same data that `tlbimp.exe` uses to generate interop assemblies.

#### What a `.tlb` file contains

A `.tlb` (OLE Automation Type Library) is a binary file described by the Microsoft OLE specification. Its `ITypeLib`/`ITypeInfo` COM interfaces expose:

| Metadata | Used for |
|---|---|
| Interface and coclass names | Type name resolution |
| Method names + parameter names | Cross-module call disambiguation |
| Parameter types (`VT_*` variant types) | `ByVal`/`ByRef` emission, return-type inference |
| `[out]`, `[in,out]` PARAMFLAG bits | `ref`/`out` keyword selection |
| Property flags (`INVOKE_PROPERTYGET`, `INVOKE_PROPERTYPUT`) | Getter vs. setter distinction |
| `FUNCFLAG_FRESTRICTED` / `VARFLAG_FHIDDEN` | Skip internal implementation details |
| `[defaultvalue]` / `[optional]` | Optional parameter handling |
| Default members (`FUNCFLAG_FDEFAULTBIND` or `DISPID_VALUE = 0`) | Bang-access and default-member expansion |
| Indexer markers (`INVOKE_FUNC` on a property returning a collection) | `[]` vs `()` disambiguation |

The `.vbp` `Reference=` line already gives us the `.tlb` path:
```
Reference=*\G{GUID}#2.0#0#C:\Program Files\Common Files\System\ado\msado15.dll#Microsoft ActiveX Data Objects 2.0 Library
```

The path (`msado15.dll`) may be a `.dll` with an embedded type library resource or a standalone `.tlb` / `.olb` file.

#### Approach

**Phase 1 — Windows COM API reader (primary path)**

Use `LoadTypeLib` / `ITypeLib` / `ITypeInfo` COM interfaces via P/Invoke or a thin C# COM-interop wrapper. This is the simplest correct implementation because the OS already knows how to parse every `.tlb` format variant:

```csharp
// VB6toCS.Core/TypeLibraries/TlbReader.cs
public sealed class TlbReader
{
    // Loads a .tlb or embedded type library from a .dll path.
    // Returns null if the path is inaccessible (non-Windows, file not found).
    public static ComTypeLibrary? TryLoad(string path) { ... }
}
```

Returns a `ComTypeLibrary` record with:
```csharp
public sealed record ComTypeLibrary(
    string LibraryName,
    IReadOnlyDictionary<string, ComTypeInfo> Types);    // typeName (ci) → type info

public sealed record ComTypeInfo(
    string Name,
    ComTypeKind Kind,                                   // Interface, Dispatch, CoClass, Enum, …
    IReadOnlyDictionary<string, ComMemberInfo> Members);

public sealed record ComMemberInfo(
    string Name,
    ComMemberKind Kind,                                 // Method, PropertyGet, PropertyPut, PropertyPutRef
    bool IsDefaultMember,
    bool IsIndexed,                                     // DISPID_VALUE with parameters → indexer
    IReadOnlyList<ComParamInfo> Parameters,
    string? ReturnTypeName);                            // mapped to C# type name, null if unresolvable

public sealed record ComParamInfo(
    string Name,
    string TypeName,                                    // C# type name
    bool IsOut,
    bool IsRef);
```

**Phase 2 — Cache / descriptor files (portability)**

To support running the translator on Linux (e.g. in CI), cache the parsed metadata as a JSON descriptor file alongside the `.tlb`:

```
msado15.tlb.vb6tocs.json
```

If the JSON cache exists and is newer than the `.tlb`, use it directly. This means the type library only needs to be read once on a Windows machine; the cache travels with the project.

**Phase 3 — Binary `.tlb` parser (optional, cross-platform)**

A pure-C# binary parser for the SLTG / MSFT `.tlb` formats (both documented in the Open Specifications). This eliminates the Windows dependency entirely. More complex to implement; defer until the COM API + cache approach proves insufficient.

#### Integration with the pipeline

The `.tlb` reader runs **between Stage 0 (VbpReader) and Stage 2 (Parser)**, before any AST work:

```
VbpReader → TlbReader (one per Reference= entry) → build ComTypeRegistry → Stages 1–5
```

The `ComTypeRegistry` (keyed by type name, case-insensitive) is passed to:

1. **`ComTypeMap`** (Stage 5, CodeGenerator) — replaces the static `IndexedMembers` set with a live lookup. `IsIndexedMember(ownerType, memberName)` queries the registry first; falls back to the static table if the type is not in the registry (e.g. COM reference path not found).

2. **`BuildMethodParamMap`** (Program.cs, between Stages 4 and 5) — currently only covers project-internal methods. Extended to also include COM methods from the registry, so cross-COM `ref` parameters are emitted correctly.

3. **`TryGetSimpleType`** (CodeGenerator) — currently resolves only declared-type identifiers. With COM metadata, `pRs.Fields[i]` can be resolved to `ADODB.Field`, enabling further chained type lookups.

4. **`CollectionTypeInferrer`** — currently only uses declared VB6 types for `.Add()` site inference. With COM metadata, `rs.Fields[i]` can be typed as `ADODB.Field`, enabling the inferrer to resolve more element types.

#### Known limitations

- **Path resolution**: the path in `Reference=` is the path on the machine where the VB6 project was last built. It is often a Windows absolute path (`C:\Program Files\...`) that does not exist on the translation machine. The tool must gracefully fall back to the static `ComTypeMap` when the `.tlb` path is not accessible.
- **Variant types**: many COM methods use `VARIANT` (`VT_VARIANT`) for parameters, which maps to `object` — no better than the current situation. Inference still applies.
- **Late binding**: VB6 code using `Dim x As Object` and late-binding COM calls cannot be typed statically regardless of `.tlb` metadata.
- **COM redirection**: some COM libraries register type libraries in the Windows Registry separately from the path in `.vbp`. A registry fallback lookup (`HKEY_CLASSES_ROOT\TypeLib\{GUID}\version`) may be needed.

#### Build order position

This is an enhancement to the existing pipeline, not a new stage. It slots between Stage 0 and Stage 1 in the build order and improves output quality at Stage 5. It should be implemented after Stage 6 (Roslyn formatting) since it has no dependencies on later stages.

---

### Stage 6: Post-Processing / Formatting
Run output through Roslyn's formatter for clean, idiomatic C#. Roslyn's analyzer flags anything that needs manual review.

---

## Technology Choices

- **Language**: C# / .NET 8
- **Parser**: Hand-written recursive descent — no external tool or Java dependency
- **Output formatting**: Roslyn (Microsoft.CodeAnalysis)
- **Architecture**: Pipeline of passes over the AST, each pass handling one concern — keeps it testable and extensible

---

## Build Order

1. ~~Lexer (tokenizer)~~ ✅
2. ~~AST node model~~ ✅
3. ~~Recursive descent parser~~ ✅
4. ~~VbpReader — `.vbp` project file parser~~ ✅
5. ~~Parser validated against real-world 137-file project~~ ✅
6. ~~Semantic analysis (symbol table, CallOrIndex resolution, property grouping)~~ ✅
7. ~~IR transformation (type normalization + error handling restructuring)~~ ✅
8. ~~C# code generator (AST → C# source text)~~ ✅
9. ~~CsprojWriter — generate `.csproj`~~ ✅
10. Roslyn formatter + "needs manual review" annotation system
11. COM `.tlb` type library reader — dynamic `ComTypeMap` + `ref`/`out` resolution for arbitrary COM references

---

## Realistic Expectations

For clean ActiveX DLL / standard module code, **80–90% automation** is achievable. The remainder will be flagged for human review. 100% automatic translation is not a goal.

---

## Integration Test Target

`samples/syas/D46O003_1080.vbp` — a real-world production ActiveX DLL from a large legacy codebase.

- **136 classes** (`.cls`), **1 standard module** (`.bas`), **0 forms** — 100% in scope
- **6 COM references**: OLE Automation, ADO 2.0, Scripting Runtime, DAO 3.51, MSSOAP 3.0, MSXML 3.0
- The surrounding `syas/` directory contains 16 VB6 projects and 562 files total; only `D46O003_1080.vbp` is the target

**Parser milestone:** All 137 files parse successfully (Stage 2 complete). The project produces 137 placeholder `.cs` files and a `.csproj` today. A fully successful run — producing a compilable C# project with no hard failures, only review comments — is the definition of "done" for the full pipeline.
