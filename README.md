# OpenVB6toCS — VB6 to C# Translator

An open-source, free Visual Basic 6 to C# source code translator targeting **ActiveX DLLs** — business logic components with no UI.

![logo](https://user-images.githubusercontent.com/17026744/79247080-68442000-7e50-11ea-9137-faeec5209107.png)

## Scope

**Supported input:**
- `.vbp` project files (primary mode) — translates all classes and modules in one go
- `.cls` class modules → `public class`
- `.bas` standard modules → `public static class`

**In scope:**
- Fields, constants, enums, user-defined types (`Type`)
- Methods (`Sub`, `Function`), properties (`Property Get/Let/Set`)
- Control flow: `If/ElseIf/Else`, `Select Case`, `For/Next`, `For Each`, `While`, `Do/Loop`, `With`
- Error handling: structured `On Error GoTo` patterns → `try/catch/finally`; complex patterns preserved as comments for manual review
- COM interfaces: `Implements`
- Parameters: `ByVal`, `ByRef`, `Optional`, `ParamArray`
- Win32 API declarations (`Declare Sub/Function`)
- External COM references — three-tier strategy: known mappings translated automatically, unknown types compiled via COM Interop with `<COMReference>` in the `.csproj`, and usage annotated for review

**Out of scope:**
- `.frm` forms, `.ctl` UserControls — skipped with a warning
- File I/O statements (`Open`, `Close`, `Print #`, etc.) — skipped with a `// TODO` comment

## Usage

```bash
# Translate a full VB6 project (primary mode)
dotnet run --project src/VB6toCS.Cli -- MyProject.vbp

# Translate up to a specific stage (0=list files, 1=lex, 2=parse, 3=analyse, 4=transform, 5=generate)
dotnet run --project src/VB6toCS.Cli -- MyProject.vbp -UpToStage 3

# Parse a single file and print its AST tree (diagnostic mode)
dotnet run --project src/VB6toCS.Cli -- Calculator.cls

# Specify source file encoding (default: windows-1252)
dotnet run --project src/VB6toCS.Cli -- MyProject.vbp -Encoding utf-8
```

### Project mode output

```
MyProject/
  ClassName.cs        ← one .cs per source file
  ModuleName.cs
  ...
  MyProject.csproj    ← SDK-style, targets net8.0, ready to build
```

COM libraries with no known .NET mapping get a `<COMReference>` entry in the `.csproj` so the project compiles immediately via interop.

### Requirements

- .NET 8 SDK
- `Option Explicit` is **required** in every source file. Files without it are rejected with a clear error — add it to the top of the file and re-run.

## Current status

| Stage | Description | Status |
|---|---|---|
| 0 | VBP project reader + `.csproj` writer | ✅ Complete |
| 1 | Lexer (tokenizer) | ✅ Complete |
| 2 | Parser → AST | ✅ Complete |
| 3 | Semantic analysis (symbol table, type resolution, property grouping) | ✅ Complete |
| 4 | IR transformation (type normalization + structured error handling) | ✅ Complete |
| 5 | C# code generation | ✅ Complete |
| 6 | Roslyn output formatting | Planned |

**Integration test:** `D46O003_1080.vbp` — a real-world production ActiveX DLL (136 classes, 1 module, 0 forms, 6 COM references). All 137 files translate end-to-end to `.cs` files. The generated project builds with **zero genuine translator errors**; all remaining build failures are due to absent COM/cross-project assembly references, which are expected in a standalone build.

## What the generated C# looks like

### Basic function

VB6 input:
```vb
Public Function Add(ByVal a As Double, ByVal b As Double) As Double
    Add = a + b
End Function
```

Generated C#:
```csharp
public double Add(double a, double b)
{
    double _result = default;
    _result = a + b;
    return _result;
}
```

### Error handling — structured `try/catch/finally`

The translator detects the canonical VB6 error handling pattern and converts it to clean, idiomatic C#. When the pattern is detected, no residual `GoTo`, labels, or `// VB6:` comments appear in the output.

VB6 input:
```vb
Function LoadData(ByVal sTable As String) As Boolean
    On Error GoTo ErrorHandler
    OpenConnection
    If Not FetchData(sTable) Then GoTo ErrorHandler
    LogSuccess
ExitGracefully:
    CloseConnection
    Exit Function
ErrorHandler:
    LogError
    GoTo ExitGracefully
End Function
```

Generated C#:
```csharp
public bool LoadData(string sTable)
{
    bool _result = default;
    try
    {
        OpenConnection();
        if (!FetchData(sTable)) throw new Exception();
        LogSuccess();
    }
    catch (Exception _ex)
    {
        LogError();
    }
    finally
    {
        CloseConnection();
    }
    return _result;
}
```

`GoTo ErrorHandler` inside the try body — the VB6 way of signalling a failure — is automatically converted to `throw new Exception()`. When the pattern cannot be safely auto-converted (multiple handlers, `Resume`, complex flow), all error-related nodes are preserved as structured comments so the developer can see the original intent and restructure manually.

### Collections

VB6 `Collection` objects are analysed across the entire project. The translator infers whether each collection is keyed (→ `Dictionary<string, T>`) or unkeyed (→ `List<T>`), and resolves the element type from `.Add` call sites.

VB6 input:
```vb
Dim colItems As New Collection
colItems.Add New clsOrder, "K001"

Dim item As clsOrder
For Each item In colItems
    item.Process
Next
```

Generated C#:
```csharp
Dictionary<string, clsOrder> colItems = new();

clsOrder item;
foreach (var item in colItems.Values)
{
    item.Process();
}
```

### ADO Recordset field access

VB6 uses `Recordset.Fields("column")` (which returns `object`) in typed contexts. The translator wraps these calls in the appropriate `Convert.To*()` call based on the target type.

VB6 input:
```vb
Dim total As Double
total = rs.Fields("Amount") + rs.Fields("Tax")

If rs!Status = 1 Then   ' bang-access shortcut for Fields()
    ...
End If
```

Generated C#:
```csharp
double total = default;
total = Convert.ToDouble(rs.Fields(@"Amount")) + Convert.ToDouble(rs.Fields(@"Tax"));

if (Convert.ToInt32(rs.Fields(@"Status")) == 1)
{
    ...
}
```

Bang-access (`!`) is expanded to the default member of the object's type (`Fields` for recordsets), and the result is wrapped with the correct `Convert.To*()` based on context.

### Enum qualification across modules

VB6 enums are global by default. In C#, they are nested types of their declaring class. The translator automatically qualifies cross-module enum type references.

VB6 input (in `clsOrder.cls`):
```vb
Public Cod_Ramo As e_CodRamo   ' e_CodRamo defined in stcConfig.cls
```

Generated C#:
```csharp
public stcConfig.e_CodRamo Cod_Ramo;
```

## Architecture

See [ARCHITECTURE.md](ARCHITECTURE.md) for a full Mermaid flowchart of the translation pipeline, including all decision branches for error handling, COM references, expression types, and code generation cases.

```
VB6toCS.sln
src/
  VB6toCS.Core/           ← translator engine (pure .NET 8, no external parser deps)
    Lexing/               ← hand-written VB6 lexer
    Parsing/              ← recursive descent parser, AST node types, AstPrinter
    Analysis/             ← semantic analysis: symbol table, type resolution, Collection<T> inference
    Transformation/       ← IR transformation: type normalization, structured error handling
    CodeGeneration/       ← C# code emitter, built-in function map, COM type map
    Projects/             ← VbpReader, CsprojWriter
  VB6toCS.Cli/            ← command-line entry point, pipeline orchestration
samples/
  Calculator.vbp/.cls     ← minimal sample project for quick testing
  syas/                   ← real-world legacy VB6 codebase (16 projects, 562 files)
```

## Translation features

- **Type mapping:** `Long` → `int`, `String` → `string`, `Boolean` → `bool`, `Date` → `DateTime`, `Currency` → `decimal`, `Variant` → `object`, `Object` → `object`
- **String literals:** emitted as C# verbatim strings (`@"..."`) — VB6 `""` quoting and backslashes work correctly with no transformation
- **Properties:** VB6 `Property Get/Let/Set` triples are grouped into a single C# `{ get; set; }` property
- **Function returns:** VB6 assigns the return value by name (`FunctionName = value`); translated to a `_result` local with `return _result`
- **ByRef parameters:** `ref` keyword emitted at call sites for all `ByRef` parameters, resolved cross-module
- **Optional parameters:** trailing optional args omitted at call sites; middle missing args use named-argument form
- **Collections:** element type and kind (`Dictionary`/`List`) inferred from `.Add` call sites across the entire project
- **Enums:** members qualified at use sites (`EnumName.Member`); cross-module references qualified as `ModuleName.EnumName`; cross-module enum type names in declarations also qualified
- **ADO/COM field access:** `Fields("key")` and `!key` (bang-access) wrapped with `Convert.To*()` based on the target type context
- **Error handling:** canonical `On Error GoTo` / cleanup label / error label pattern → `try/catch/finally`; `GoTo ErrorHandler` inside try body → `throw new Exception()`; unrecognised patterns → structured comments
- **`With` blocks:** tracked on a stack; `.Member` expressions resolved to the correct `With` object
- **`Select Case`:** mapped to C# `switch`; VB6 `Is` comparisons and range patterns supported
- **`For Each` over collections:** `.Values` appended automatically when iterating a `Dictionary`

## Known limitations

- **COM type accuracy:** indexed COM members (e.g. `ADODB.Fields(i)`) are identified via a static lookup table (`ComTypeMap.cs`) covering ADODB, DAO, and MSXML. COM types not in the table fall back to method-call syntax and may need manual adjustment.
- **`ref`/`out` parameters on COM methods:** resolved statically for known types; unknown COM method parameters default to `ref`.
- **`Collection` element types:** inferred from `.Add` call sites within the same project. Cross-module or late-bound element types fall back to `Dictionary<string, object>` with a `// REVIEW:` comment.
- **VB6 `Static` locals:** do not persist between calls in the generated C# (VB6 semantics differ). Marked with `// VB6 Static local — hoist to field if persistence needed`.
- **Complex error handling:** `On Error GoTo` patterns that don't match the canonical cleanup-label structure (multiple handlers, `Resume`, retry semantics) are preserved as `// VB6:` comments for manual restructuring.
- **GoSub/Return:** translated as labeled blocks with comments; not idiomatic C#.
- **Optional parameters with mixed ordering:** VB6 allows optional parameters before required ones; C# does not. Such functions require manual parameter reordering.
- **Date literals** (`#1/15/2020#`) and `DateAdd`/`DateDiff`/`DatePart` calls emit `default /* TODO */` placeholders for manual completion.
