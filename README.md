# OpenVB6toCS — VB6 to C# Translator

An open-source, free Visual Basic 6 to C# source code translator targeting **ActiveX DLLs** — business logic components with no UI.

## Scope

**Supported input:**
- `.vbp` project files (primary mode) — translates all classes and modules in one go
- `.cls` class modules → `public class`
- `.bas` standard modules → `public static class`

**In scope:**
- Fields, constants, enums, user-defined types (`Type`)
- Methods (`Sub`, `Function`), properties (`Property Get/Let/Set`)
- Control flow: `If/ElseIf/Else`, `Select Case`, `For/Next`, `For Each`, `While`, `Do/Loop`, `With`
- Error handling: `On Error GoTo`, `On Error Resume Next`, `Resume`, `Err.Raise`
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
| 4 | IR transformation (type normalization + error handling restructuring) | ✅ Complete |
| 5 | C# code generation | ✅ Complete |
| 6 | Roslyn output formatting | Planned |

**Integration test:** `D46O003_1080.vbp` — a real-world production ActiveX DLL (136 classes, 1 module, 0 forms, 6 COM references). All 137 files translate end-to-end to `.cs` files.

## What the generated C# looks like

VB6 input (`Calculator.cls`):
```vb
Public Function Add(ByVal a As Double, ByVal b As Double) As Double
    Add = a + b
End Function
```

Generated C# output (`Calculator.cs`):
```csharp
public class Calculator
{
    public double Add(double a, double b)
    {
        double _result = default;
        _result = a + b;
        return _result;
    }
}
```

## Architecture

```
VB6toCS.sln
src/
  VB6toCS.Core/           ← translator engine (pure .NET 8, no external parser deps)
    Lexing/               ← hand-written VB6 lexer
    Parsing/              ← recursive descent parser, AST node types, AstPrinter
    Analysis/             ← semantic analysis: symbol table, type resolution, Collection<T> inference
    Transformation/       ← IR transformation: type normalization, error handling restructuring
    CodeGeneration/       ← C# code emitter, built-in function map, COM type map
    Projects/             ← VbpReader, CsprojWriter
  VB6toCS.Cli/            ← command-line entry point, pipeline orchestration
samples/
  Calculator.vbp/.cls     ← minimal sample project for quick testing
  syas/                   ← real-world legacy VB6 codebase (16 projects, 562 files)
```

## Known limitations

- **COM type accuracy:** indexed COM members (e.g. `ADODB.Fields(i)`) are identified via a static lookup table (`ComTypeMap.cs`) covering ADODB, DAO, and MSXML. COM types not in the table fall back to method-call syntax and may need manual adjustment. A `.tlb` COM type library reader is planned to eliminate this limitation.
- **`ref`/`out` parameters on COM methods:** resolved statically for known types; unknown COM method parameters default to `ref`.
- **`Collection<T>` element type:** inferred from `.Add` call sites within the same project. Cross-module or late-bound element types fall back to `Collection<object>` with a `// REVIEW:` comment.
- **On Error Resume Next:** translated with a `// VB6: On Error Resume Next` comment; errors are not suppressed in the C# output and must be reviewed manually.
- **GoSub/Return:** translated as labeled blocks with comments; not idiomatic C#.
