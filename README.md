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
- External COM references — handled via a three-tier strategy (see [ROADMAP.md](ROADMAP.md))

**Out of scope:**
- `.frm` forms, `.ctl` UserControls — skipped with a warning
- File I/O statements (`Open`, `Close`, `Print #`, etc.) — skipped with a `// TODO` comment
- ADO/DAO data access objects

## Usage

```bash
# Translate a full VB6 project (primary mode)
dotnet run --project src/VB6toCS.Cli -- MyProject.vbp

# Parse a single file and print its AST tree (diagnostic mode)
dotnet run --project src/VB6toCS.Cli -- Calculator.cls
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

- `Option Explicit` is **required** in every source file. Files without it are rejected with a clear error — add it to the top of the file and re-run.

## Current status

The pipeline is being built stage by stage:

| Stage | Description | Status |
|---|---|---|
| 0 | VBP project reader + `.csproj` writer | ✅ Complete |
| 1 | Lexer (tokenizer) | ✅ Complete |
| 2 | Parser → AST | ✅ Complete |
| 3 | Semantic analysis (symbol table, CallOrIndex resolution, property grouping) | ✅ Complete |
| 4 | IR transformation (type normalization + error handling restructuring) | ✅ Complete |
| 5 | C# code generation | Next |
| 6 | Roslyn formatting + review annotations | Planned |

**Integration test:** `D46O003_1080.vbp` — a real-world production ActiveX DLL (136 classes, 1 module, 0 forms, 6 COM references). All 137 files pass through Stage 4 with 0 errors.

## Architecture

```
VB6toCS.sln
src/
  VB6toCS.Core/       ← translator engine (pure .NET 8, no external parser deps)
    Lexing/           ← hand-written VB6 lexer
    Parsing/          ← recursive descent parser, AST node types, AstPrinter
    Analysis/         ← semantic analysis (Stage 3)
    Transformation/   ← IR transformation (Stage 4)
    Projects/         ← VbpReader, CsprojWriter
  VB6toCS.Cli/        ← command-line entry point
samples/
  Calculator.vbp/.cls ← sample project for quick testing
```

