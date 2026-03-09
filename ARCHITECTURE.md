# OpenVB6toCS — Translation Pipeline Architecture

This document describes the full translation pipeline from VB6 source to C# output.

```mermaid
flowchart TD
    INPUT([".vbp / .cls / .bas<br/>input file"]) --> DISPATCH

    DISPATCH{Input type?}
    DISPATCH -->|".vbp project"| S0
    DISPATCH -->|".cls / .bas single file"| S1

    %% ── Stage 0: Project Reader ───────────────────────────────────────────────
    subgraph S0["Stage 0 — VBP Project Reader"]
        VBP_LINE{Line type?}
        VBP_LINE -->|"Class= / Module="| ADD_SRC[Add to source file list]
        VBP_LINE -->|"Form= / UserControl="| SKIP_UI["Skip — warn user<br/>UI out of scope"]
        VBP_LINE -->|"Reference= COM lib"| COM_TIER
        VBP_LINE -->|"Name= / ExeName32="| PROJ_NAME[Record project name]

        COM_TIER{COM mapping tier}
        COM_TIER -->|"Known .NET equiv<br/>e.g. Scripting.Dictionary"| COM_MAP["Translate silently<br/>e.g. → Dictionary&lt;K,V&gt;"]
        COM_TIER -->|"Unknown COM type"| COM_INTEROP["COM Interop fallback<br/>add &lt;COMReference&gt; to .csproj<br/>annotate usage with INTEROP:"]
    end

    ADD_SRC --> S1

    %% ── Stage 1: Lexer ────────────────────────────────────────────────────────
    subgraph S1["Stage 1 — Lexer"]
        LEX["Tokenise source text<br/>hand-written scanner"]
        LEX --> OPT_EXP{"Option Explicit<br/>present?"}
        OPT_EXP -->|No| REJECT["Reject file<br/>with clear error"]
        OPT_EXP -->|Yes| TOK_LIST[Token list]

        LEX_NOTES["Special handling:<br/>• Line continuation _ → silent<br/>• Type suffixes %&amp;!#@$ → consumed, not appended<br/>• &amp;H1F hex → 0x1F<br/>• #date# → DateLiteralToken<br/>• ! followed by letter → BangToken<br/>• Comments → preserved as tokens"]
    end

    TOK_LIST --> S2

    %% ── Stage 2: Parser ───────────────────────────────────────────────────────
    subgraph S2["Stage 2 — Recursive Descent Parser"]
        PARSE["Parse token stream<br/>into typed AST"]
        PARSE --> AST["ModuleNode root<br/>59 AST node types"]

        PARSE_NODES["Node categories:<br/>• Declarations: Field, Const, Enum, UDT, Implements<br/>• Procedures: Sub, Function, Property Get/Let/Set<br/>• Statements: If, Select, For, ForEach, While, Do,<br/>  With, OnError, GoTo, Label, Dim, Assignment, Call<br/>• Expressions: BinaryExpr, UnaryExpr, MemberAccess,<br/>  BangAccess, CallOrIndex (ambiguous), NewObject, …<br/>• Literals: Integer, Double, String, Bool, Date, Nothing"]
    end

    AST --> S3

    %% ── Stage 3: Semantic Analysis ────────────────────────────────────────────
    subgraph S3["Stage 3 — Semantic Analysis (3 passes)"]
        P3_1["Pass 1 — Build module symbol table<br/>All declarations → SymbolTable<br/>One entry per field / method / property / enum / UDT"]
        P3_2["Pass 2 — Transform bodies<br/>• CallOrIndexNode → IndexNode  (known arrays)<br/>• FunctionName = value → FunctionReturnNode<br/>• Identifier casing normalised to declaration site"]
        P3_3["Pass 3 — Property grouping<br/>Property Get + Let + Set triads<br/>→ single CsPropertyNode"]
        P3_1 --> P3_2 --> P3_3

        COLL_INF["CollectionTypeInferrer<br/>(cross-module, runs after all Pass 1s)<br/>Infers Collection element type from .Add call sites<br/>→ collectionTypeMap: (module,field) → elemType"]
    end

    P3_3 --> MAPS
    COLL_INF --> MAPS

    %% ── Cross-module maps ─────────────────────────────────────────────────────
    subgraph MAPS["Cross-module Map Building (Program.cs)"]
        MAP1["collectionTypeMap<br/>(module,field) → elemType"]
        MAP2["enumMemberMap<br/>memberName → (module,enumName)"]
        MAP3["enumTypedFieldNames<br/>field names whose type is an enum"]
        MAP4["enumTypeMap<br/>enumTypeName → moduleName"]
        MAP5["methodParamMap<br/>module → method → ParameterMode list"]
        MAP6["globalVarMap<br/>fieldName → typeName  (public fields)"]
    end

    MAPS --> S4

    %% ── Stage 4: IR Transformation ────────────────────────────────────────────
    subgraph S4["Stage 4 — IR Transformation (2 passes)"]
        T4_1["Pass 1 — Type normalisation<br/>String→string  Long→int  Boolean→bool<br/>Date→DateTime  Currency→decimal  Variant→object<br/>VB6 intrinsic enums (VbMsgBoxStyle etc.) → int<br/>Applied to ALL TypeRefNode instances"]

        T4_2["Pass 2 — Error handling restructuring"]
        T4_1 --> T4_2

        EH_DETECT{"Structured pattern<br/>detected?"}
        T4_2 --> EH_DETECT

        EH_DETECT -->|"Yes — canonical<br/>On Error GoTo / cleanup label / error label"| EH_CONVERT

        subgraph EH_CONVERT["Structured conversion"]
            TRY_BODY["try body = statements before cleanup label"]
            CATCH_BODY["catch body = error handler block"]
            FINALLY_BODY["finally body = cleanup label block"]
            GOTO_THROW["GoTo errLabel inside try body<br/>→ ThrowNode  (throw new Exception)"]
            TRY_BODY --- CATCH_BODY --- FINALLY_BODY --- GOTO_THROW
        end

        EH_DETECT -->|"No — multiple handlers /<br/>Resume / complex flow"| EH_PRESERVE["Preserve as structured comments<br/>for manual developer review"]
    end

    EH_CONVERT --> S5
    EH_PRESERVE --> S5
    T4_1 --> S5

    %% ── Stage 5: Code Generation ──────────────────────────────────────────────
    subgraph S5["Stage 5 — C# Code Generation"]

        subgraph MEMBER_GEN["Module member emission"]
            MG_FIELD["FieldNode → public/private field<br/>with C# type"]
            MG_ENUM["EnumNode → public enum Name : int<br/>Member qualif.: same module → EnumName.M<br/>cross-module → ModuleName.EnumName.M"]
            MG_UDT["UdtNode → public struct"]
            MG_SUB["SubNode → void method<br/>(.bas module → static)"]
            MG_FUNC["FunctionNode → typed method<br/>_result local + return _result pattern<br/>Return type inferred when declared As Variant<br/>and all return expressions share one type"]
            MG_PROP["CsPropertyNode → { get; set; } property<br/>or split get/set if asymmetric"]
            MG_DECL["Declare Sub/Function → [DllImport] extern"]
        end

        subgraph STMT_GEN["Statement branches"]
            ST_TRYCATCH["TryCatchNode → try/catch/finally"]
            ST_THROW["ThrowNode → throw new Exception()"]
            ST_ONERR["OnErrorNode → // VB6: On Error ... comment"]
            ST_FOREACH["ForEachNode → foreach<br/>(Dictionary → .Values appended)"]
            ST_WITH["WithNode → using _withStack<br/>(nested With supported via Stack&lt;string&gt;)"]
            ST_SELECT["SelectCaseNode → switch<br/>CaseIs → relational case<br/>CaseRange → case x when val >= lo &amp;&amp; val &lt;= hi"]
            ST_ASSIGN["AssignmentNode<br/>Set obj = ... → plain assignment<br/>Let primitive = ... → plain assignment"]
            ST_REDIM["ReDimNode → arr = new T[n]"]
        end

        subgraph EXPR_GEN["Expression branches"]
            EX_BANG["BangAccessNode (obj!Key)<br/>→ resolve obj type → look up default member<br/>→ obj.Fields(@'Key') for Recordset<br/>or obj[@'Key'] fallback"]
            EX_FIELDS["Fields() / BangAccess in typed context<br/>→ Convert.ToDouble() / Convert.ToInt32() / etc.<br/>based on LHS type"]
            EX_CALL["CallOrIndexNode<br/>• COM indexed member → obj[args]<br/>• ByRef params → ref prepended<br/>• Optional trailing args → omitted<br/>• Optional middle args → named + default"]
            EX_ENUM["Enum member identifier<br/>→ qualified via enumMemberMap"]
            EX_TYPE["TypeStr()<br/>→ cross-module enum type → ModuleName.Type<br/>→ VB6 Collection → Collection&lt;T&gt; or Dictionary&lt;string,T&gt;<br/>→ built-in type map (Long→int, etc.)"]
            EX_BUILTIN["Built-in functions<br/>CStr→.ToString()  CInt→Convert.ToInt32()<br/>Left/Right/Mid/InStr → string methods<br/>IIf→ternary  Nz→??  Len→.Length"]
        end

    end

    S5 --> S6

    %% ── Stage 6 / Output ──────────────────────────────────────────────────────
    subgraph S6["Output"]
        CS_FILES[".cs files<br/>one per source module"]
        CSPROJ[".csproj<br/>SDK-style, net8.0, Library<br/>+ &lt;COMReference&gt; for unmapped COM libs"]
        ROSLYN["Stage 6 — Roslyn formatting<br/>(planned — normalise whitespace and style)"]
        CS_FILES --> ROSLYN
    end

    CSPROJ --> DONE([Build-ready C# project])
    ROSLYN --> DONE
```

## Pipeline stages summary

| Stage | Component | Description |
|---|---|---|
| 0 | `VbpReader` + `CsprojWriter` | Reads `.vbp` project, classifies source files, writes SDK `.csproj` |
| 1 | `Lexer` | Hand-written VB6 tokeniser; enforces `Option Explicit` |
| 2 | `Parser` | Recursive descent, produces 59-node typed AST |
| 3 | `Analyser` + `CollectionTypeInferrer` | Symbol table, type resolution, property grouping, Collection element inference |
| — | Map builders (`Program.cs`) | Six cross-module maps linking all 3+ modules |
| 4 | `Transformer` | Type normalisation + structured error-handling conversion |
| 5 | `CodeGenerator` | AST → C# source text with all cross-module resolution |
| 6 | Roslyn formatter | *(Planned)* Normalise whitespace and style |

## Key design decisions

- **No external parser library.** Pure hand-written C# recursive descent — no ANTLR, no YAML, no Java toolchain.
- **Records for all AST nodes.** `abstract record AstNode` gives free value equality and clean `with`-expression transforms.
- **`CallOrIndexNode` ambiguity resolved in Stage 3.** `foo(x)` is lexically ambiguous between a call and an array index; the symbol table resolves it.
- **Error handling pattern detection.** Seven structural checks validate the canonical `On Error GoTo / cleanup label / error label` shape before converting. Patterns that fail any check are preserved as comments.
- **`GoTo ErrorHandler` → `throw new Exception()`.** Recursive replacement of `GoTo errLabel` nodes inside the try body; handles the common VB6 validation pattern (`If condition Then GoTo ErrorHandler`).
- **Cross-module maps.** Six maps built once after Stage 3 and passed into Stage 5 to resolve enum members, parameter modes, collection types, and field types across all modules without a global type registry.
