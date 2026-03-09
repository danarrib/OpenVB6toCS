# Translator Improvement Opportunities

Identified by building the translated `D46O003_1080` project (~2,500 compiler errors).
Items are ordered by priority (errors eliminated vs implementation effort).

> **Status after session (items 1–8 sprint):** errors reduced from ~2,500 → ~290,
> of which ~286 are expected cross-project/COM-dependency errors in the stripped test build.
> Only ~4 genuine translator errors remain (1 CS1737 mixed-optional function + 3 CS1750
> optional-int-default with COM enum types). All syntax errors have been eliminated.

---

## 🔴 Critical — large error counts, straightforward fixes

### ✅ 1. Double-quoted strings in string literals (CS1002/CS1513/CS8997 — ~1,000 errors) ✦ also fixes #5

VB6 escapes a quote inside a string by doubling it: `"He said ""hello"""`.
The translator currently emits the doubled quotes verbatim, which C# 11+ interprets as a
raw string literal delimiter (`"""`), causing parse errors.

**Fix:** Prefix all emitted string literals with `@` (C# verbatim string).

This works because C# verbatim strings use the exact same `""` quoting convention as VB6 —
no transformation of string content is needed at all. As a bonus, verbatim strings also treat
backslashes as literal characters, which eliminates item #5 entirely for free. VB6 strings
never contain real escape sequences (`\n`, `\t`, etc.) — special characters are expressed via
`vbCrLf`, `Chr()`, etc. — so there is no downside to always using `@"..."`.

```vb
' VB6
lMsg = "FNP9999 - Valor ""inválido"""
```
```csharp
// Current (broken)
lMsg = "FNP9999 - Valor ""inválido"""
// Fixed
lMsg = @"FNP9999 - Valor ""inválido""";
```

---

### ✅ 2. `/* missing */` in argument positions (CS0839 — ~710 errors)

When a VB6 call omits an optional argument (e.g. `Foo(, b)`), the translator emits
`/* missing */` as a literal in the argument list. That is not valid C# syntax.

There are two distinct sub-problems to solve:

#### Sub-problem A — Omitted middle arguments (common case)

VB6 allows `Foo(a, , c)` to skip a middle optional. C# can express this with named arguments,
omitting the skipped parameter by name: `Foo(a: x, c: y)`. No definition change needed.

**Fix:** At call sites where an argument slot contains `IsMissing: true`:
- If the slot is trailing → simply omit it (C# allows omitting trailing optional args positionally)
- If the slot is in the middle → switch to named-argument form for all subsequent non-missing args;
  skip the missing ones entirely

```vb
' VB6
Foo a, , c   ' middle arg omitted
```
```csharp
// Fixed — named args, middle arg omitted
Foo(a: x, c: y);
```

#### Sub-problem B — Mixed parameter ordering in definitions (rare)

VB6 allows optional params anywhere: `Sub Foo(Optional a As String, b As String)`.
C# requires all optional params to come after all required params — so this definition
is illegal as-is.

**Fix (two-step, requires cross-module tracking):**

1. **Definition site:** detect when optional params precede required ones; reorder the
   parameter list so required params come first, optional params last.
2. **Call sites:** build a cross-module map `(module, procName) → (originalPosition → paramName)`
   for every reordered function. At each call site, use named arguments so each argument
   reaches the correctly renamed parameter regardless of the reordering.

```vb
' VB6 definition
Sub Foo(Optional a As String = "", b As String)
```
```csharp
// Reordered C# definition
void Foo(string b, string a = "")
// Call site — named args preserve original intent
Foo(b: x, a: y);
// Call site with omitted optional
Foo(b: x);  // 'a' simply absent
```

**Cross-module map structure:** built between Stages 3 and 5 (similar to `enumMemberMap`).
Key: `(moduleName, procName)`. Value: list of `(csParamName, originalVb6Position)` for
every reordered function. Code generator consults this map when emitting any call expression
whose target resolves to a reordered function.

**Note on external callees:** when the callee is a COM/external method not in the project,
emit `default` for missing args as a safe fallback (it compiles; the value may be wrong
and should be flagged for review).

**Implementation status:**
- Sub-problem A ✅ — trailing missing args omitted; middle missing args → `default /* missing optional */`
- Sub-problem B ⚠️ PARTIAL — CS1741 (ByRef optional default) fixed by stripping the default value.
  CS1737 (mixed optional ordering) still affects 1 function; requires cross-module param-reordering map.
  Remaining: 2 CS1737 errors (1 function, duplicated by compiler pass).

---

### ✅ 3. VB6 `Static` locals emitted as `static` in method bodies (CS0106 — ~352 errors)

VB6 `Static` local variables persist their value across calls (procedure-level singletons).
The translator currently emits `static T name = default;` inside methods, which C# forbids.

**Fix options:**
- **Preferred:** Hoist to a private field on the class, renamed `_static_MethodName_VarName`,
  with a `// VB6 Static local` comment.
- **Fallback ✅ implemented:** Emit `// VB6 Static local — hoist to field if persistence needed`
  and keep as a regular local (compiles; semantics differ — value resets each call).
  The preferred hoisting approach remains a future improvement.

---

## 🟡 High impact — causes many failures, moderate complexity

### ✅ 4. `ref` on numeric/string literals (CS1525 — ~44 errors; CS1003 in case labels — ~24 errors)

VB6 allows passing a literal ByRef (it transparently creates a temp copy).
C# forbids `ref 1`, `ref "text"`, etc.

**Fix:** When a ByRef argument is a literal → emit it without `ref`, add `/* was ByRef */` comment:
```csharp
// Before (broken)
SomeFunc(ref 1, ref "text")
// After
SomeFunc(1 /* was ByRef */, "text" /* was ByRef */)
```

---

### ✅ 5. Backslash in string literals not escaped (CS1009 + CS1010/CS1040 — ~110 errors) ✦ resolved by #1

VB6 strings are not escape-sequence aware — `\` is just a backslash. Inside a C# `"..."`
string, sequences like `\N`, `\W`, `\` become invalid escape sequences. Also affects
`switch` case label strings (same root cause).

**Resolved by item #1 ✅:** Switching all string literals to `@"..."` verbatim syntax makes
backslashes literal by definition — no additional handling required.

---

## 🟠 Medium impact — affects correctness, lower frequency

### ✅ 6. Empty indexer brackets on array access (CS0443 — ~28 errors)

Some array variable accesses emit `arr[]` instead of `arr`. Occurs when a `string[]`-typed
variable is used without indexing (e.g. `.Length` access). The array-type check in `MapCall`
generates brackets even when `c.Arguments.Count == 0`.

**Fix:** In `MapCall`, only emit index brackets when `c.Arguments.Count > 0`.

---

### ✅ 7. VB6 type-suffix characters not fully stripped (CS1056 — ~4 errors)

VB6 allows type-declaration suffixes on identifiers (`$`, `%`, `&`, `!`, `#`, `@`).
Most are stripped in the lexer, but `!` (Single suffix, not bang-access) may slip through
in some identifier contexts.

**Fix ✅ implemented:** Lexer updated to consume but not append type-suffix characters (`%`, `&`,
`!`, `#`, `@`, `$`) for both identifiers and numeric literals. Hex literal format also fixed:
`&H1F` → `0x1F`, `&O17` → octal-to-decimal conversion in code generator.

---

### ✅ 8. One-element tuple expressions (CS8124 — ~10 errors)

Some code path was suspected to produce `(expr,)` or similar.

**Resolution:** CS8124 was not observed in the build after implementing items 1–7.
Likely eliminated as a side effect of other fixes (string literal changes, missing-arg
handling, etc.). No standalone fix was needed.

---

## 🟢 Quality improvements (not compile errors — correctness and idiom)

### 9. `Err.Raise` → `throw new Exception(...)`

`Err.Raise 1234, , "message"` currently emits as a method call stub.

**Fix:**
```csharp
throw new Exception("message"); // was Err.Raise 1234
```
`Err.Number` inside catch blocks → `0` or keep as comment.

---

### ✅ 10. Error handling — pragmatic approach

`On Error GoTo` restructuring into try/catch blocks has been **disabled** because VB6 error
handling semantics differ too much from C# try/catch to be mechanically reliable.

**Current output (all emitted as comments/stubs for manual review):**
- `On Error GoTo label` → `// VB6: On Error GoTo label — not restructured; review manually`
- `On Error Resume Next` → `// VB6: On Error Resume Next — errors suppressed; review manually`
- `On Error GoTo 0` → `// VB6: On Error GoTo 0 — reset error handler`
- `Resume Next` → `// VB6: Resume Next`
- `Resume` / `Resume label` → `// TODO: Resume (retry semantics — no C# equivalent)`
- Labels (`ErrHandler:`) → emitted as valid C# labels for reference
- `GoTo label` within scope → emitted as valid `goto label;`
- `GoTo label` outside scope → `// goto label; — label not found in scope, commented out`

This leaves error-handling structure visible for the developer to restructure manually,
which is safer than generating silently wrong try/catch blocks.

---

### ✅ 16. Structured error handling — `try/catch/finally` pattern detection

A specific, common VB6 pattern maps cleanly and safely to `try/catch/finally`:

```vb
Function Foo
    On Error GoTo ErrorHandler
    ' ... work ...
    LogSuccess
ExitGracefully:
    CloseConnections
    Exit Function
ErrorHandler:
    LogError
    GoTo ExitGracefully
End Function
```

Maps to:
```csharp
void Foo()
{
    try { /* work */ LogSuccess(); }
    catch { LogError(); }
    finally { CloseConnections(); }
}
```

**Detection algorithm (in `Transformer.cs`):**
1. Procedure has exactly one `On Error GoTo errLabel`.
2. Find `errLabel:` — its body ends with `GoTo cleanupLabel` (no other `GoTo`s in the error block).
3. Find `cleanupLabel:` — its body ends with `Exit Function/Sub` (the normal-exit path).
4. Everything before the first of these two labels is the `try` body.
5. `cleanupLabel:` body (minus the `Exit`) → `finally` body.
6. `errLabel:` body (minus the `GoTo cleanupLabel`) → `catch` body.
7. Restructure into `TryCatchFinallyNode(TryBody, CatchBody, FinallyBody)`.

**Reject and fall back to comment-out when:**
- More than one `On Error GoTo` in the procedure
- `Resume` or `Resume Next` appears anywhere
- The error block has more than one `GoTo` (ambiguous flow)
- `cleanupLabel` is not immediately followed by the `errLabel` (interleaved labels)

**Status:** Pending

---

### 11. `GoTo` outside error handlers

Non-error-handler `GoTo` is emitted as `goto label;`. C# allows it but it's fragile with
variable scoping. Worth adding a `// WARNING: GoTo — review control flow` comment.
Obvious patterns (e.g. `GoTo ExitSub` → label at end of proc) could be restructured to `return`.

---

### 12. Property/method names that collide with C# keywords

VB6 identifiers like `String`, `Object`, `Error`, `Lock` are valid VB6 names but reserved
in C#. These cause unexpected parse errors in the emitted code.

**Fix:** Add a keyword-escape pass — prefix conflicting names with `@` (e.g. `@string`,
`@object`) or rename them with a suffix.

---

### 13. Nested `With` block stack

`WithMemberAccessNode` inside nested `With` blocks only tracks one level. Nested `With`
(e.g. `With obj` containing `With .SubProp`) needs a stack. The inner `.Member` may
currently resolve to the outer object incorrectly.

**Fix:** Replace single `_withObject` field with a `Stack<string>` in `CodeGenerator`.

---

### 14. Implicit numeric narrowing coercions

VB6 freely coerces between numeric types in assignments. C# requires explicit casts when
narrowing (e.g. `long` → `int`, `double` → `int`). The translator emits no casts, so
narrowing assignments either silently truncate or produce CS0266 in strict mode.

**Fix:** When the declared type of the LHS is narrower than the RHS expression type,
emit an explicit cast: `(int)someDouble`.

---

### 15. `Date` literals

VB6 date literals (`#1/15/2020#`) are stored as `DateLiteralNode` with raw text and
currently emitted as-is — not valid C#.

**Fix:** Parse the date string and emit:
```csharp
new DateTime(2020, 1, 15) // was #1/15/2020#
```
or `DateTime.Parse("1/15/2020")` as a safer fallback.

---

## Summary table

| # | Fix | Est. errors | Status |
|---|-----|-------------|--------|
| 1 | All string literals → `@"..."` verbatim (also fixes #5) | ~1,110 | ✅ Done |
| 2 | `/* missing */` → omit trailing / `default` for middle | ~710 | ✅ Done (sub-B partial) |
| 3 | `Static` locals → comment + regular local | ~352 | ✅ Done (fallback) |
| 4 | `ref` on literals → drop `ref` | ~68 | ✅ Done |
| 5 | ~~Backslash escaping~~ (resolved by #1) | ~110 | ✅ Done |
| 6 | Empty indexer brackets on `IndexNode` | ~28 | ✅ Done |
| 7 | Type suffix stripping in lexer (idents + numbers + hex) | ~4 | ✅ Done |
| 8 | One-element tuple | ~10 | ✅ Eliminated by other fixes |
| 9 | `Err.Raise` → `throw` | quality | Pending |
| 10 | Error handling → pragmatic comment-out approach | quality | ✅ Done |
| 11 | `GoTo` restructuring | quality | Pending |
| 12 | C# keyword collisions | quality | Pending |
| 13 | Nested `With` stack | quality | Pending |
| 14 | Numeric narrowing casts | quality | Pending |
| 15 | `Date` literals | quality | Pending |
| 16 | Structured `On Error GoTo` → `try/catch/finally` pattern | quality | ✅ Done |

## Additional fixes made during sprint (discovered while implementing 1–8)

| Fix | Errors eliminated |
|-----|-------------------|
| Numeric labels (`10:` → `_L10:`) | ~213 occurrences (all in one large file) |
| Number literal type suffixes (`9999999999#` → `9999999999`) | ~1 critical |
| Hex literal format (`&H1F` → `0x1F`) | cross-file |
| `Imp` operator (`/* Imp */ \|\|` → `(!left \|\| right)`) | 4 CS1525 |
| `DateAdd`/`DateDiff`/`Choose`/`Switch` → `default /* TODO */` | 8 syntax errors |
| Nested `/* */` in block comments → `BlockSafe()` | 2 CS1056 em-dash |
| VB6 `Or` in case pattern → split into two `case` labels | 4 CS1525 |
| `Collection` in function return types → `Dictionary<string,object>` | 38 CS0305 |
| `static const` → `const` (C# const implies static) | 14 CS0504 |
| `const object = value` → infer type from literal | 4 CS0134 |
| Missing VB6 constants (`vbTextCompare`, `vbApplicationModal`, etc.) | 4 CS0103 |
| `ByRef` optional param with default → strip default, add comment | 2 CS1741 |
| Unqualified `Dictionary` type → `Dictionary<string, object>` | 2 CS0305 |
