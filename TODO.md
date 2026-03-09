# Translator Improvement Opportunities

---

## 🟢 Quality improvements (not compile errors — correctness and idiom)

### 1. `Err.Raise` → `throw new Exception(...)`

`Err.Raise 1234, , "message"` currently emits as a method call stub.

**Fix:**
```csharp
throw new Exception("message"); // was Err.Raise 1234
```
`Err.Number` inside catch blocks → `0` or keep as comment.

---

### 2. `GoTo` outside error handlers

Non-error-handler `GoTo` is emitted as `goto label;`. C# allows it but it's fragile with
variable scoping. Worth adding a `// WARNING: GoTo — review control flow` comment.
Obvious patterns (e.g. `GoTo ExitSub` → label at end of proc) could be restructured to `return`.

---

### 3. Mixed optional parameter ordering (CS1737 — 1 function remaining)

VB6 allows optional params before required ones: `Sub Foo(Optional a As String = "", b As String)`.
C# requires all optional params to come after required ones.

**Fix (two-step, requires cross-module tracking):**
1. **Definition site:** reorder params so required come first, optional last.
2. **Call sites:** build a cross-module map and use named arguments so each argument
   reaches the correctly reordered parameter.

Remaining: 2 CS1737 errors (1 function, duplicated by compiler pass).

---

### 4. Property/method names that collide with C# keywords

VB6 identifiers like `String`, `Object`, `Error`, `Lock` are valid VB6 names but reserved
in C#. These cause unexpected parse errors in the emitted code.

**Fix:** Add a keyword-escape pass — prefix conflicting names with `@` (e.g. `@string`,
`@object`) or rename them with a suffix.

---

### 5. Nested `With` block stack

`WithMemberAccessNode` inside nested `With` blocks only tracks one level. Nested `With`
(e.g. `With obj` containing `With .SubProp`) needs a stack. The inner `.Member` may
currently resolve to the outer object incorrectly.

**Fix:** Replace single `_withObject` field with a `Stack<string>` in `CodeGenerator`.

---

### 6. Implicit numeric narrowing coercions

VB6 freely coerces between numeric types in assignments. C# requires explicit casts when
narrowing (e.g. `long` → `int`, `double` → `int`). The translator emits no casts, so
narrowing assignments either silently truncate or produce CS0266 in strict mode.

**Fix:** When the declared type of the LHS is narrower than the RHS expression type,
emit an explicit cast: `(int)someDouble`.

---

### 7. `Date` literals

VB6 date literals (`#1/15/2020#`) are stored as `DateLiteralNode` with raw text and
currently emitted as-is — not valid C#.

**Fix:** Parse the date string and emit:
```csharp
new DateTime(2020, 1, 15) // was #1/15/2020#
```
or `DateTime.Parse("1/15/2020")` as a safer fallback.

---

## Summary table

| # | Fix | Status |
|---|-----|--------|
| 1 | `Err.Raise` → `throw new Exception(...)` | Pending |
| 2 | `GoTo` outside error handlers | Pending |
| 3 | Mixed optional parameter ordering (CS1737) | Pending |
| 4 | C# keyword collisions | Pending |
| 5 | Nested `With` stack | Pending |
| 6 | Numeric narrowing casts | Pending |
| 7 | `Date` literals | Pending |
