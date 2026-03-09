# Translator Improvement Opportunities

Build baseline: `D46O003_1080_nocom.csproj` (137 translated files, no COM references).
Current status: **zero genuine translator errors**. All remaining build failures are
expected COM/cross-project assembly reference absences.

---

## 🟡 Remaining compile errors (genuine translator issues)

### 1. Mixed optional parameter ordering (CS1737 — 1 function)

VB6 allows optional params before required ones: `Sub Foo(Optional a As String = "", b As String)`.
C# requires all optional params to come after required ones.

**Fix (two-step, requires cross-module tracking):**
1. **Definition site:** reorder params so required come first, optional last.
2. **Call sites:** build a cross-module map and use named arguments so each argument
   reaches the correctly reordered parameter.

Remaining: 2 CS1737 errors (1 function, duplicated by compiler pass).

---

## 🟢 Quality improvements (not compile errors — correctness and idiom)

### 2. `Err.Raise` → `throw new Exception(...)`

`Err.Raise 1234, , "message"` currently emits as a plain method call stub.

**Fix:**
```csharp
throw new Exception("message"); // was Err.Raise 1234
```

---

### 3. `GoTo` outside error handlers

Non-error-handler `GoTo` statements (not jumping to the error label) are emitted
as `goto label;`. Obvious patterns (e.g. `GoTo ExitSub` → label at end of proc)
could be restructured to `return`.

---

### 4. Property/method names that collide with C# keywords

VB6 identifiers like `String`, `Object`, `Error`, `Lock` are valid VB6 names but
reserved in C#.

**Fix:** Prefix conflicting names with `@` (e.g. `@string`, `@object`).

---

### 5. Nested `With` block stack

`WithMemberAccessNode` inside nested `With` blocks only tracks one level.

**Fix:** Replace single `_withObject` field with a `Stack<string>`.

---

### 6. Implicit numeric narrowing coercions

VB6 freely coerces between numeric types. C# requires explicit casts when narrowing
(e.g. `double` → `int`).

**Fix:** Emit `(int)someDouble` when LHS is narrower than RHS expression type.

---

### 7. `Date` literals

VB6 date literals (`#1/15/2020#`) currently emit `DateTime.Parse(rawText)`.

**Fix:** Parse the date string and emit `new DateTime(2020, 1, 15)`.

---

## Summary table

| # | Fix | Status |
|---|-----|--------|
| 1 | Mixed optional parameter ordering (CS1737) | Pending |
| 2 | `Err.Raise` → `throw new Exception(...)` | Pending |
| 3 | `GoTo` restructuring | Pending |
| 4 | C# keyword collisions | Pending |
| 5 | Nested `With` stack | Pending |
| 6 | Numeric narrowing casts | Pending |
| 7 | `Date` literals | Pending |
