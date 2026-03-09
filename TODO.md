# Translator Improvement Opportunities

Build baseline: `D46O003_1080_nocom.csproj` (137 translated files, no COM references).

---

## 🔴 Genuine translator errors

### ✅ 8. Labels before closing brace (CS1525 / CS1002 — 236 errors)

C# requires a label to be followed by at least one statement. VB6 allows labels
as the very last line of a block (e.g. just before `End If` or `Next`), which
generated bare `label:` with no following statement.

**Fix:** Emit `label: ;` (label + C# empty statement) for every `LabelNode`.

---

### 9. Cross-module nested enum type names (CS0246 — ~100 errors)

In the generated C#, all enums are **nested types** of their declaring class
(e.g. `e_CodRamo` is defined inside `stcA46V702B`). When another class uses
`e_CodRamo` as a field type, C# requires the fully-qualified form
`stcA46V702B.e_CodRamo`. The translator currently emits the bare enum name,
causing CS0246 "type not found".

Affected enum type names (examples):
`e_CodRamo`, `e_TipPessoa`, `E_IDConsist`, `eTipCalcRPT`, `eRegraPlano`,
`e_TipoPrazo`, `e_TipEmissao`, `e_TipDoc`, `E_TipConsist`, `e_SimNao`,
`e_FormaPagamento`.

**Fix:** Build a cross-module `enumTypeMap` (`enumTypeName → moduleName`) alongside
the existing `enumMemberMap`. In `TypeStr()` (code generator), when emitting a
`TypeRefNode` whose type name matches an enum from a *different* module, qualify it
as `ModuleName.EnumTypeName`.

Same logic already used for enum *member* identifiers — just extend it to the
type name itself in declarations, parameter types, and return types.

---

### 10. VB6 built-in enum types as optional parameter defaults (CS1750 — 2 errors)

Two functions use VB6/VBA built-in enum types as optional parameter types with
integer defaults:

```csharp
// Generated (broken)
void gpGraLog(byte pOpcao, string pMensagem, VbMsgBoxStyle pMsgBoxStyle = 0)
int Find(string sToFind, int lStartIndex = 1, VbCompareMethod compare = 1)
```

`VbMsgBoxStyle` and `VbCompareMethod` are VB6/VBA runtime enums not available
in .NET. CS1750 fires because `int` literal `0` / `1` can't be implicitly
converted to an undefined enum type.

**Fix options (in order of preference):**
1. Map known VB6 enum types to their .NET equivalents in `BuiltInMap`:
   - `VbCompareMethod` → `StringComparison` (0=Binary, 1=Text)
   - `VbMsgBoxStyle` → `int` (no .NET equivalent; downgrade the param type)
2. For unmapped VB6 enum types, emit `int` as the parameter type and keep
   the integer default value — preserves compilability.

---

### 11. Mixed optional parameter ordering (CS1737 — 2 errors, 1 function)

Already in the TODO. One function in `clsA28V720A` has optional parameters
before required ones (VB6 allows this; C# does not). Cross-module param
reordering map required.

---

## 🟡 Expected / non-translator errors

These errors appear in the no-COM build and are **expected** — they would be
resolved by the full build with COM interop assemblies:

| Error | Count | Cause |
|-------|-------|-------|
| CS0246 `ADODB` | 174 | ADO COM library not referenced in no-COM build |
| CS0246 cross-project types | ~8 | Types from other syas projects not compiled together |

---

## 🟢 Quality improvements (not compile errors — correctness and idiom)

### 1. `Err.Raise` → `throw new Exception(...)`

`Err.Raise 1234, , "message"` currently emits as a plain method call stub.

**Fix:**
```csharp
throw new Exception("message"); // was Err.Raise 1234
```

---

### 2. `GoTo` outside error handlers

Non-error-handler `GoTo` statements that don't jump to the error label are
emitted as `goto label;`. C# allows it but it's fragile with variable scoping.
Obvious patterns (e.g. `GoTo ExitSub` → label at end of proc) could be
restructured to `return`.

---

### 3. Mixed optional parameter ordering (CS1737 — 1 function)

See item #11 above (also a compile error).

---

### 4. Property/method names that collide with C# keywords

VB6 identifiers like `String`, `Object`, `Error`, `Lock` are valid VB6 names
but reserved in C#.

**Fix:** Prefix conflicting names with `@` (e.g. `@string`, `@object`).

---

### 5. Nested `With` block stack

`WithMemberAccessNode` inside nested `With` blocks only tracks one level.

**Fix:** Replace single `_withObject` field with a `Stack<string>`.

---

### 6. Implicit numeric narrowing coercions

VB6 freely coerces between numeric types. C# requires explicit casts when
narrowing (e.g. `double` → `int`).

**Fix:** Emit `(int)someDouble` when LHS is narrower than RHS expression type.

---

### 7. `Date` literals

VB6 date literals (`#1/15/2020#`) are currently emitted as
`DateTime.Parse(rawText) /* date literal */` which is not quite right.

**Fix:** Parse the date string and emit `new DateTime(2020, 1, 15)`.

---

## Summary table

| # | Fix | Errors | Status |
|---|-----|--------|--------|
| 8 | Labels before closing brace → `label: ;` | 236 | ✅ Done |
| 9 | Cross-module enum type qualification | ~100 | Pending |
| 10 | VB6 built-in enum types as param defaults | 2 | Pending |
| 11 | Mixed optional parameter ordering | 2 | Pending |
| — | ADODB / cross-project types | ~182 | Expected (not translator) |
| 1 | `Err.Raise` → `throw` | quality | Pending |
| 2 | `GoTo` restructuring | quality | Pending |
| 4 | C# keyword collisions | quality | Pending |
| 5 | Nested `With` stack | quality | Pending |
| 6 | Numeric narrowing casts | quality | Pending |
| 7 | `Date` literals | quality | Pending |
