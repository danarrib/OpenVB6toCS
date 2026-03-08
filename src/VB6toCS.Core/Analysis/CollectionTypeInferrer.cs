using VB6toCS.Core.Parsing.Nodes;

namespace VB6toCS.Core.Analysis;

/// <summary>
/// Cross-module inference pass that deduces the element type of VB6 Collection fields
/// and properties by scanning all .Add(item) call sites and Set-assignment sites across
/// the entire project, then propagating types transitively until stable.
///
/// Algorithm (runs on Stage-3 ASTs, before Transformer):
///
///   Pass 1 — Build member maps (field types + property return types per module).
///   Pass 2 — Scan every .Add(item) call site; record (ownerModule, field) → elemType.
///   Pass 3 — Initialise _resolvedTypes from unambiguous Add-scan results.
///   Pass 4+ (loop) —
///     a. Scan every AssignmentNode: if RHS is a resolved Collection, propagate its
///        element type to the LHS Collection field/property.
///     b. Scan every Property getter FunctionReturnNode: if the returned expression is
///        a resolved Collection, record the property's return element type.
///     Repeat until no new types are resolved.
///
/// A field/property is resolved to Collection&lt;object&gt; as fallback when inference fails.
/// </summary>
public sealed class CollectionTypeInferrer
{
    // module.Name → { memberName (ci) → VB6 declared/return type }
    // Includes both FieldNode declarators and CsPropertyNode getter return types.
    private readonly Dictionary<string, Dictionary<string, string>> _moduleMembers =
        new(StringComparer.OrdinalIgnoreCase);

    // (moduleName, memberName) → distinct element types observed at Add call sites only.
    private readonly Dictionary<(string, string), HashSet<string>> _hits = new();

    // Working set of resolved Collection element types (fields + properties).
    // Updated during both the Add scan and the propagation loop.
    private readonly Dictionary<(string, string), string> _resolvedTypes =
        new();

    // Keys that received conflicting types from different sites — excluded from results.
    private readonly HashSet<(string, string)> _conflicted = new();

    // For each (module, field) that received at least one Add call:
    //   _keyedAdds    — had at least one Add with a non-missing key argument (arg[1])
    //   _unkeyedAdds  — had at least one Add without a key argument
    private readonly HashSet<(string, string)> _keyedAdds   = new();
    private readonly HashSet<(string, string)> _unkeyedAdds = new();

    // ── Public API ────────────────────────────────────────────────────────────

    /// <summary>Run all inference passes on the full set of Stage-3 modules.</summary>
    public void Analyse(IEnumerable<ModuleNode> modules)
    {
        var list = modules.ToList();

        // Pass 1: build member maps (fields + property return types).
        foreach (var module in list)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var f in module.Members.OfType<FieldNode>())
                foreach (var d in f.Declarators)
                    if (d.TypeRef != null)
                        map[d.Name] = d.TypeRef.TypeName;
            foreach (var p in module.Members.OfType<CsPropertyNode>())
                if (p.Type?.TypeName != null)
                    map[p.Name] = p.Type.TypeName;
            _moduleMembers[module.Name] = map;
        }

        // Pass 2: scan Add call sites.
        foreach (var module in list)
            ScanModule(module);

        // Pass 3: seed _resolvedTypes from unambiguous Add-scan results.
        foreach (var (key, types) in _hits)
            if (types.Count == 1)
                TryResolve(key, types.First(), out _);

        // Pass 4+: propagation loop — assignment propagation + property return inference.
        bool anyNew;
        do
        {
            anyNew = false;
            foreach (var module in list)
                anyNew |= PropagateModule(module);
        } while (anyNew);
    }

    /// <summary>
    /// Returns a read-only snapshot of all resolved Collection element types after
    /// inference is complete. Key: (moduleName, fieldOrPropertyName). Value: element type.
    /// </summary>
    public IReadOnlyDictionary<(string, string), string> GetInferredTypes() =>
        _resolvedTypes;

    /// <summary>
    /// Returns the single inferred element type for a specific field/property,
    /// or null if it could not be resolved.
    /// </summary>
    public string? GetElementType(string moduleName, string fieldName) =>
        _resolvedTypes.TryGetValue((moduleName, fieldName), out var t) ? t : null;

    /// <summary>
    /// Returns a map from (moduleName, fieldName) to the inferred <see cref="CollectionKind"/>
    /// for every Collection field/property where Add usage was observed.
    /// Fields with mixed key/no-key usage are omitted (they default to <see cref="CollectionKind.Collection"/>).
    /// </summary>
    public IReadOnlyDictionary<(string, string), CollectionKind> GetInferredKinds()
    {
        var result = new Dictionary<(string, string), CollectionKind>();
        var allKeys = new HashSet<(string, string)>(_keyedAdds);
        allKeys.UnionWith(_unkeyedAdds);

        foreach (var key in allKeys)
        {
            // List only when every observed Add call had no key argument.
            // Dictionary for always-keyed AND mixed (keyed + unkeyed) calls.
            result[key] = _unkeyedAdds.Contains(key) && !_keyedAdds.Contains(key)
                ? CollectionKind.List
                : CollectionKind.Dictionary;
        }
        return result;
    }

    // ── Add-call scanning (Pass 2) ────────────────────────────────────────────

    private void ScanModule(ModuleNode module)
    {
        foreach (var member in module.Members)
        {
            switch (member)
            {
                case SubNode s:
                    ScanProcedureForAdds(module.Name, s.Parameters, s.Body);
                    break;
                case FunctionNode f:
                    ScanProcedureForAdds(module.Name, f.Parameters, f.Body);
                    break;
                case CsPropertyNode p:
                    if (p.GetBody != null) ScanProcedureForAdds(module.Name, p.GetParameters, p.GetBody);
                    if (p.LetBody != null) ScanProcedureForAdds(module.Name, p.LetParameters, p.LetBody);
                    if (p.SetBody != null) ScanProcedureForAdds(module.Name, p.SetParameters, p.SetBody);
                    break;
            }
        }
    }

    private void ScanProcedureForAdds(
        string moduleName,
        IReadOnlyList<ParameterNode> parameters,
        IReadOnlyList<AstNode> body)
    {
        var locals = BuildLocals(parameters, body);
        // Pre-pass: infer element types for local Collection variables within this procedure
        // (e.g. "Dim lCampos As Collection" where lCampos.Add lCampo (clsYasField) is called).
        // This lets us emit "Collection<clsYasField>" rather than bare "Collection" when lCampos
        // is subsequently added to a module-level Collection field.
        var localCollElemTypes = BuildLocalCollElemTypes(body, moduleName, locals);
        ScanBodyForAdds(body, moduleName, locals, localCollElemTypes, new Stack<string?>());
    }

    private void ScanBodyForAdds(
        IReadOnlyList<AstNode> body,
        string moduleName,
        Dictionary<string, string> locals,
        IReadOnlyDictionary<string, string> localCollElemTypes,
        Stack<string?> withStack)
    {
        foreach (var node in body)
        {
            RecordAddCall(node, moduleName, locals, localCollElemTypes, withStack);

            if (node is WithNode w)
            {
                withStack.Push(InferMemberType(w.Object, moduleName, locals));
                ScanBodyForAdds(w.Body, moduleName, locals, localCollElemTypes, withStack);
                withStack.Pop();
            }
            else
            {
                foreach (var child in ChildBodies(node))
                    ScanBodyForAdds(child, moduleName, locals, localCollElemTypes, withStack);
            }
        }
    }

    private void RecordAddCall(
        AstNode node,
        string moduleName,
        Dictionary<string, string> locals,
        IReadOnlyDictionary<string, string> localCollElemTypes,
        Stack<string?> withStack)
    {
        if (node is not CallStatementNode call) return;

        // col.Add item      → Target=MemberAccess(col,"Add"), Args=[item]
        // col.Add(item)     → Target=CallOrIndex(MemberAccess(col,"Add"),[item]), Args=[]
        MemberAccessNode? addMa;
        IReadOnlyList<ArgumentNode> addArgs;

        if (call.Target is MemberAccessNode ma1 &&
            ma1.MemberName.Equals("Add", StringComparison.OrdinalIgnoreCase))
        {
            addMa = ma1;  addArgs = call.Arguments;
        }
        else if (call.Target is CallOrIndexNode coi &&
                 coi.Target is MemberAccessNode ma2 &&
                 ma2.MemberName.Equals("Add", StringComparison.OrdinalIgnoreCase))
        {
            addMa = ma2;  addArgs = coi.Arguments;
        }
        else return;

        // First non-missing argument is the item being added.
        // For named args like "Item:=x", the first arg still carries the value.
        var itemArg = addArgs.FirstOrDefault(a => !a.IsMissing);
        if (itemArg?.Value == null) return;

        string? elemType = InferMemberType(itemArg.Value, moduleName, locals);
        if (elemType == null) return;

        // If the item being added is itself a local Collection variable whose element type
        // we pre-resolved (e.g. lCampos whose items are clsYasField), upgrade the element
        // type from bare "Collection" to the parameterised "Collection<innerElem>" so that
        // the outer field is typed Collection<Collection<innerElem>> rather than
        // Collection<Collection>.
        if (elemType.Equals("Collection", StringComparison.OrdinalIgnoreCase) &&
            itemArg.Value is IdentifierNode itemId &&
            localCollElemTypes.TryGetValue(itemId.Name, out var innerElem))
        {
            // Local Collection variables are now translated to Dictionary<string,T> by default,
            // so represent the nested element type accordingly.
            elemType = $"Dictionary<string, {innerElem}>";
        }

        var target = ResolveCollectionTarget(addMa.Object, moduleName, locals, withStack);
        if (target == null) return;

        var key = target.Value;
        if (!_hits.TryGetValue(key, out var set))
            _hits[key] = set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        set.Add(elemType);

        // Track whether this Add call includes a string key (second argument).
        bool hasKeyArg = addArgs.Count >= 2 &&
                         !addArgs[1].IsMissing &&
                         addArgs[1].Value != null;
        if (hasKeyArg) _keyedAdds.Add(key);
        else           _unkeyedAdds.Add(key);
    }

    // ── Local Collection element-type pre-pass ────────────────────────────────

    /// <summary>
    /// Scans a procedure body to infer element types for local Collection variables.
    /// For each <c>Dim x As Collection</c> where only one distinct item type is ever
    /// added to <c>x</c>, records <c>x → elemType</c>.  Conflicting types are excluded.
    /// </summary>
    private Dictionary<string, string> BuildLocalCollElemTypes(
        IReadOnlyList<AstNode> body,
        string moduleName,
        Dictionary<string, string> locals)
    {
        var result    = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var conflicts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var localColls = locals
            .Where(kv => kv.Value.Equals("Collection", StringComparison.OrdinalIgnoreCase))
            .Select(kv => kv.Key)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        if (localColls.Count == 0) return result;

        ScanBodyForLocalCollAdds(body, moduleName, locals, localColls, result, conflicts);
        return result;
    }

    private void ScanBodyForLocalCollAdds(
        IReadOnlyList<AstNode> body,
        string moduleName,
        Dictionary<string, string> locals,
        HashSet<string> localColls,
        Dictionary<string, string> result,
        HashSet<string> conflicts)
    {
        foreach (var node in body)
        {
            if (node is CallStatementNode call)
            {
                MemberAccessNode? addMa = null;
                IReadOnlyList<ArgumentNode> addArgs = [];

                if (call.Target is MemberAccessNode ma1 &&
                    ma1.MemberName.Equals("Add", StringComparison.OrdinalIgnoreCase))
                { addMa = ma1; addArgs = call.Arguments; }
                else if (call.Target is CallOrIndexNode coi &&
                         coi.Target is MemberAccessNode ma2 &&
                         ma2.MemberName.Equals("Add", StringComparison.OrdinalIgnoreCase))
                { addMa = ma2; addArgs = coi.Arguments; }

                if (addMa?.Object is IdentifierNode collId &&
                    localColls.Contains(collId.Name) &&
                    !conflicts.Contains(collId.Name))
                {
                    var itemArg = addArgs.FirstOrDefault(a => !a.IsMissing);
                    if (itemArg?.Value != null)
                    {
                        string? elemType = InferMemberType(itemArg.Value, moduleName, locals);
                        if (elemType != null)
                        {
                            if (!result.TryGetValue(collId.Name, out var existing))
                                result[collId.Name] = elemType;
                            else if (!existing.Equals(elemType, StringComparison.OrdinalIgnoreCase))
                            {
                                result.Remove(collId.Name);
                                conflicts.Add(collId.Name);
                            }
                        }
                    }
                }
            }

            if (node is WithNode w)
                ScanBodyForLocalCollAdds(w.Body, moduleName, locals, localColls, result, conflicts);
            else
                foreach (var child in ChildBodies(node))
                    ScanBodyForLocalCollAdds(child, moduleName, locals, localColls, result, conflicts);
        }
    }

    // ── Propagation (Pass 4+) ─────────────────────────────────────────────────

    private bool PropagateModule(ModuleNode module)
    {
        bool anyNew = false;
        foreach (var member in module.Members)
        {
            switch (member)
            {
                case SubNode s:
                    anyNew |= PropagateBody(
                        s.Body, module.Name,
                        BuildLocals(s.Parameters, s.Body),
                        new Stack<string?>(), returnKey: null);
                    break;

                case FunctionNode f:
                    anyNew |= PropagateBody(
                        f.Body, module.Name,
                        BuildLocals(f.Parameters, f.Body),
                        new Stack<string?>(), returnKey: null);
                    break;

                case CsPropertyNode p:
                    // Getter: propagate return type as well as assignments inside.
                    if (p.GetBody != null)
                    {
                        var getLocals = BuildLocals(p.GetParameters, p.GetBody);
                        anyNew |= PropagateBody(
                            p.GetBody, module.Name, getLocals,
                            new Stack<string?>(),
                            returnKey: (module.Name, p.Name));
                    }
                    if (p.LetBody != null)
                        anyNew |= PropagateBody(
                            p.LetBody, module.Name,
                            BuildLocals(p.LetParameters, p.LetBody),
                            new Stack<string?>(), returnKey: null);
                    if (p.SetBody != null)
                        anyNew |= PropagateBody(
                            p.SetBody, module.Name,
                            BuildLocals(p.SetParameters, p.SetBody),
                            new Stack<string?>(), returnKey: null);
                    break;
            }
        }
        return anyNew;
    }

    private bool PropagateBody(
        IReadOnlyList<AstNode> body,
        string moduleName,
        Dictionary<string, string> locals,
        Stack<string?> withStack,
        (string, string)? returnKey)
    {
        bool anyNew = false;
        foreach (var node in body)
        {
            anyNew |= TryPropagateStatement(node, moduleName, locals, withStack, returnKey);

            if (node is WithNode w)
            {
                withStack.Push(InferMemberType(w.Object, moduleName, locals));
                anyNew |= PropagateBody(w.Body, moduleName, locals, withStack, returnKey);
                withStack.Pop();
            }
            else
            {
                foreach (var child in ChildBodies(node))
                    anyNew |= PropagateBody(child, moduleName, locals, withStack, returnKey);
            }
        }
        return anyNew;
    }

    private bool TryPropagateStatement(
        AstNode node,
        string moduleName,
        Dictionary<string, string> locals,
        Stack<string?> withStack,
        (string, string)? returnKey)
    {
        // ── Assignment: X = Y where Y is a resolved Collection ────────────────
        if (node is AssignmentNode a)
        {
            string? elemType = InferCollectionElementType(a.Value, moduleName, locals, withStack);
            if (elemType != null)
            {
                var target = ResolveCollectionTarget(a.Target, moduleName, locals, withStack);
                if (target != null)
                {
                    TryResolve(target.Value, elemType, out bool changed);
                    return changed;
                }
            }
        }

        // ── Property getter return: FunctionReturnNode → infer property type ──
        if (node is FunctionReturnNode ret && returnKey != null)
        {
            string? elemType = InferCollectionElementType(ret.Value, moduleName, locals, withStack);
            if (elemType != null)
            {
                TryResolve(returnKey.Value, elemType, out bool changed);
                return changed;
            }
        }

        return false;
    }

    // ── Shared resolution helpers ─────────────────────────────────────────────

    /// <summary>
    /// Resolves the collection being targeted (field or property) as (ownerModule, name).
    /// Returns null when the target is not a recognisable unresolved Collection reference.
    /// </summary>
    private (string mod, string name)? ResolveCollectionTarget(
        ExpressionNode target,
        string moduleName,
        Dictionary<string, string> locals,
        Stack<string?> withStack)
    {
        switch (target)
        {
            case IdentifierNode id:
                if (_moduleMembers.TryGetValue(moduleName, out var mf) &&
                    mf.TryGetValue(id.Name, out var type) &&
                    type.Equals("Collection", StringComparison.OrdinalIgnoreCase))
                    return (moduleName, id.Name);
                break;

            case WithMemberAccessNode wma:
            {
                string? withType = withStack.Count > 0 ? withStack.Peek() : null;
                if (withType != null &&
                    _moduleMembers.TryGetValue(withType, out var wf) &&
                    wf.TryGetValue(wma.MemberName, out var ft) &&
                    ft.Equals("Collection", StringComparison.OrdinalIgnoreCase))
                    return (withType, wma.MemberName);
                break;
            }

            case MemberAccessNode ma:
            {
                string? ownerType = InferMemberType(ma.Object, moduleName, locals);
                if (ownerType != null &&
                    _moduleMembers.TryGetValue(ownerType, out var of) &&
                    of.TryGetValue(ma.MemberName, out var ft) &&
                    ft.Equals("Collection", StringComparison.OrdinalIgnoreCase))
                    return (ownerType, ma.MemberName);
                break;
            }
        }
        return null;
    }

    /// <summary>
    /// Returns the Collection element type of an expression if it is already resolved,
    /// or null when the element type is unknown.
    /// </summary>
    private string? InferCollectionElementType(
        ExpressionNode expr,
        string moduleName,
        Dictionary<string, string> locals,
        Stack<string?> withStack)
    {
        switch (expr)
        {
            case IdentifierNode id:
                if (_resolvedTypes.TryGetValue((moduleName, id.Name), out var t1)) return t1;
                break;

            case WithMemberAccessNode wma:
            {
                string? withType = withStack.Count > 0 ? withStack.Peek() : null;
                if (withType != null &&
                    _resolvedTypes.TryGetValue((withType, wma.MemberName), out var t2)) return t2;
                break;
            }

            case MemberAccessNode ma:
            {
                string? ownerType = InferMemberType(ma.Object, moduleName, locals);
                if (ownerType != null &&
                    _resolvedTypes.TryGetValue((ownerType, ma.MemberName), out var t3)) return t3;
                break;
            }
        }
        return null;
    }

    /// <summary>
    /// Best-effort VB6 type name for an expression (used to resolve member owners).
    /// Returns null when the type cannot be statically determined.
    /// </summary>
    private string? InferMemberType(
        ExpressionNode expr,
        string moduleName,
        Dictionary<string, string> locals)
    {
        switch (expr)
        {
            case NewObjectNode n:
                return n.TypeName;

            case IdentifierNode id:
                if (locals.TryGetValue(id.Name, out var lt)) return lt;
                if (_moduleMembers.TryGetValue(moduleName, out var mf) &&
                    mf.TryGetValue(id.Name, out var ft)) return ft;
                return null;

            case MeNode:
                return moduleName;

            case MemberAccessNode ma:
            {
                string? ownerType = InferMemberType(ma.Object, moduleName, locals);
                if (ownerType != null &&
                    _moduleMembers.TryGetValue(ownerType, out var of) &&
                    of.TryGetValue(ma.MemberName, out var memberType))
                    return memberType;
                return null;
            }

            default:
                return null;
        }
    }

    // ── Resolution bookkeeping ────────────────────────────────────────────────

    /// <summary>
    /// Attempts to record elemType for key.
    /// If key was already resolved to the same type: no-op.
    /// If key was already resolved to a different type: marks as conflicted (removed).
    /// Sets changed=true only when the resolution set actually changes.
    /// </summary>
    private void TryResolve((string, string) key, string elemType, out bool changed)
    {
        if (_conflicted.Contains(key)) { changed = false; return; }

        if (_resolvedTypes.TryGetValue(key, out var existing))
        {
            if (existing.Equals(elemType, StringComparison.OrdinalIgnoreCase))
            {
                changed = false;
            }
            else
            {
                // Conflict: two different types → demote to unresolved.
                _resolvedTypes.Remove(key);
                _conflicted.Add(key);
                changed = true;
            }
        }
        else
        {
            _resolvedTypes[key] = elemType;
            changed = true;
        }
    }

    // ── Local variable map builder ────────────────────────────────────────────

    private static Dictionary<string, string> BuildLocals(
        IReadOnlyList<ParameterNode> parameters,
        IReadOnlyList<AstNode> body)
    {
        var locals = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var p in parameters)
            if (p.TypeRef != null)
                locals[p.Name] = p.TypeRef.TypeName;
        CollectLocals(body, locals);
        return locals;
    }

    private static void CollectLocals(IReadOnlyList<AstNode> body, Dictionary<string, string> locals)
    {
        foreach (var node in body)
        {
            if (node is LocalDimNode dim)
                foreach (var d in dim.Declarators)
                    if (d.TypeRef != null)
                        locals[d.Name] = d.TypeRef.TypeName;
            foreach (var child in ChildBodies(node))
                CollectLocals(child, locals);
        }
    }

    // ── Child-body extractor ──────────────────────────────────────────────────

    private static IEnumerable<IReadOnlyList<AstNode>> ChildBodies(AstNode node)
    {
        switch (node)
        {
            case IfNode i:
                yield return i.ThenBody;
                foreach (var ei in i.ElseIfClauses) yield return ei.Body;
                if (i.ElseBody != null) yield return i.ElseBody;
                break;
            case SingleLineIfNode s:
                yield return [s.ThenStatement];
                if (s.ElseStatement != null) yield return [s.ElseStatement];
                break;
            case SelectCaseNode s:
                foreach (var c in s.Cases) yield return c.Body;
                break;
            case ForNextNode  f: yield return f.Body; break;
            case ForEachNode  f: yield return f.Body; break;
            case WhileNode    w: yield return w.Body; break;
            case DoLoopNode   d: yield return d.Body; break;
            // WithNode is handled directly in ScanBodyForAdds / PropagateBody (push/pop stack).
        }
    }
}
