using VB6toCS.Core.Parsing.Nodes;

namespace VB6toCS.Core.Analysis;

/// <summary>
/// Cross-module inference pass that deduces the element type of VB6 Collection fields
/// by scanning all .Add(item) call sites across the entire project.
///
/// A Collection field is assigned a concrete element type when every observed Add call
/// passes items of the same type.  If zero Add calls are found, or items of more than
/// one type are added, the field is left unresolved (will become Collection&lt;object&gt;).
///
/// Must run on Stage-3 ASTs (after Analyser, before Transformer).
/// </summary>
public sealed class CollectionTypeInferrer
{
    // module.Name → { fieldName (ci) → declared VB6 type name }
    private readonly Dictionary<string, Dictionary<string, string>> _moduleFields =
        new(StringComparer.OrdinalIgnoreCase);

    // (moduleName, fieldName) → distinct element type names observed at Add call sites
    private readonly Dictionary<(string, string), HashSet<string>> _hits = new();

    // ── Public API ────────────────────────────────────────────────────────────

    /// <summary>Run the two-pass inference on all Stage-3 modules.</summary>
    public void Analyse(IEnumerable<ModuleNode> modules)
    {
        var list = modules.ToList();

        // Pass 1: record every field's declared type for every module.
        foreach (var module in list)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var f in module.Members.OfType<FieldNode>())
                foreach (var d in f.Declarators)
                    if (d.TypeRef != null)
                        map[d.Name] = d.TypeRef.TypeName;
            _moduleFields[module.Name] = map;
        }

        // Pass 2: scan every procedure body for collection.Add(item) calls.
        foreach (var module in list)
            ScanModule(module);
    }

    /// <summary>
    /// Returns a dictionary of all unambiguously inferred collection types.
    /// Key: (moduleName, fieldName).  Value: element type name.
    /// Only fields where every Add call agreed on a single type are included.
    /// </summary>
    public IReadOnlyDictionary<(string, string), string> GetInferredTypes()
    {
        var result = new Dictionary<(string, string), string>();
        foreach (var (key, types) in _hits)
            if (types.Count == 1)
                result[key] = types.First();
        return result;
    }

    /// <summary>
    /// Returns the single inferred element type for the given Collection field,
    /// or null when the type could not be unambiguously determined.
    /// </summary>
    public string? GetElementType(string moduleName, string fieldName) =>
        _hits.TryGetValue((moduleName, fieldName), out var types) && types.Count == 1
            ? types.First()
            : null;

    // ── Module / procedure scanning ───────────────────────────────────────────

    private void ScanModule(ModuleNode module)
    {
        foreach (var member in module.Members)
        {
            switch (member)
            {
                case SubNode s:
                    ScanProcedure(module.Name, s.Parameters, s.Body);
                    break;
                case FunctionNode f:
                    ScanProcedure(module.Name, f.Parameters, f.Body);
                    break;
                case CsPropertyNode p:
                    if (p.GetBody != null) ScanProcedure(module.Name, p.GetParameters, p.GetBody);
                    if (p.LetBody != null) ScanProcedure(module.Name, p.LetParameters, p.LetBody);
                    if (p.SetBody != null) ScanProcedure(module.Name, p.SetParameters, p.SetBody);
                    break;
            }
        }
    }

    private void ScanProcedure(
        string moduleName,
        IReadOnlyList<ParameterNode> parameters,
        IReadOnlyList<AstNode> body)
    {
        var locals = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var p in parameters)
            if (p.TypeRef != null)
                locals[p.Name] = p.TypeRef.TypeName;
        CollectLocals(body, locals);

        ScanBody(body, moduleName, locals);
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

    private void ScanBody(
        IReadOnlyList<AstNode> body,
        string moduleName,
        Dictionary<string, string> locals)
    {
        foreach (var node in body)
        {
            RecordAddCall(node, moduleName, locals);
            foreach (var child in ChildBodies(node))
                ScanBody(child, moduleName, locals);
        }
    }

    // ── Add-call detection ────────────────────────────────────────────────────

    private void RecordAddCall(
        AstNode node,
        string moduleName,
        Dictionary<string, string> locals)
    {
        if (node is not CallStatementNode call) return;

        // Two representations depending on whether parens were used in VB6:
        //   col.Add item      → CallStatement { Target=MemberAccess(col,"Add"), Args=[item] }
        //   col.Add(item)     → CallStatement { Target=CallOrIndex(MemberAccess(col,"Add"),[item]), Args=[] }
        MemberAccessNode? addMa;
        IReadOnlyList<ArgumentNode> addArgs;

        if (call.Target is MemberAccessNode ma1 &&
            ma1.MemberName.Equals("Add", StringComparison.OrdinalIgnoreCase))
        {
            addMa = ma1;
            addArgs = call.Arguments;
        }
        else if (call.Target is CallOrIndexNode coi &&
                 coi.Target is MemberAccessNode ma2 &&
                 ma2.MemberName.Equals("Add", StringComparison.OrdinalIgnoreCase))
        {
            addMa = ma2;
            addArgs = coi.Arguments;
        }
        else return;

        // First non-missing argument is the item being added to the collection.
        var itemArg = addArgs.FirstOrDefault(a => !a.IsMissing);
        if (itemArg?.Value == null) return;

        string? elemType = InferType(itemArg.Value, moduleName, locals);
        if (elemType == null) return;

        var target = ResolveCollectionTarget(addMa.Object, moduleName, locals);
        if (target == null) return;

        var key = target.Value;
        if (!_hits.TryGetValue(key, out var set))
            _hits[key] = set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        set.Add(elemType);
    }

    // ── Type / target resolution ──────────────────────────────────────────────

    /// <summary>
    /// Resolves the collection variable being written to as (ownerModuleName, fieldName).
    /// Returns null when the target is not a recognisable Collection field reference.
    /// </summary>
    private (string mod, string field)? ResolveCollectionTarget(
        ExpressionNode target,
        string moduleName,
        Dictionary<string, string> locals)
    {
        switch (target)
        {
            // colField.Add ... — direct module-level field in the current class
            case IdentifierNode id:
                if (_moduleFields.TryGetValue(moduleName, out var mf) &&
                    mf.TryGetValue(id.Name, out var type) &&
                    type.Equals("Collection", StringComparison.OrdinalIgnoreCase))
                    return (moduleName, id.Name);
                break;

            // obj.colField.Add ... — field belonging to another class
            case MemberAccessNode ma:
            {
                string? ownerType = InferType(ma.Object, moduleName, locals);
                if (ownerType != null &&
                    _moduleFields.TryGetValue(ownerType, out var ownerFields) &&
                    ownerFields.TryGetValue(ma.MemberName, out var fieldType) &&
                    fieldType.Equals("Collection", StringComparison.OrdinalIgnoreCase))
                    return (ownerType, ma.MemberName);
                break;
            }
        }

        return null;
    }

    /// <summary>
    /// Best-effort type name for an expression.
    /// Returns null when the type cannot be statically determined.
    /// </summary>
    private string? InferType(
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
                if (_moduleFields.TryGetValue(moduleName, out var mf) &&
                    mf.TryGetValue(id.Name, out var ft)) return ft;
                return null;

            case MeNode:
                return moduleName;

            case MemberAccessNode ma:
            {
                // One level of member access: obj.TypedField
                string? ownerType = InferType(ma.Object, moduleName, locals);
                if (ownerType != null &&
                    _moduleFields.TryGetValue(ownerType, out var ownerFields) &&
                    ownerFields.TryGetValue(ma.MemberName, out var memberType))
                    return memberType;
                return null;
            }

            default:
                return null;
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
            case WithNode     w: yield return w.Body; break;
        }
    }
}
