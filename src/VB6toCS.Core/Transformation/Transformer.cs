using VB6toCS.Core.Analysis;
using VB6toCS.Core.Parsing.Nodes;

namespace VB6toCS.Core.Transformation;

/// <summary>
/// Stage 4 — IR Transformation.
///
/// Two passes over a Stage-3 ModuleNode:
///   Pass 1 — Type normalization: maps VB6 primitive type names (String, Long, Boolean, …)
///             to their C# equivalents in every TypeRefNode in the AST. Known COM type names
///             (Scripting.Dictionary, etc.) are also remapped.
///   Pass 2 — Error handling: transforms the VB6 pattern
///               On Error GoTo label / [guarded body] / Exit / label: / [handler body]
///             into a TryCatchNode in each procedure body.
///
/// Returns a new ModuleNode; the input is not modified.
/// Diagnostics are available via <see cref="Diagnostics"/> after the call.
/// </summary>
public sealed class Transformer
{
    private readonly List<TransformDiagnostic> _diagnostics = [];
    public IReadOnlyList<TransformDiagnostic> Diagnostics => _diagnostics;

    // Inferred Collection element types from the cross-module CollectionTypeInferrer pass.
    // Key: (moduleName, fieldName).  Value: the single inferred element type.
    private readonly IReadOnlyDictionary<(string, string), string>? _collectionTypes;

    // Inferred Collection kind (Dictionary / List / Collection) per field.
    // Only fields with unambiguous Add-call patterns appear in this map.
    private readonly IReadOnlyDictionary<(string, string), CollectionKind>? _collectionKinds;

    private string _currentContext = "<module>";
    private string _currentModuleName = "";

    public Transformer(
        IReadOnlyDictionary<(string, string), string>? collectionTypes = null,
        IReadOnlyDictionary<(string, string), CollectionKind>? collectionKinds = null)
    {
        _collectionTypes = collectionTypes;
        _collectionKinds = collectionKinds;
    }

    // VB6 primitive type names → C# equivalents
    private static readonly Dictionary<string, string> TypeMap =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["String"]     = "string",
            ["Integer"]    = "int",
            ["Long"]       = "int",
            ["Single"]     = "float",
            ["Double"]     = "double",
            ["Boolean"]    = "bool",
            ["Byte"]       = "byte",
            ["Date"]       = "DateTime",
            ["Currency"]   = "decimal",
            ["Variant"]    = "object",
            ["Object"]     = "object",
            // Scripting.Dictionary used unqualified → generic Dictionary
            ["Dictionary"] = "Dictionary<string, object>",
        };

    public ModuleNode Transform(ModuleNode module)
    {
        _currentModuleName = module.Name;
        var members = module.Members.Select(TransformMember).ToList();
        return module with { Members = members };
    }

    // ── Member-level transformation ─────────────────────────────────────────

    private AstNode TransformMember(AstNode node)
    {
        // Set context so NormalizeType diagnostics carry a meaningful member name.
        _currentContext = node switch
        {
            SubNode s          => s.Name,
            FunctionNode f     => f.Name,
            CsPropertyNode p   => p.Name,
            FieldNode          => "<field>",
            ConstDeclarationNode => "<const>",
            UdtNode u          => u.Name,
            DeclareNode d      => d.Name,
            _                  => "<module>",
        };

        return node switch
        {
            SubNode s => s with
            {
                Parameters = NormalizeParams(s.Parameters),
                Body       = TransformBody(s.Body, s.Name, s.Line),
            },
            FunctionNode f => f with
            {
                Parameters = NormalizeParams(f.Parameters),
                ReturnType = NormalizeReturnType(f.ReturnType, f.Name),
                Body       = TransformBody(f.Body, f.Name, f.Line),
            },
            CsPropertyNode p => p with
            {
                Type          = NormalizeReturnType(p.Type, p.Name),
                GetParameters = NormalizeParams(p.GetParameters),
                GetBody       = p.GetBody  != null ? TransformBody(p.GetBody,  p.Name + ".Get", p.Line) : null,
                LetParameters = NormalizeParams(p.LetParameters),
                LetBody       = p.LetBody  != null ? TransformBody(p.LetBody,  p.Name + ".Let", p.Line) : null,
                SetParameters = NormalizeParams(p.SetParameters),
                SetBody       = p.SetBody  != null ? TransformBody(p.SetBody,  p.Name + ".Set", p.Line) : null,
            },
            FieldNode f            => f with { Declarators = NormalizeFieldDeclarators(f.Declarators) },
            ConstDeclarationNode c => c with { Declarators = NormalizeDeclarators(c.Declarators) },
            UdtNode u              => u with
            {
                Fields = u.Fields.Select(f => f with { TypeRef = NormalizeType(f.TypeRef)! }).ToList(),
            },
            DeclareNode d => d with
            {
                Parameters = NormalizeParams(d.Parameters),
                ReturnType = NormalizeType(d.ReturnType),
            },
            _ => node,
        };
    }

    // ── Body-level transformation ───────────────────────────────────────────

    private IReadOnlyList<AstNode> TransformBody(
        IReadOnlyList<AstNode> body, string context, int contextLine)
    {
        // Error handling restructuring runs first on the flat statement list.
        var result = TransformErrorHandling(body, context, contextLine);
        // Then recurse into each statement for nested bodies and type normalization.
        return result.Select(TransformStatement).ToList();
    }

    private AstNode TransformStatement(AstNode node) => node switch
    {
        LocalDimNode d => d with { Declarators = NormalizeDeclarators(d.Declarators) },
        ReDimNode    r => r with { Declarators = NormalizeDeclarators(r.Declarators) },

        IfNode i => i with
        {
            ThenBody      = i.ThenBody.Select(TransformStatement).ToList(),
            ElseIfClauses = i.ElseIfClauses
                .Select(ei => ei with { Body = ei.Body.Select(TransformStatement).ToList() })
                .ToList(),
            ElseBody = i.ElseBody?.Select(TransformStatement).ToList(),
        },

        SingleLineIfNode si => si with
        {
            ThenStatement = TransformStatement(si.ThenStatement),
            ElseStatement = si.ElseStatement != null
                ? TransformStatement(si.ElseStatement) : null,
        },

        SelectCaseNode s => s with
        {
            Cases = s.Cases
                .Select(c => c with { Body = c.Body.Select(TransformStatement).ToList() })
                .ToList(),
        },

        ForNextNode  f  => f  with { Body = f.Body.Select(TransformStatement).ToList() },
        ForEachNode  fe => fe with { Body = fe.Body.Select(TransformStatement).ToList() },
        WhileNode    w  => w  with { Body = w.Body.Select(TransformStatement).ToList() },
        DoLoopNode   dl => dl with { Body = dl.Body.Select(TransformStatement).ToList() },
        WithNode     w  => w  with { Body = w.Body.Select(TransformStatement).ToList() },

        TryCatchNode tc => tc with
        {
            TryBody     = tc.TryBody.Select(TransformStatement).ToList(),
            CatchBody   = tc.CatchBody.Select(TransformStatement).ToList(),
            FinallyBody = tc.FinallyBody?.Select(TransformStatement).ToList(),
        },

        _ => node,
    };

    // ── Error handling transformation ───────────────────────────────────────

    private IReadOnlyList<AstNode> TransformErrorHandling(
        IReadOnlyList<AstNode> body, string context, int contextLine)
    {
        // Try to detect the structured try/catch/finally pattern. If it matches all
        // validation checks, return a clean TryCatchNode with no residual comments or
        // labels. If anything fails, return the body untouched — the code generator
        // will emit all error-related nodes as comments for manual review.
        return TryDetectStructuredPattern(body) ?? body;
    }

    /// <summary>
    /// Attempts to detect the canonical VB6 try/catch/finally pattern:
    /// <code>
    ///   On Error GoTo ErrLabel
    ///   [try body]
    ///   CleanupLabel:
    ///   [finally body]
    ///   Exit Sub / Function / Property
    ///   ErrLabel:
    ///   [catch body]
    ///   GoTo CleanupLabel
    /// </code>
    /// Returns a restructured body containing a clean <see cref="TryCatchNode"/> on success,
    /// or <c>null</c> if any validation check fails (caller falls back to comment-out).
    /// </summary>
    private static IReadOnlyList<AstNode>? TryDetectStructuredPattern(IReadOnlyList<AstNode> body)
    {
        // ── 1. Exactly one "On Error GoTo <label>" in the flat top-level body ─────────
        int onErrorIdx = -1;
        string? errLabel = null;
        foreach (var (node, i) in body.Select((n, i) => (n, i)))
        {
            if (node is OnErrorNode o && o.Kind == OnErrorKind.GoTo && o.LabelName != null)
            {
                if (onErrorIdx >= 0) return null; // more than one
                onErrorIdx = i;
                errLabel = o.LabelName;
            }
        }
        if (onErrorIdx < 0) return null;

        // ── 2. No Resume anywhere in the entire body (including nested) ──────────────
        if (FindAllDeep<ResumeNode>(body).Any()) return null;

        // ── 3. Find errLabel in the flat body (must come after On Error GoTo) ────────
        int errIdx = FindLabelIdx(body, errLabel!);
        if (errIdx < 0 || errIdx <= onErrorIdx) return null;

        // ── 4. Catch section validation ───────────────────────────────────────────────
        var catchSection = body.Skip(errIdx + 1).ToList();

        // 4a. Last meaningful statement must be a GoTo <cleanupLabel>
        int lastMeaningful = LastMeaningfulIdx(catchSection);
        if (lastMeaningful < 0) return null;
        if (catchSection[lastMeaningful] is not GoToNode finalGoTo) return null;
        string cleanupLabel = finalGoTo.Label;

        // 4b. No GoTo in catch section other than that final one (at any nesting depth)
        if (FindAllDeep<GoToNode>(catchSection).Count() != 1) return null;

        // 4c. No nested On Error in catch section
        if (FindAllDeep<OnErrorNode>(catchSection).Any()) return null;

        // ── 5. Find cleanupLabel (must appear before errLabel in the flat body) ───────
        int cleanupIdx = FindLabelIdx(body, cleanupLabel);
        if (cleanupIdx < 0 || cleanupIdx >= errIdx) return null;

        // ── 6. Finally section validation (cleanupLabel+1 .. errLabel-1) ─────────────
        var finallySection = body.Skip(cleanupIdx + 1).Take(errIdx - cleanupIdx - 1).ToList();

        // 6a. Must contain an Exit Sub / Function / Property
        if (!finallySection.OfType<ExitNode>()
                .Any(e => e.What is "Sub" or "Function" or "Property"))
            return null;

        // 6b. No GoTo or On Error in finally section (at any nesting depth)
        if (FindAllDeep<GoToNode>(finallySection).Any()) return null;
        if (FindAllDeep<OnErrorNode>(finallySection).Any()) return null;

        // ── 7. Try section validation (On Error GoTo+1 .. cleanupLabel-1) ────────────
        var trySection = body.Skip(onErrorIdx + 1).Take(cleanupIdx - onErrorIdx - 1).ToList();

        // 7a. No GoTo that escapes to errLabel or cleanupLabel (at any nesting depth)
        bool hasEscapeGoTo = FindAllDeep<GoToNode>(trySection).Any(g =>
            g.Label.Equals(errLabel,     StringComparison.OrdinalIgnoreCase) ||
            g.Label.Equals(cleanupLabel, StringComparison.OrdinalIgnoreCase));
        if (hasEscapeGoTo) return null;

        // 7b. No nested On Error in try section
        if (FindAllDeep<OnErrorNode>(trySection).Any()) return null;

        // ── All checks passed — build the clean try/catch/finally ─────────────────────
        // Finally body: strip the trailing Exit (implicit in try/finally control flow)
        var finallyBody = StripTrailingExit(finallySection);
        // Catch body: strip the trailing GoTo (now represented by the finally clause)
        var catchBody = catchSection.Take(lastMeaningful).ToList();
        // Preamble: anything before On Error GoTo (typically variable declarations)
        var preamble = body.Take(onErrorIdx).ToList();

        var tryCatch = new TryCatchNode(
            body[onErrorIdx].Line, body[onErrorIdx].Column,
            trySection, "_ex", catchBody, finallyBody);

        return [.. preamble, tryCatch];
    }

    // ── Pattern-detection helpers ────────────────────────────────────────────

    private static int FindLabelIdx(IReadOnlyList<AstNode> body, string labelName)
    {
        for (int i = 0; i < body.Count; i++)
            if (body[i] is LabelNode l &&
                l.Name.Equals(labelName, StringComparison.OrdinalIgnoreCase))
                return i;
        return -1;
    }

    private static int LastMeaningfulIdx(IReadOnlyList<AstNode> list)
    {
        for (int i = list.Count - 1; i >= 0; i--)
            if (list[i] is not CommentNode) return i;
        return -1;
    }

    /// <summary>
    /// Recursively yields all nodes of type <typeparamref name="T"/> anywhere in the body,
    /// including inside nested statement blocks (If, For, With, etc.).
    /// </summary>
    private static IEnumerable<T> FindAllDeep<T>(IReadOnlyList<AstNode> body) where T : AstNode
    {
        foreach (var node in body)
        {
            if (node is T t) yield return t;
            foreach (var child in ChildBodies(node))
                foreach (var found in FindAllDeep<T>(child))
                    yield return found;
        }
    }

    private static IEnumerable<IReadOnlyList<AstNode>> ChildBodies(AstNode node)
    {
        switch (node)
        {
            case IfNode i:
                yield return i.ThenBody;
                foreach (var ei in i.ElseIfClauses) yield return ei.Body;
                if (i.ElseBody != null) yield return i.ElseBody;
                break;
            case SelectCaseNode s:
                foreach (var c in s.Cases) yield return c.Body;
                break;
            case ForNextNode f:  yield return f.Body; break;
            case ForEachNode fe: yield return fe.Body; break;
            case WhileNode w:    yield return w.Body;  break;
            case DoLoopNode d:   yield return d.Body;  break;
            case WithNode w:     yield return w.Body;  break;
            case TryCatchNode tc:
                yield return tc.TryBody;
                yield return tc.CatchBody;
                if (tc.FinallyBody != null) yield return tc.FinallyBody;
                break;
        }
    }

    private static List<AstNode> StripTrailingExit(List<AstNode> body)
    {
        int i = body.Count - 1;
        while (i >= 0 && body[i] is CommentNode) i--;

        if (i >= 0 && body[i] is ExitNode e &&
            e.What is "Sub" or "Function" or "Property")
        {
            var result = new List<AstNode>(body);
            result.RemoveAt(i);
            return result;
        }

        return body;
    }

    // ── Type normalization helpers ───────────────────────────────────────────

    private TypeRefNode? NormalizeType(TypeRefNode? t)
    {
        if (t == null) return null;
        return TypeMap.TryGetValue(t.TypeName, out var csName)
            ? t with { TypeName = csName }
            : t; // COM types, user-defined types, etc. — unchanged
    }

    /// <summary>
    /// Like NormalizeType, but also resolves Collection return types to Dictionary/List.
    /// Used for function and property return types (which don't have a field name for
    /// element-type inference, so they always fall back to Dictionary&lt;string, object&gt;).
    /// </summary>
    private TypeRefNode? NormalizeReturnType(TypeRefNode? t, string procName)
    {
        if (t == null) return null;
        if (t.TypeName.Equals("Collection", StringComparison.OrdinalIgnoreCase))
            return ResolveCollectionType(t, procName, isField: false);
        return NormalizeType(t);
    }

    /// <summary>
    /// Normalizes declarators for module-level fields.
    /// For Collection fields: uses the inferred element type when available;
    /// falls back to Collection&lt;object&gt; with a warning.
    /// </summary>
    private IReadOnlyList<VariableDeclaratorNode> NormalizeFieldDeclarators(
        IReadOnlyList<VariableDeclaratorNode> declarators)
    {
        if (declarators.Count == 0) return declarators;
        return declarators.Select(d =>
        {
            if (d.TypeRef?.TypeName.Equals("Collection", StringComparison.OrdinalIgnoreCase) == true)
                return d with { TypeRef = ResolveCollectionType(d.TypeRef, d.Name, isField: true) };
            return d with { TypeRef = NormalizeType(d.TypeRef) };
        }).ToList();
    }

    /// <summary>
    /// Normalizes declarators for local variables and ReDim statements.
    /// Collection locals always become Collection&lt;object&gt; with a warning
    /// (local variable inference is not currently implemented).
    /// </summary>
    private IReadOnlyList<VariableDeclaratorNode> NormalizeDeclarators(
        IReadOnlyList<VariableDeclaratorNode> declarators)
    {
        if (declarators.Count == 0) return declarators;
        return declarators.Select(d =>
        {
            if (d.TypeRef?.TypeName.Equals("Collection", StringComparison.OrdinalIgnoreCase) == true)
                return d with { TypeRef = ResolveCollectionType(d.TypeRef, d.Name, isField: false) };
            return d with { TypeRef = NormalizeType(d.TypeRef) };
        }).ToList();
    }

    private IReadOnlyList<ParameterNode> NormalizeParams(
        IReadOnlyList<ParameterNode> parameters)
    {
        if (parameters.Count == 0) return parameters;
        return parameters.Select(p =>
        {
            if (p.TypeRef?.TypeName.Equals("Collection", StringComparison.OrdinalIgnoreCase) == true)
                return p with { TypeRef = ResolveCollectionType(p.TypeRef, p.Name, isField: false) };
            return p with { TypeRef = NormalizeType(p.TypeRef) };
        }).ToList();
    }

    /// <summary>
    /// Resolves a Collection TypeRefNode to the most appropriate C# collection type.
    /// For fields, consults the inferred kind and element type maps:
    /// <list type="bullet">
    ///   <item><see cref="CollectionKind.Dictionary"/> → <c>Dictionary&lt;string, T&gt;</c></item>
    ///   <item><see cref="CollectionKind.List"/>       → <c>List&lt;T&gt;</c></item>
    ///   <item>otherwise                               → <c>Collection&lt;T&gt;</c> with REVIEW warning</item>
    /// </list>
    /// Non-field (locals, parameters) always become <c>Collection&lt;object&gt;</c> for now.
    /// </summary>
    private TypeRefNode ResolveCollectionType(TypeRefNode t, string name, bool isField)
    {
        string? elemType = null;
        CollectionKind kind = CollectionKind.Dictionary; // default: Dictionary unless inferred as List

        if (isField && _collectionTypes != null)
            _collectionTypes.TryGetValue((_currentModuleName, name), out elemType);
        if (isField && _collectionKinds != null)
            _collectionKinds.TryGetValue((_currentModuleName, name), out kind);
        // Non-fields (locals, parameters): no Add-call data available → keep default Dictionary

        string elem = elemType ?? "object";

        if (kind == CollectionKind.List)
            return t with { TypeName = $"List<{elem}>" };

        // Dictionary (explicit or default)
        if (elemType == null)
            Warn(_currentContext,
                $"'{name}' is a Collection whose element type could not be inferred; " +
                "translated to Dictionary<string, object>. Review key and value types.",
                t.Line, t.Column);
        return t with { TypeName = $"Dictionary<string, {elem}>" };
    }

    // ── Diagnostic helpers ───────────────────────────────────────────────────

    private void Warn(string context, string message, int line, int col) =>
        _diagnostics.Add(new TransformDiagnostic(
            DiagnosticSeverity.Warning,
            $"'{context}': {message}",
            line, col));
}
