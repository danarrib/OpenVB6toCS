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
                ReturnType = NormalizeType(f.ReturnType),
                Body       = TransformBody(f.Body, f.Name, f.Line),
            },
            CsPropertyNode p => p with
            {
                Type          = NormalizeType(p.Type),
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
            TryBody   = tc.TryBody.Select(TransformStatement).ToList(),
            CatchBody = tc.CatchBody.Select(TransformStatement).ToList(),
        },

        _ => node,
    };

    // ── Error handling transformation ───────────────────────────────────────

    private IReadOnlyList<AstNode> TransformErrorHandling(
        IReadOnlyList<AstNode> body, string context, int contextLine)
    {
        // Locate all "On Error GoTo label" nodes in the direct statement list.
        int errorNodeIdx = -1;
        int goToCount = 0;
        string? label = null;
        int errorLine = contextLine, errorCol = 1;

        for (int i = 0; i < body.Count; i++)
        {
            if (body[i] is OnErrorNode o && o.Kind == OnErrorKind.GoTo && o.LabelName != null)
            {
                goToCount++;
                if (goToCount == 1)
                {
                    errorNodeIdx = i;
                    label        = o.LabelName;
                    errorLine    = o.Line;
                    errorCol     = o.Column;
                }
            }
        }

        if (goToCount == 0)
            return body; // nothing to transform

        if (goToCount > 1)
        {
            Warn(context,
                $"{goToCount} 'On Error GoTo' handlers — structural error-handling transform skipped; manual review required.",
                contextLine, 1);
            return body;
        }

        // Find the matching label node.
        int labelIdx = -1;
        for (int i = 0; i < body.Count; i++)
        {
            if (body[i] is LabelNode ln &&
                ln.Name.Equals(label, StringComparison.OrdinalIgnoreCase))
            {
                labelIdx = i;
                break;
            }
        }

        if (labelIdx < 0)
        {
            Warn(context,
                $"label '{label}' referenced by 'On Error GoTo' not found in procedure body — transform skipped.",
                errorLine, errorCol);
            return body;
        }

        // Split the body:
        //   preamble  = [0 .. errorNodeIdx)
        //   tryBody   = (errorNodeIdx .. labelIdx)
        //   catchBody = (labelIdx .. end)
        var preamble  = body.Take(errorNodeIdx).ToList();
        var tryBody   = body.Skip(errorNodeIdx + 1).Take(labelIdx - errorNodeIdx - 1).ToList();
        var catchBody = body.Skip(labelIdx + 1).ToList();

        // Warn if there's no Exit Sub/Function before the label — the handler is reachable
        // without an error in VB6 (fall-through), which differs from C# try/catch semantics.
        bool hadExit = tryBody.OfType<ExitNode>()
            .Any(e => e.What is "Sub" or "Function" or "Property");

        if (!hadExit)
        {
            Warn(context,
                $"error handler '{label}' may be reachable without an error (no Exit before label); review generated try/catch.",
                body[labelIdx].Line, 1);
        }

        // Remove trailing Exit Sub/Function from tryBody (it just guards fall-through).
        tryBody = StripTrailingExit(tryBody);

        // Remove trailing Resume/Resume Next from catchBody.
        catchBody = StripTrailingResume(catchBody, context);

        var tryCatch = new TryCatchNode(
            body[errorNodeIdx].Line, body[errorNodeIdx].Column,
            tryBody, "_ex", catchBody);

        return [.. preamble, tryCatch];
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

    private List<AstNode> StripTrailingResume(List<AstNode> body, string context)
    {
        int i = body.Count - 1;
        while (i >= 0 && body[i] is CommentNode) i--;

        if (i >= 0 && body[i] is ResumeNode r)
        {
            if (r.IsNext)
            {
                // Resume Next: execution continues after the catch block — correct.
                var result = new List<AstNode>(body);
                result.RemoveAt(i);
                return result;
            }

            // Resume / Resume label: no direct C# equivalent; leave in body + warn.
            var resumeDesc = r.LabelName != null ? $"Resume {r.LabelName}" : "Resume";
            Warn(context,
                $"'{resumeDesc}' in error handler has no C# equivalent; manual review required.",
                r.Line, r.Column);
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
