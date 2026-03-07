using VB6toCS.Core.Parsing.Nodes;

namespace VB6toCS.Core.Analysis;

/// <summary>
/// Stage 3 — Semantic analysis.
///
/// Performs three passes over a parsed ModuleNode:
///   Pass 1 — Build the module-level symbol table from all declarations.
///   Pass 2 — Walk procedure bodies: resolve CallOrIndexNode (call vs. array
///             index), detect function return assignments, normalize identifier
///             casing to match the declaration site.
///   Pass 3 — Group sibling Property Get/Let/Set nodes into a single
///             CsPropertyNode, which maps naturally to a C# property.
///
/// Returns a new ModuleNode; the input is not modified.
/// Diagnostics (warnings for unresolvable names, etc.) are available via
/// <see cref="Diagnostics"/> after the call.
/// </summary>
public sealed class Analyser
{
    private readonly SymbolTable _symbols = new();
    private readonly List<AnalysisDiagnostic> _diagnostics = [];

    public IReadOnlyList<AnalysisDiagnostic> Diagnostics => _diagnostics;

    public ModuleNode Analyse(ModuleNode module)
    {
        // Pass 1: collect module-level declarations
        BuildModuleSymbols(module.Members);

        // Pass 2: transform procedure bodies
        var transformed = TransformMembers(module.Members);

        // Pass 3: group Property Get/Let/Set into CsPropertyNode
        var grouped = GroupProperties(transformed);

        return module with { Members = grouped };
    }

    // ── Pass 1: build module symbol table ───────────────────────────────────

    private void BuildModuleSymbols(IReadOnlyList<AstNode> members)
    {
        foreach (var node in members)
        {
            switch (node)
            {
                case FieldNode f:
                    foreach (var d in f.Declarators)
                        _symbols.AddModule(new Symbol(d.Name, SymbolKind.Field,
                            d.TypeRef, d.TypeRef?.IsArray ?? false));
                    break;

                case ConstDeclarationNode c:
                    foreach (var d in c.Declarators)
                        _symbols.AddModule(new Symbol(d.Name, SymbolKind.Constant,
                            d.TypeRef, false));
                    break;

                case EnumNode e:
                    _symbols.AddModule(new Symbol(e.Name, SymbolKind.Enum, null, false));
                    foreach (var m in e.Members)
                        _symbols.AddModule(new Symbol(m.Name, SymbolKind.EnumMember, null, false));
                    break;

                case UdtNode u:
                    _symbols.AddModule(new Symbol(u.Name, SymbolKind.Udt, null, false));
                    break;

                case SubNode s:
                    _symbols.AddModule(new Symbol(s.Name, SymbolKind.Sub, null, false));
                    break;

                case FunctionNode f:
                    _symbols.AddModule(new Symbol(f.Name, SymbolKind.Function,
                        f.ReturnType, f.ReturnType?.IsArray ?? false));
                    break;

                case PropertyNode p:
                    // All Get/Let/Set share the same name; only register once.
                    if (_symbols.Lookup(p.Name) == null)
                    {
                        var type = p.ReturnType ?? ValueParamType(p);
                        _symbols.AddModule(new Symbol(p.Name, SymbolKind.Property, type, false));
                    }
                    break;

                case DeclareNode d:
                    _symbols.AddModule(new Symbol(d.Name,
                        d.IsSub ? SymbolKind.Sub : SymbolKind.Function,
                        d.ReturnType, false));
                    break;
            }
        }
    }

    /// <summary>
    /// For Property Let/Set, the value being assigned is the last parameter.
    /// Returns its type, or null if there are no parameters.
    /// </summary>
    private static TypeRefNode? ValueParamType(PropertyNode p) =>
        p.Parameters.Count > 0 ? p.Parameters[^1].TypeRef : null;

    // ── Pass 2: transform members ───────────────────────────────────────────

    private IReadOnlyList<AstNode> TransformMembers(IReadOnlyList<AstNode> members)
    {
        var result = new List<AstNode>(members.Count);
        foreach (var node in members)
        {
            result.Add(node switch
            {
                SubNode s      => TransformSub(s),
                FunctionNode f => TransformFunction(f),
                PropertyNode p => TransformProperty(p),
                _              => node   // FieldNode, EnumNode, etc. — no bodies to walk
            });
        }
        return result;
    }

    private SubNode TransformSub(SubNode s)
    {
        _symbols.EnterProcedure(s.Name, returnName: null);
        AddParametersToScope(s.Parameters);
        var body = TransformBody(s.Body);
        _symbols.ExitProcedure();
        return s with { Body = body };
    }

    private FunctionNode TransformFunction(FunctionNode f)
    {
        // Function name is the return-value pseudo-variable.
        _symbols.EnterProcedure(f.Name, returnName: f.Name);
        AddParametersToScope(f.Parameters);
        var body = TransformBody(f.Body);
        _symbols.ExitProcedure();
        return f with { Body = body };
    }

    private PropertyNode TransformProperty(PropertyNode p)
    {
        // Property Get has a return value (assigned via PropertyName = expr).
        string? retName = p.Kind == PropertyKind.Get ? p.Name : null;
        _symbols.EnterProcedure(p.Name, returnName: retName);
        AddParametersToScope(p.Parameters);
        var body = TransformBody(p.Body);
        _symbols.ExitProcedure();
        return p with { Body = body };
    }

    private void AddParametersToScope(IReadOnlyList<ParameterNode> parameters)
    {
        foreach (var p in parameters)
            _symbols.AddLocal(new Symbol(p.Name, SymbolKind.Parameter,
                p.TypeRef, p.TypeRef?.IsArray ?? false));
    }

    // ── Statement transformation ────────────────────────────────────────────

    private IReadOnlyList<AstNode> TransformBody(IReadOnlyList<AstNode> body) =>
        body.Select(TransformStatement).ToList();

    private AstNode TransformStatement(AstNode node)
    {
        return node switch
        {
            AssignmentNode a    => TransformAssignment(a),

            CallStatementNode c => c with
            {
                Target    = TransformExpression(c.Target),
                Arguments = TransformArgs(c.Arguments),
            },

            LocalDimNode d => RegisterLocals(d),

            ReDimNode r => r with
            {
                Declarators = r.Declarators.Select(RegisterArrayDeclarator).ToList(),
            },

            IfNode i => i with
            {
                Condition     = TransformExpression(i.Condition),
                ThenBody      = TransformBody(i.ThenBody),
                ElseIfClauses = i.ElseIfClauses
                    .Select(ei => ei with
                    {
                        Condition = TransformExpression(ei.Condition),
                        Body      = TransformBody(ei.Body),
                    })
                    .ToList(),
                ElseBody = i.ElseBody != null ? TransformBody(i.ElseBody) : null,
            },

            SingleLineIfNode si => si with
            {
                Condition     = TransformExpression(si.Condition),
                ThenStatement = TransformStatement(si.ThenStatement),
                ElseStatement = si.ElseStatement != null
                    ? TransformStatement(si.ElseStatement) : null,
            },

            SelectCaseNode s => s with
            {
                TestExpression = TransformExpression(s.TestExpression),
                Cases = s.Cases
                    .Select(c => c with
                    {
                        Patterns = c.Patterns.Select(TransformCasePattern).ToList(),
                        Body     = TransformBody(c.Body),
                    })
                    .ToList(),
            },

            ForNextNode f => f with
            {
                Start = TransformExpression(f.Start),
                End   = TransformExpression(f.End),
                Step  = f.Step != null ? TransformExpression(f.Step) : null,
                Body  = TransformBody(f.Body),
            },

            ForEachNode f => f with
            {
                Collection = TransformExpression(f.Collection),
                Body       = TransformBody(f.Body),
            },

            WhileNode w => w with
            {
                Condition = TransformExpression(w.Condition),
                Body      = TransformBody(w.Body),
            },

            DoLoopNode d => d with
            {
                Condition = d.Condition != null ? TransformExpression(d.Condition) : null,
                Body      = TransformBody(d.Body),
            },

            WithNode w => w with
            {
                Object = TransformExpression(w.Object),
                Body   = TransformBody(w.Body),
            },

            ErrorStatementNode e => e with
            {
                ErrorNumber = TransformExpression(e.ErrorNumber),
            },

            // Leaf statements — nothing to transform
            GoToNode or GoSubNode or ReturnNode or ExitNode or LabelNode
                or CommentNode or ResumeNode or OnErrorNode or EndStatementNode => node,

            _ => node,
        };
    }

    private AstNode TransformAssignment(AssignmentNode a)
    {
        var target = TransformExpression(a.Target);
        var value  = TransformExpression(a.Value);

        // Detect return-value assignment:
        //   FunctionName = expr       inside a Function body
        //   PropertyName = expr       inside a Property Get body
        // Both Set and non-Set variants are detected (e.g. Set GetObj = New Foo).
        if (_symbols.CurrentReturnName != null &&
            target is IdentifierNode id &&
            id.Name.Equals(_symbols.CurrentReturnName, StringComparison.OrdinalIgnoreCase))
        {
            return new FunctionReturnNode(a.Line, a.Column, value);
        }

        return a with { Target = target, Value = value };
    }

    private LocalDimNode RegisterLocals(LocalDimNode d)
    {
        foreach (var decl in d.Declarators)
            _symbols.AddLocal(new Symbol(decl.Name, SymbolKind.LocalVar,
                decl.TypeRef, decl.TypeRef?.IsArray ?? false));
        return d;
    }

    private VariableDeclaratorNode RegisterArrayDeclarator(VariableDeclaratorNode d)
    {
        // ReDim may introduce a new name if not already in scope (unusual but valid).
        if (_symbols.Lookup(d.Name) == null)
            _symbols.AddLocal(new Symbol(d.Name, SymbolKind.LocalVar, d.TypeRef, true));
        return d;
    }

    private CasePatternNode TransformCasePattern(CasePatternNode p) => p switch
    {
        CaseValuePattern v => v with { Value = TransformExpression(v.Value) },
        CaseRangePattern r => r with
        {
            Low  = TransformExpression(r.Low),
            High = TransformExpression(r.High),
        },
        CaseIsPattern i => i with { Value = TransformExpression(i.Value) },
        _               => p,
    };

    // ── Expression transformation ───────────────────────────────────────────

    private ExpressionNode TransformExpression(ExpressionNode e)
    {
        return e switch
        {
            CallOrIndexNode c      => ResolveCallOrIndex(c),

            MemberAccessNode m     => m with { Object = TransformExpression(m.Object) },
            BangAccessNode b       => b with { Object = TransformExpression(b.Object) },

            BinaryExpressionNode b => b with
            {
                Left  = TransformExpression(b.Left),
                Right = TransformExpression(b.Right),
            },

            UnaryExpressionNode u  => u with { Operand = TransformExpression(u.Operand) },
            TypeOfIsNode t         => t with { Operand = TransformExpression(t.Operand) },

            IdentifierNode id      => NormalizeIdentifier(id),

            // Literals, MeNode, NothingNode, NewObjectNode, WithMemberAccessNode
            _ => e,
        };
    }

    private ExpressionNode ResolveCallOrIndex(CallOrIndexNode c)
    {
        var target = TransformExpression(c.Target);
        var args   = TransformArgs(c.Arguments);

        // Resolution is only possible when the target is a bare identifier that
        // we can look up in the symbol table.  Member-access chains (obj.Prop(i))
        // require type information and are left as CallOrIndexNode for Stage 4/5.
        if (target is IdentifierNode id)
        {
            var sym = _symbols.Lookup(id.Name);
            if (sym != null && sym.IsArray)
                return new IndexNode(c.Line, c.Column, target, args);

            // Sub / Function / DeclareNode / Property → it's a call; keep as CallOrIndexNode
            // Unresolved (built-in functions, COM methods) → also keep as CallOrIndexNode
        }

        return c with { Target = target, Arguments = args };
    }

    private IdentifierNode NormalizeIdentifier(IdentifierNode id)
    {
        var sym = _symbols.Lookup(id.Name);
        if (sym != null && !string.Equals(sym.Name, id.Name, StringComparison.Ordinal))
            return id with { Name = sym.Name };
        return id;
    }

    private IReadOnlyList<ArgumentNode> TransformArgs(IReadOnlyList<ArgumentNode> args) =>
        args.Select(a => a is { IsMissing: false, Value: not null }
            ? a with { Value = TransformExpression(a.Value) }
            : a).ToList();

    // ── Pass 3: group Property Get/Let/Set ──────────────────────────────────

    private static IReadOnlyList<AstNode> GroupProperties(IReadOnlyList<AstNode> members)
    {
        var result  = new List<AstNode>(members.Count);
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        for (int i = 0; i < members.Count; i++)
        {
            if (members[i] is not PropertyNode first)
            {
                result.Add(members[i]);
                continue;
            }

            if (visited.Contains(first.Name))
                continue; // already emitted as part of a CsPropertyNode

            visited.Add(first.Name);

            // Collect all property accessors with this name (may be non-consecutive).
            PropertyNode? getProp = null, letProp = null, setProp = null;
            for (int j = i; j < members.Count; j++)
            {
                if (members[j] is PropertyNode p &&
                    p.Name.Equals(first.Name, StringComparison.OrdinalIgnoreCase))
                {
                    switch (p.Kind)
                    {
                        case PropertyKind.Get: getProp = p; break;
                        case PropertyKind.Let: letProp = p; break;
                        case PropertyKind.Set: setProp = p; break;
                    }
                }
            }

            // Type: prefer Get return type; fall back to the value parameter of Let/Set.
            TypeRefNode? type = getProp?.ReturnType
                ?? ValueParamType(letProp ?? setProp!);

            result.Add(new CsPropertyNode(
                first.Line, first.Column,
                first.Access,
                first.IsStatic,
                first.Name,
                type,
                getProp?.Parameters ?? [],
                getProp?.Body,
                letProp?.Parameters ?? [],
                letProp?.Body,
                setProp?.Parameters ?? [],
                setProp?.Body));
        }

        return result;
    }
}
