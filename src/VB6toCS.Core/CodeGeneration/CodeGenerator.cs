using VB6toCS.Core.Lexing;
using VB6toCS.Core.Parsing.Nodes;

namespace VB6toCS.Core.CodeGeneration;

/// <summary>
/// Stage 5 — C# Code Generation.
/// Walks the Stage-4 AST and emits a complete C# source file as a string.
/// </summary>
public sealed class CodeGenerator
{
    private readonly CodeWriter _w = new();
    private readonly bool _isStatic;    // true for .bas standard modules
    private readonly Stack<string> _withStack = new();  // tracks With object expressions
    private readonly HashSet<string> _requiredUsings = new();  // populated during generation

    // Cross-module enum member map: memberName → (moduleName, enumName).
    // Built by the caller from all Stage-3 ASTs; null in single-file diagnostic mode.
    private readonly IReadOnlyDictionary<string, (string ModuleName, string EnumName)>? _enumMemberMap;

    // Set of field/property/UDT-field names (any module) whose declared type is an enum.
    // Used to suppress the (int) cast when both sides of a comparison are enum-typed.
    private readonly IReadOnlySet<string>? _enumTypedFieldNames;

    private string _currentModuleName = "";

    private CodeGenerator(bool isStatic,
        IReadOnlyDictionary<string, (string ModuleName, string EnumName)>? enumMemberMap,
        IReadOnlySet<string>? enumTypedFieldNames)
    {
        _isStatic = isStatic;
        _enumMemberMap = enumMemberMap;
        _enumTypedFieldNames = enumTypedFieldNames;
    }

    public static string Generate(ModuleNode module, bool isStaticModule,
        IReadOnlyDictionary<string, (string ModuleName, string EnumName)>? enumMemberMap = null,
        IReadOnlySet<string>? enumTypedFieldNames = null)
    {
        var gen = new CodeGenerator(isStaticModule, enumMemberMap, enumTypedFieldNames);
        gen.GenerateModule(module);

        string classCode = gen._w.ToString();
        if (gen._requiredUsings.Count == 0)
            return classCode;

        var sb = new System.Text.StringBuilder();
        foreach (var ns in gen._requiredUsings.OrderBy(x => x))
            sb.AppendLine($"using {ns};");
        sb.AppendLine();
        sb.Append(classCode);
        return sb.ToString();
    }

    // ── Module ───────────────────────────────────────────────────────────────

    private void GenerateModule(ModuleNode module)
    {
        _currentModuleName = module.Name;
        string staticKw  = _isStatic ? "static " : "";
        string baseClause = module.Implements.Count > 0
            ? " : " + string.Join(", ", module.Implements)
            : "";

        _w.WriteLine($"public {staticKw}class {module.Name}{baseClause}");
        _w.OpenBrace();

        bool first = true;
        foreach (var member in module.Members)
        {
            if (member is CommentNode) { GenerateMember(member); continue; }
            if (!first) _w.WriteLine();
            GenerateMember(member);
            first = false;
        }

        _w.CloseBrace();
    }

    // ── Members ──────────────────────────────────────────────────────────────

    private void GenerateMember(AstNode node)
    {
        switch (node)
        {
            case CommentNode c:
                _w.WriteLine($"// {c.Text.TrimStart('\'').Trim()}");
                break;

            case FieldNode f:
                foreach (var d in f.Declarators)
                    _w.WriteLine($"{Access(f.Access)}{StaticMod()}{TypeStr(d.TypeRef, d.Name)} {d.Name};{ReviewComment(d.TypeRef)}");
                break;

            case ConstDeclarationNode c:
                foreach (var d in c.Declarators)
                    _w.WriteLine($"{Access(c.Access)}{StaticMod()}const {TypeStr(d.TypeRef, d.Name)} {d.Name} = {(d.DefaultValue != null ? Expr(d.DefaultValue) : "default")};{ReviewComment(d.TypeRef)}");
                break;

            case EnumNode e:
                _w.WriteLine($"{Access(e.Access)}enum {e.Name} : int");
                _w.OpenBrace();
                foreach (var m in e.Members)
                {
                    foreach (var tc in m.TrailingComments)
                        _w.WriteLine($"// {tc.Text.TrimStart('\'').Trim()}");
                    string val = m.Value != null ? $" = {Expr(m.Value)}" : "";
                    _w.WriteLine($"{m.Name}{val},");
                }
                _w.CloseBrace();
                break;

            case UdtNode u:
                _w.WriteLine($"{Access(u.Access)}struct {u.Name}");
                _w.OpenBrace();
                foreach (var f in u.Fields)
                    _w.WriteLine($"public {TypeStr(f.TypeRef, f.Name)} {f.Name};");
                _w.CloseBrace();
                break;

            case DeclareNode d:
                _w.WriteLine($"// TODO: P/Invoke — Declare {(d.IsSub ? "Sub" : "Function")} {d.Name} Lib \"{d.LibName}\"");
                break;

            case SubNode s:
                GenerateSub(s);
                break;

            case FunctionNode f:
                GenerateFunction(f);
                break;

            case CsPropertyNode p:
                GenerateProperty(p);
                break;
        }
    }

    // ── Methods ──────────────────────────────────────────────────────────────

    private void GenerateSub(SubNode s)
    {
        string staticKw = _isStatic || s.IsStatic ? "static " : "";
        _w.WriteLine($"{Access(s.Access)}{staticKw}void {s.Name}({Params(s.Parameters)})");
        _w.OpenBrace();
        GenerateBody(s.Body, returnType: null);
        _w.CloseBrace();
    }

    private void GenerateFunction(FunctionNode f)
    {
        string staticKw  = _isStatic || f.IsStatic ? "static " : "";
        string returnType = TypeStr(f.ReturnType, f.Name);
        _w.WriteLine($"{Access(f.Access)}{staticKw}{returnType} {f.Name}({Params(f.Parameters)})");
        _w.OpenBrace();
        GenerateBodyWithReturn(f.Body, returnType);
        _w.CloseBrace();
    }

    private void GenerateProperty(CsPropertyNode p)
    {
        string staticKw  = _isStatic || p.IsStatic ? "static " : "";
        string propType  = TypeStr(p.Type, p.Name);
        _w.WriteLine($"{Access(p.Access)}{staticKw}{propType} {p.Name}");
        _w.OpenBrace();

        if (p.GetBody != null)
        {
            _w.WriteLine("get");
            _w.OpenBrace();
            GenerateBodyWithReturn(p.GetBody, propType);
            _w.CloseBrace();
        }

        if (p.LetBody != null)
        {
            // value parameter is the last parameter
            string valueParam = p.LetParameters.Count > 0
                ? p.LetParameters[^1].Name : "value";
            _w.WriteLine($"set  // Let — parameter: {valueParam}");
            _w.OpenBrace();
            GenerateBody(p.LetBody, returnType: null);
            _w.CloseBrace();
        }

        if (p.SetBody != null)
        {
            string valueParam = p.SetParameters.Count > 0
                ? p.SetParameters[^1].Name : "value";
            _w.WriteLine($"set  // Set — parameter: {valueParam}");
            _w.OpenBrace();
            GenerateBody(p.SetBody, returnType: null);
            _w.CloseBrace();
        }

        _w.CloseBrace();
    }

    // ── Body generation ──────────────────────────────────────────────────────

    /// <summary>Generates a body for a Sub or property setter (no return value).</summary>
    private void GenerateBody(IReadOnlyList<AstNode> body, string? returnType)
    {
        foreach (var stmt in body)
            GenerateStatement(stmt, returnType, resultVar: null);
    }

    /// <summary>
    /// Generates a body for a Function or property getter.
    /// Declares a _result local, replaces FunctionReturnNode with assignments,
    /// and appends "return _result;" at the end.
    /// </summary>
    private void GenerateBodyWithReturn(IReadOnlyList<AstNode> body, string returnType)
    {
        _w.WriteLine($"{returnType} _result = default;");
        foreach (var stmt in body)
            GenerateStatement(stmt, returnType, resultVar: "_result");
        _w.WriteLine("return _result;");
    }

    // ── Statements ───────────────────────────────────────────────────────────

    private void GenerateStatement(AstNode node, string? returnType, string? resultVar)
    {
        switch (node)
        {
            case CommentNode c:
                _w.WriteLine($"// {c.Text.TrimStart('\'').Trim()}");
                break;

            case LocalDimNode d:
                foreach (var dec in d.Declarators)
                {
                    string typeStr = TypeStr(dec.TypeRef, dec.Name);
                    string init    = dec.DefaultValue != null ? $" = {Expr(dec.DefaultValue)}" : "";
                    _w.WriteLine($"{(d.IsStatic ? "static " : "")}{typeStr} {dec.Name}{init};{ReviewComment(dec.TypeRef)}");
                }
                break;

            case ReDimNode r:
                foreach (var dec in r.Declarators)
                {
                    string dims = dec.DefaultValue != null ? Expr(dec.DefaultValue) : "";
                    string baseType = dec.TypeRef != null
                        ? TypeStr(new TypeRefNode(dec.TypeRef.TypeName, false, false, 0), dec.Name)
                        : "object";
                    string preserve = r.IsPreserve ? " /* Preserve */" : "";
                    // ReDim with dimensions stored in DefaultValue (as-is from parser)
                    _w.WriteLine($"{dec.Name} = new {baseType}[{dims}]{preserve};");
                }
                break;

            case FunctionReturnNode r:
                if (resultVar != null)
                    _w.WriteLine($"{resultVar} = {Expr(r.Value)};");
                else
                    _w.WriteLine($"return {Expr(r.Value)};");
                break;

            case AssignmentNode a:
                // Drop "Set" keyword — object assignment is the same in C#
                _w.WriteLine($"{Expr(a.Target)} = {Expr(a.Value)};");
                break;

            case CallStatementNode c:
                string callTarget = Expr(c.Target);
                string callArgs   = c.Arguments.Count > 0
                    ? "(" + Args(c.Arguments) + ")"
                    : c.ExplicitCall ? "()" : "";
                // Detect Err.Raise → throw
                if (callTarget == "Err.Raise")
                {
                    string desc = c.Arguments.Count >= 3 ? Expr(c.Arguments[2].Value!) : "\"\"";
                    _w.WriteLine($"throw new Exception({desc}); // Err.Raise");
                }
                else
                {
                    _w.WriteLine($"{callTarget}{callArgs};");
                }
                break;

            case IfNode i:
                _w.WriteLine($"if ({Expr(i.Condition)})");
                _w.OpenBrace();
                foreach (var s in i.ThenBody) GenerateStatement(s, returnType, resultVar);
                _w.CloseBrace();
                foreach (var ei in i.ElseIfClauses)
                {
                    _w.WriteLine($"else if ({Expr(ei.Condition)})");
                    _w.OpenBrace();
                    foreach (var s in ei.Body) GenerateStatement(s, returnType, resultVar);
                    _w.CloseBrace();
                }
                if (i.ElseBody != null)
                {
                    _w.WriteLine("else");
                    _w.OpenBrace();
                    foreach (var s in i.ElseBody) GenerateStatement(s, returnType, resultVar);
                    _w.CloseBrace();
                }
                break;

            case SingleLineIfNode si:
                _w.WriteLine($"if ({Expr(si.Condition)}) {{ ");
                GenerateStatement(si.ThenStatement, returnType, resultVar);
                if (si.ElseStatement != null)
                {
                    _w.WriteLine("} else {");
                    GenerateStatement(si.ElseStatement, returnType, resultVar);
                }
                _w.WriteLine("}");
                break;

            case SelectCaseNode s:
                _w.WriteLine($"switch ({Expr(s.TestExpression)})");
                _w.OpenBrace();
                foreach (var c in s.Cases)
                {
                    if (c.IsElse)
                        _w.WriteLine("default:");
                    else
                        foreach (var pat in c.Patterns)
                            _w.WriteLine($"case {CasePattern(pat)}:");
                    _w.OpenBrace();
                    foreach (var stmt in c.Body)
                        GenerateStatement(stmt, returnType, resultVar);
                    _w.WriteLine("break;");
                    _w.CloseBrace();
                }
                _w.CloseBrace();
                break;

            case ForNextNode f:
                string forStep = f.Step != null ? $" + {Expr(f.Step)}" : "++";
                string forCond = f.Step != null ? "<= /* REVIEW: step direction */" : "<=";
                string stepStr = f.Step != null ? $" += {Expr(f.Step)}" : "++";
                _w.WriteLine($"for ({f.VariableName} = {Expr(f.Start)}; {f.VariableName} {forCond} {Expr(f.End)}; {f.VariableName}{stepStr})");
                _w.OpenBrace();
                foreach (var s in f.Body) GenerateStatement(s, returnType, resultVar);
                _w.CloseBrace();
                break;

            case ForEachNode fe:
                _w.WriteLine($"foreach (var {fe.VariableName} in {Expr(fe.Collection)})");
                _w.OpenBrace();
                foreach (var s in fe.Body) GenerateStatement(s, returnType, resultVar);
                _w.CloseBrace();
                break;

            case WhileNode w:
                _w.WriteLine($"while ({Expr(w.Condition)})");
                _w.OpenBrace();
                foreach (var s in w.Body) GenerateStatement(s, returnType, resultVar);
                _w.CloseBrace();
                break;

            case DoLoopNode d:
                GenerateDoLoop(d, returnType, resultVar);
                break;

            case WithNode w:
                string withExpr = Expr(w.Object);
                _withStack.Push(withExpr);
                _w.WriteLine($"{{ // With {withExpr}");
                _w.OpenBrace();
                foreach (var s in w.Body) GenerateStatement(s, returnType, resultVar);
                _w.CloseBrace();
                _w.WriteLine($"}} // End With");
                _withStack.Pop();
                break;

            case TryCatchNode tc:
                _w.WriteLine("try");
                _w.OpenBrace();
                foreach (var s in tc.TryBody)  GenerateStatement(s, returnType, resultVar);
                _w.CloseBrace();
                _w.WriteLine($"catch (Exception {tc.CatchVariable})");
                _w.OpenBrace();
                foreach (var s in tc.CatchBody) GenerateStatement(s, returnType, resultVar);
                _w.CloseBrace();
                break;

            case OnErrorNode o:
                switch (o.Kind)
                {
                    case OnErrorKind.ResumeNext:
                        _w.WriteLine("// VB6: On Error Resume Next — errors suppressed; review manually");
                        break;
                    case OnErrorKind.GoToZero:
                        _w.WriteLine("// VB6: On Error GoTo 0 — reset error handler");
                        break;
                    case OnErrorKind.GoTo:
                        _w.WriteLine($"// VB6: On Error GoTo {o.LabelName} — not restructured; review manually");
                        break;
                }
                break;

            case ResumeNode r:
                if (r.IsNext)
                    _w.WriteLine("// VB6: Resume Next");
                else if (r.LabelName != null)
                    _w.WriteLine($"goto {r.LabelName}; // VB6: Resume {r.LabelName}");
                else
                    _w.WriteLine("// TODO: Resume (retry semantics — no C# equivalent)");
                break;

            case ExitNode e:
                switch (e.What)
                {
                    case "Sub" or "Function" or "Property":
                        if (resultVar != null)
                            _w.WriteLine($"return {resultVar}; // Exit {e.What}");
                        else
                            _w.WriteLine($"return; // Exit {e.What}");
                        break;
                    case "For" or "Do" or "While":
                        _w.WriteLine("break;");
                        break;
                    default:
                        _w.WriteLine($"break; // Exit {e.What}");
                        break;
                }
                break;

            case GoToNode g:
                _w.WriteLine($"goto {g.Label};");
                break;

            case GoSubNode g:
                _w.WriteLine($"// TODO: GoSub {g.Label} — convert to method call");
                break;

            case ReturnNode:
                _w.WriteLine("// VB6: Return (from GoSub)");
                break;

            case LabelNode l:
                // Labels must be at the method body level; dedent by one for the label itself
                _w.WriteLine($"{l.Name}:;");
                break;

            case ErrorStatementNode e:
                _w.WriteLine($"throw new Exception(\"Error \" + {Expr(e.ErrorNumber)});");
                break;

            case EndStatementNode:
                _w.WriteLine("// VB6: End — terminates the application");
                _w.WriteLine("System.Environment.Exit(0);");
                break;
        }
    }

    private void GenerateDoLoop(DoLoopNode d, string? returnType, string? resultVar)
    {
        switch (d.Kind)
        {
            case DoLoopKind.DoWhileTop:
                _w.WriteLine($"while ({Expr(d.Condition!)})");
                _w.OpenBrace();
                foreach (var s in d.Body) GenerateStatement(s, returnType, resultVar);
                _w.CloseBrace();
                break;

            case DoLoopKind.DoUntilTop:
                _w.WriteLine($"while (!({Expr(d.Condition!)}))");
                _w.OpenBrace();
                foreach (var s in d.Body) GenerateStatement(s, returnType, resultVar);
                _w.CloseBrace();
                break;

            case DoLoopKind.DoWhileBottom:
                _w.WriteLine("do");
                _w.OpenBrace();
                foreach (var s in d.Body) GenerateStatement(s, returnType, resultVar);
                _w.CloseBrace($" while ({Expr(d.Condition!)});");
                break;

            case DoLoopKind.DoUntilBottom:
                _w.WriteLine("do");
                _w.OpenBrace();
                foreach (var s in d.Body) GenerateStatement(s, returnType, resultVar);
                _w.CloseBrace($" while (!({Expr(d.Condition!)}));");
                break;

            case DoLoopKind.DoForever:
                _w.WriteLine("while (true)");
                _w.OpenBrace();
                foreach (var s in d.Body) GenerateStatement(s, returnType, resultVar);
                _w.CloseBrace();
                break;
        }
    }

    // ── Expressions ──────────────────────────────────────────────────────────

    private string Expr(ExpressionNode e) => e switch
    {
        NothingNode                 => "null",
        MeNode                      => "this",
        BoolLiteralNode b           => b.Value ? "true" : "false",
        IntegerLiteralNode i        => i.RawText,
        DoubleLiteralNode d         => d.RawText,
        StringLiteralNode s         => s.RawText,
        DateLiteralNode d           => $"DateTime.Parse({d.RawText}) /* date literal */",

        IdentifierNode id           => MapIdentifier(id.Name),
        NewObjectNode n             => NewObjectExpr(n),
        TypeOfIsNode t              => $"({Expr(t.Operand)} is {t.TypeName})",

        MemberAccessNode m          => $"{Expr(m.Object)}.{m.MemberName}",
        BangAccessNode b            => $"{Expr(b.Object)}[\"{b.MemberName}\"] /* ! */",
        WithMemberAccessNode w      => $"{CurrentWith()}.{w.MemberName}",

        IndexNode ix                => $"{Expr(ix.Target)}[{Args(ix.Arguments)}]",
        CallOrIndexNode c           => MapCall(c),

        UnaryExpressionNode u       => MapUnary(u),
        BinaryExpressionNode b      => MapBinary(b),

        _                           => $"/* {e.GetType().Name} */"
    };

    private string NewObjectExpr(NewObjectNode n)
    {
        // VB6 "New Collection" → C# "new()" (target-typed new, C# 9+).
        // The element type is inferred by the compiler from the assignment target's
        // declared type, which is already correctly set to Collection<T> or
        // Collection<object> by the Transformer.
        if (n.TypeName.Equals("Collection", StringComparison.OrdinalIgnoreCase))
        {
            _requiredUsings.Add("System.Collections.ObjectModel");
            return "new()";
        }
        return $"new {n.TypeName}()";
    }

    private string MapIdentifier(string name)
    {
        if (BuiltInMap.TryGetConstant(name, out var csConst))
            return csConst;

        // Qualify enum members from this or other modules.
        // Same module → "EnumName.MemberName"; other module → "ModuleName.EnumName.MemberName".
        if (_enumMemberMap != null && _enumMemberMap.TryGetValue(name, out var ep))
        {
            return ep.ModuleName.Equals(_currentModuleName, StringComparison.OrdinalIgnoreCase)
                ? $"{ep.EnumName}.{name}"
                : $"{ep.ModuleName}.{ep.EnumName}.{name}";
        }

        return name;
    }

    private string MapCall(CallOrIndexNode c)
    {
        // Bare identifier calls — check for built-in functions
        if (c.Target is IdentifierNode id)
        {
            string[] argArr = c.Arguments.Select(a => a.IsMissing ? "/* missing */" : Expr(a.Value!)).ToArray();
            if (BuiltInMap.TryGetFunction(id.Name, argArr, out var mapped))
                return mapped;
        }

        // Member access calls — check for Debug.Print etc.
        if (c.Target is MemberAccessNode ma)
        {
            string fullName = $"{Expr(ma.Object)}.{ma.MemberName}";
            string[] argArr = c.Arguments.Select(a => a.IsMissing ? "/* missing */" : Expr(a.Value!)).ToArray();
            if (BuiltInMap.TryGetFunction(fullName, argArr, out var mapped))
                return mapped;

            // Err object property mappings
            if (Expr(ma.Object) == "Err")
            {
                return ma.MemberName switch
                {
                    "Number"      => "_ex.HResult /* Err.Number */",
                    "Description" => "_ex.Message",
                    "Source"      => "_ex.Source ?? \"\" /* Err.Source */",
                    "Clear"       => "/* Err.Clear() */",
                    _             => $"Err.{ma.MemberName} /* TODO */"
                };
            }
        }

        // General call
        string target = Expr(c.Target);
        string args   = c.Arguments.Count > 0 ? $"({Args(c.Arguments)})" : "()";
        return $"{target}{args}";
    }

    private string MapUnary(UnaryExpressionNode u)
    {
        string operand = Expr(u.Operand);
        return u.Operator switch
        {
            TokenKind.KwNot => $"!({operand})",
            TokenKind.Minus => $"-{operand}",
            TokenKind.Plus  => operand,
            _               => $"/* {u.Operator} */ {operand}"
        };
    }

    // Comparison operators that may involve enum members compared against int fields.
    private static readonly HashSet<TokenKind> NumericComparisonOps =
    [
        TokenKind.Equals, TokenKind.NotEqual,
        TokenKind.LessThan, TokenKind.GreaterThan,
        TokenKind.LessEqual, TokenKind.GreaterEqual,
    ];

    /// <summary>
    /// Returns true when the expression is a bare identifier that resolves to
    /// an enum member in the cross-module enum map.
    /// </summary>
    private bool IsEnumMember(ExpressionNode e) =>
        _enumMemberMap != null &&
        e is IdentifierNode id &&
        _enumMemberMap.ContainsKey(id.Name);

    /// <summary>
    /// Returns true when the terminal member name of the expression is a field,
    /// property, or UDT field whose VB6 declared type is an enum type.
    /// Used to suppress the (int) cast when the other side is already enum-typed.
    /// </summary>
    private bool IsEnumTypedField(ExpressionNode e)
    {
        if (_enumTypedFieldNames == null) return false;
        string? name = e switch
        {
            IdentifierNode id    => id.Name,
            MemberAccessNode m   => m.MemberName,
            _                    => null,
        };
        return name != null && _enumTypedFieldNames.Contains(name);
    }

    private string MapBinary(BinaryExpressionNode b)
    {
        string left  = Expr(b.Left);
        string right = Expr(b.Right);

        // Exponentiation → Math.Pow
        if (b.Operator == TokenKind.Caret)
            return $"Math.Pow({left}, {right})";

        // Integer division — emit cast to make intent clear
        if (b.Operator == TokenKind.Backslash)
            return $"((int)({left}) / (int)({right}))";

        // In C# enum values cannot be implicitly compared with int fields.
        // In VB6 all enums are plain integers, so the comparison always worked.
        // Cast the enum side to (int) — but only when the other side is NOT already
        // an enum-typed field (if both sides are the same enum type, no cast is needed).
        if (NumericComparisonOps.Contains(b.Operator))
        {
            if (IsEnumMember(b.Left)  && !IsEnumTypedField(b.Right)) left  = $"(int){left}";
            if (IsEnumMember(b.Right) && !IsEnumTypedField(b.Left))  right = $"(int){right}";
        }

        string op = b.Operator switch
        {
            TokenKind.Plus         => "+",
            TokenKind.Minus        => "-",
            TokenKind.Star         => "*",
            TokenKind.Slash        => "/",
            TokenKind.Ampersand    => "+",   // string concat
            TokenKind.KwMod        => "%",
            TokenKind.Equals       => "==",
            TokenKind.NotEqual     => "!=",
            TokenKind.LessThan     => "<",
            TokenKind.GreaterThan  => ">",
            TokenKind.LessEqual    => "<=",
            TokenKind.GreaterEqual => ">=",
            TokenKind.KwAnd        => "&&",
            TokenKind.KwOr         => "||",
            TokenKind.KwXor        => "^",
            TokenKind.KwIs         => "==",  // reference equality
            TokenKind.KwEqv        => "==",  // approx
            TokenKind.KwImp        => "/* Imp */ ||",  // !left || right
            TokenKind.KwLike       => $"/* Like — use Regex */ ==",
            _                      => $"/* {b.Operator} */"
        };

        return $"({left} {op} {right})";
    }

    private string CasePattern(CasePatternNode p) => p switch
    {
        CaseValuePattern v  => Expr(v.Value),
        CaseRangePattern r  => $"{Expr(r.Low)}: case {Expr(r.High)} /* range */",
        CaseIsPattern i     => $"/* Is */ {MapCaseOp(i.Operator)} {Expr(i.Value)}",
        _                   => "default"
    };

    private static string MapCaseOp(TokenKind op) => op switch
    {
        TokenKind.Equals       => "==",
        TokenKind.NotEqual     => "!=",
        TokenKind.LessThan     => "<",
        TokenKind.GreaterThan  => ">",
        TokenKind.LessEqual    => "<=",
        TokenKind.GreaterEqual => ">=",
        _                      => op.ToString()
    };

    // ── Helpers ───────────────────────────────────────────────────────────────

    private string Params(IReadOnlyList<ParameterNode> parameters)
    {
        if (parameters.Count == 0) return "";
        return string.Join(", ", parameters.Select(ParamStr));
    }

    private string ParamStr(ParameterNode p)
    {
        string modifier  = p.Mode == ParameterMode.ByRef ? "ref " : "";
        string optional  = p.IsOptional ? "/* Optional */ " : "";
        string paramArr  = p.IsParamArray ? "params " : "";
        string typeStr   = TypeStr(p.TypeRef, p.Name);
        string defVal    = p.DefaultValue != null ? $" = {Expr(p.DefaultValue)}" : "";
        return $"{optional}{paramArr}{modifier}{typeStr} {p.Name}{defVal}";
    }

    private string Args(IReadOnlyList<ArgumentNode> args) =>
        string.Join(", ", args.Select(a =>
            a.IsMissing ? "/* missing */" :
            a.Name != null ? $"/* {a.Name}: */ {Expr(a.Value!)}" :
            Expr(a.Value!)));

    private string TypeStr(TypeRefNode? t, string context)
    {
        if (t == null) return "object";
        string name = t.TypeName;

        if (name.StartsWith("Collection<"))
            _requiredUsings.Add("System.Collections.ObjectModel");

        // Arrays
        if (t.IsArray)
        {
            string dims = t.ArrayRank > 1 ? string.Concat(Enumerable.Repeat(",", t.ArrayRank - 1)) : "";
            return $"{name}[{dims}]";
        }
        return name;
    }

    /// <summary>
    /// Returns an inline // REVIEW: comment to append to a declaration line,
    /// or null when no review is needed.
    /// </summary>
    private static string? ReviewComment(TypeRefNode? t) =>
        t?.TypeName == "Collection<object>"
            ? " // REVIEW: Collection<object> — replace with List<T> or Dictionary<string, T> once element type is known"
            : null;

    private static string Access(AccessModifier a) => a switch
    {
        AccessModifier.Public  => "public ",
        AccessModifier.Private => "private ",
        AccessModifier.Friend  => "internal ",
        AccessModifier.Global  => "public ",
        _                      => "public "
    };

    private string StaticMod() => _isStatic ? "static " : "";

    private string CurrentWith() =>
        _withStack.Count > 0 ? _withStack.Peek() : "/* With object */";
}
