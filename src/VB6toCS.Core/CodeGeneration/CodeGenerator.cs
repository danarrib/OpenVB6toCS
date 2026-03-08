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

    // Cross-module default member map: className → defaultMemberName (VB_UserMemId = 0).
    // Used to expand VB6 default member calls: obj(arg) → obj.DefaultMember(arg).
    private readonly IReadOnlyDictionary<string, string>? _defaultMemberMap;

    // Cross-module method parameter mode map: moduleName → (methodName → [ParameterMode, ...]).
    // Used to emit `ref` at call sites for ByRef parameters.
    private readonly IReadOnlyDictionary<string, IReadOnlyDictionary<string, IReadOnlyList<Parsing.Nodes.ParameterMode>>>? _methodParamMap;

    // Cross-module global variable map: publicFieldName → typeName (collected from all modules).
    // Used to resolve qualified calls like `M46V999.gfObtSYASMULTINI` where M46V999 is a
    // public field declared in another module (not in the current local/module scope).
    private readonly IReadOnlyDictionary<string, string>? _globalVarMap;

    private string _currentModuleName = "";

    // Type scope for implicit-coercion detection.
    // Module-level field types (populated once per module, never cleared).
    private readonly Dictionary<string, string> _moduleFieldTypes = new(StringComparer.OrdinalIgnoreCase);
    // Procedure-level types: parameters + locals (cleared on each procedure entry).
    private readonly Dictionary<string, string> _procTypes = new(StringComparer.OrdinalIgnoreCase);

    private CodeGenerator(bool isStatic,
        IReadOnlyDictionary<string, (string ModuleName, string EnumName)>? enumMemberMap,
        IReadOnlySet<string>? enumTypedFieldNames,
        IReadOnlyDictionary<string, string>? defaultMemberMap,
        IReadOnlyDictionary<string, IReadOnlyDictionary<string, IReadOnlyList<Parsing.Nodes.ParameterMode>>>? methodParamMap,
        IReadOnlyDictionary<string, string>? globalVarMap)
    {
        _isStatic = isStatic;
        _enumMemberMap = enumMemberMap;
        _enumTypedFieldNames = enumTypedFieldNames;
        _defaultMemberMap = defaultMemberMap;
        _methodParamMap = methodParamMap;
        _globalVarMap = globalVarMap;
    }

    public static string Generate(ModuleNode module, bool isStaticModule,
        IReadOnlyDictionary<string, (string ModuleName, string EnumName)>? enumMemberMap = null,
        IReadOnlySet<string>? enumTypedFieldNames = null,
        IReadOnlyDictionary<string, string>? defaultMemberMap = null,
        IReadOnlyDictionary<string, IReadOnlyDictionary<string, IReadOnlyList<Parsing.Nodes.ParameterMode>>>? methodParamMap = null,
        IReadOnlyDictionary<string, string>? globalVarMap = null)
    {
        var gen = new CodeGenerator(isStaticModule, enumMemberMap, enumTypedFieldNames, defaultMemberMap, methodParamMap, globalVarMap);
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

        // Collect module-level field and property types for type-aware code generation.
        _moduleFieldTypes.Clear();
        foreach (var member in module.Members)
        {
            if (member is FieldNode f)
                foreach (var d in f.Declarators)
                    if (d.TypeRef != null)
                        _moduleFieldTypes[d.Name] = d.TypeRef.TypeName;
            // Note: CsPropertyNode pattern match fails with stale obj; use explicit cast
            var csP = member as CsPropertyNode;
            if (csP?.Type != null)
                _moduleFieldTypes[csP.Name] = csP.Type.TypeName;
        }

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
        EnterProcScope(s.Parameters);
        GenerateBody(s.Body, returnType: null);
        _w.CloseBrace();
    }

    private void GenerateFunction(FunctionNode f)
    {
        string staticKw   = _isStatic || f.IsStatic ? "static " : "";
        string returnType = TypeStr(f.ReturnType, f.Name);
        returnType = InferActualReturnType(f.Parameters, f.Body, returnType);
        _w.WriteLine($"{Access(f.Access)}{staticKw}{returnType} {f.Name}({Params(f.Parameters)})");
        _w.OpenBrace();
        EnterProcScope(f.Parameters);
        GenerateBodyWithReturn(f.Body, returnType);
        _w.CloseBrace();
    }

    private void GenerateProperty(CsPropertyNode p)
    {
        // Parameterized properties cannot be represented as C# properties.
        // A Property Get with parameters, or a Let/Set with more than one parameter
        // (the extra ones beyond the value), must become methods.
        bool getHasParams  = p.GetBody  != null && p.GetParameters.Count  > 0;
        bool letHasExtra   = p.LetBody  != null && p.LetParameters.Count  > 1;
        bool setHasExtra   = p.SetBody  != null && p.SetParameters.Count  > 1;
        if (getHasParams || letHasExtra || setHasExtra)
        {
            GenerateParameterizedPropertyAsMethods(p);
            return;
        }

        string staticKw  = _isStatic || p.IsStatic ? "static " : "";
        string propType  = TypeStr(p.Type, p.Name);
        _w.WriteLine($"{Access(p.Access)}{staticKw}{propType} {p.Name}");
        _w.OpenBrace();

        if (p.GetBody != null)
        {
            EnterProcScope(p.GetParameters);
            _w.WriteLine("get");
            _w.OpenBrace();
            GenerateBodyWithReturn(p.GetBody, propType);
            _w.CloseBrace();
        }

        if (p.LetBody != null)
        {
            EnterProcScope(p.LetParameters);
            string valueParam = p.LetParameters.Count > 0
                ? p.LetParameters[^1].Name : "value";
            _w.WriteLine($"set  // Let — parameter: {valueParam}");
            _w.OpenBrace();
            GenerateBody(p.LetBody, returnType: null);
            _w.CloseBrace();
        }

        if (p.SetBody != null)
        {
            EnterProcScope(p.SetParameters);
            string valueParam = p.SetParameters.Count > 0
                ? p.SetParameters[^1].Name : "value";
            _w.WriteLine($"set  // Set — parameter: {valueParam}");
            _w.OpenBrace();
            GenerateBody(p.SetBody, returnType: null);
            _w.CloseBrace();
        }

        _w.CloseBrace();
    }

    /// <summary>
    /// Emits a VB6 Property Get/Let/Set that has extra parameters as one or more C# methods.
    /// C# properties cannot have parameters; Property Get with params becomes a getter method,
    /// Property Let/Set with extra params becomes a setter method.
    /// </summary>
    private void GenerateParameterizedPropertyAsMethods(CsPropertyNode p)
    {
        string staticKw = _isStatic || p.IsStatic ? "static " : "";
        string propType = TypeStr(p.Type, p.Name);
        _w.WriteLine("// This was a VB6 Property with parameters. C# properties cannot have parameters;");
        _w.WriteLine("// converted to method(s). Consider implementing as an indexer (this[...]) if appropriate.");

        if (p.GetBody != null)
        {
            EnterProcScope(p.GetParameters);
            _w.WriteLine($"{Access(p.Access)}{staticKw}{propType} {p.Name}({Params(p.GetParameters)})");
            _w.OpenBrace();
            GenerateBodyWithReturn(p.GetBody, propType);
            _w.CloseBrace();
        }

        if (p.LetBody != null)
        {
            EnterProcScope(p.LetParameters);
            // All parameters including the trailing value parameter
            _w.WriteLine($"{Access(p.Access)}{staticKw}void {p.Name}({Params(p.LetParameters)})  // Property Let");
            _w.OpenBrace();
            GenerateBody(p.LetBody, returnType: null);
            _w.CloseBrace();
        }

        if (p.SetBody != null)
        {
            EnterProcScope(p.SetParameters);
            _w.WriteLine($"{Access(p.Access)}{staticKw}void {p.Name}({Params(p.SetParameters)})  // Property Set");
            _w.OpenBrace();
            GenerateBody(p.SetBody, returnType: null);
            _w.CloseBrace();
        }
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
                    if (dec.TypeRef != null) _procTypes[dec.Name] = dec.TypeRef.TypeName;
                    string typeStr = TypeStr(dec.TypeRef, dec.Name);
                    string init    = dec.DefaultValue != null
                        ? $" = {Expr(dec.DefaultValue)}"
                        : $" = {DefaultInit(typeStr)}";
                    _w.WriteLine($"{(d.IsStatic ? "static " : "")}{typeStr} {dec.Name}{init};{ReviewComment(dec.TypeRef)}");
                }
                break;

            case ReDimNode r:
                for (int _ri = 0; _ri < r.Declarators.Count; _ri++)
                {
                    var dec      = r.Declarators[_ri];
                    string baseType = dec.TypeRef != null
                        ? TypeStr(new TypeRefNode(dec.TypeRef.TypeName, false, false, 0), dec.Name)
                        : "object";
                    string preserve = r.IsPreserve ? " /* Preserve */" : "";
                    // Emit dimension expressions captured by the parser.
                    var dimExprs = _ri < r.DimensionLists.Count ? r.DimensionLists[_ri] : null;
                    string dims  = dimExprs != null && dimExprs.Count > 0
                        ? string.Join(", ", dimExprs.Select(Expr))
                        : "";
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
                // When ExplicitCall is true but the target is already a CallOrIndexNode (e.g.
                // `Call Foo(x)` where the parser absorbed the parens into the node), don't add
                // another "()" — that would produce double parentheses like `Foo(x)()`.
                // If there are arguments, emit them.
                // If the target is a CallOrIndexNode, its Expr() already includes "(...)" so
                // don't add another pair. Otherwise always emit "()" — even when ExplicitCall
                // is false (VB6 Sub calls without the Call keyword omit parens, but C# requires them).
                string callArgs   = c.Arguments.Count > 0
                    ? "(" + ArgsWithRef(c.Arguments, TryResolveProcParamModes(c.Target)) + ")"
                    : c.Target is not CallOrIndexNode ? "()" : "";
                // Detect Err.Raise → throw
                if (callTarget == "Err.Raise")
                {
                    string desc = c.Arguments.Count >= 3 ? Expr(c.Arguments[2].Value!) : "\"\"";
                    _w.WriteLine($"throw new Exception({desc}); // Err.Raise");
                }
                else if (callTarget == "Err.Clear" || callTarget == "/* Err.Clear() */")
                {
                    _w.WriteLine("// Err.Clear(); // no direct C# equivalent — review error-handling logic");
                }
                else
                {
                    _w.WriteLine($"{callTarget}{callArgs};");
                }
                break;

            case IfNode i:
                string ifComment = i.ThenComment != null ? $"  // {i.ThenComment.Text.TrimStart('\'').Trim()}" : "";
                _w.WriteLine($"if ({Expr(i.Condition)}){ifComment}");
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
        // Bare identifier calls — type-aware overrides + built-in function mapping
        if (c.Target is IdentifierNode id)
        {
            // IsMissing(x) — VB6 Optional parameter omission check.
            // Maps to (x == null) for reference/object types; (x == default) for value types.
            if (id.Name.Equals("IsMissing", StringComparison.OrdinalIgnoreCase) &&
                c.Arguments.Count == 1 && !c.Arguments[0].IsMissing)
            {
                string argExpr = Expr(c.Arguments[0].Value!);
                string? argType = TryGetSimpleType(c.Arguments[0].Value!);
                string sentinel = (argType != null && NonNullableValueTypes.Contains(argType))
                    ? "default" : "null";
                return $"({argExpr} == {sentinel})";
            }

            // Type-aware Val() — choose the conversion based on the argument's known type.
            if ((id.Name.Equals("Val", StringComparison.OrdinalIgnoreCase) ||
                 id.Name.Equals("CDbl", StringComparison.OrdinalIgnoreCase)) &&
                c.Arguments.Count == 1 && !c.Arguments[0].IsMissing)
            {
                var argNode = c.Arguments[0].Value!;
                string argExpr = Expr(argNode);
                string? argType = TryGetSimpleType(argNode);
                string converted = IsNullableValueType(argType)
                    ? $"Convert.ToDouble({argExpr}.GetValueOrDefault())"
                    : $"Convert.ToDouble({argExpr})";
                if (IsDefaultMemberCallArg(argNode))
                    converted += " /* TODO: default-member object — may need .Value */";
                return converted;
            }

            // Type-aware string-coercion functions (Trim, LTrim, RTrim, UCase, LCase, Len).
            // For string args: method call directly (no .ToString() needed).
            // For nullable value types: .GetValueOrDefault().ToString() first.
            // For non-nullable value types / unknown: .ToString() (safe, compiles for all).
            string? strFnSuffix = id.Name.ToLowerInvariant() switch
            {
                "trim"  or "trim$"  => ".Trim()",
                "ltrim" or "ltrim$" => ".TrimStart()",
                "rtrim" or "rtrim$" => ".TrimEnd()",
                "ucase" or "ucase$" => ".ToUpper()",
                "lcase" or "lcase$" => ".ToLower()",
                "len"               => ".Length",
                _ => null
            };
            if (strFnSuffix != null && c.Arguments.Count == 1 && !c.Arguments[0].IsMissing)
            {
                const string todoMarker = " /* TODO: default-member object — may need .Value */";
                var argNode = c.Arguments[0].Value!;
                string argExpr = Expr(argNode);
                string? argType = TryGetSimpleType(argNode);

                // If the inner expression already carries a TODO marker (e.g. from Val()),
                // strip it before composing, then re-append at the outermost level.
                bool hasTodo = argExpr.EndsWith(todoMarker);
                if (hasTodo) argExpr = argExpr[..^todoMarker.Length];

                string strExpr = argType == "string"
                    ? argExpr
                    : IsNullableValueType(argType)
                        ? $"{argExpr}.GetValueOrDefault().ToString()"
                        : $"{argExpr}.ToString()";
                string result = $"{strExpr}{strFnSuffix}";
                return hasTodo ? result + todoMarker : result;
            }

            // Type-aware string-position functions (Left, Right, Mid, InStr, InStrRev).
            // If the first string argument is not a string type, wrap it with .ToString().
            // This prevents compile errors when e.g. Left(intVar, 2) is called.
            string lcFn = id.Name.ToLowerInvariant().TrimEnd('$');
            if (lcFn is "left" or "right" or "mid" or "instr" or "instrrev" &&
                c.Arguments.Count >= 2 && !c.Arguments[0].IsMissing)
            {
                switch (lcFn)
                {
                    case "left" when c.Arguments.Count >= 2:
                    {
                        string str = EnsureStringExpr(c.Arguments[0].Value!);
                        string len = Expr(c.Arguments[1].Value!);
                        return $"{str}.Substring(0, {len})";
                    }
                    case "right" when c.Arguments.Count >= 2:
                    {
                        string str = EnsureStringExpr(c.Arguments[0].Value!);
                        string len = Expr(c.Arguments[1].Value!);
                        return $"{str}.Substring({str}.Length - {len})";
                    }
                    case "mid" when c.Arguments.Count >= 2:
                    {
                        string str   = EnsureStringExpr(c.Arguments[0].Value!);
                        string start = Expr(c.Arguments[1].Value!);
                        return c.Arguments.Count >= 3 && !c.Arguments[2].IsMissing
                            ? $"{str}.Substring({start} - 1, {Expr(c.Arguments[2].Value!)})"
                            : $"{str}.Substring({start} - 1)";
                    }
                    case "instr" when c.Arguments.Count >= 2:
                    {
                        // 2-arg form: InStr(str, pattern); 3-arg: InStr(start, str, pattern)
                        if (c.Arguments.Count >= 3 && !c.Arguments[2].IsMissing)
                        {
                            string str     = EnsureStringExpr(c.Arguments[1].Value!);
                            string pattern = EnsureStringExpr(c.Arguments[2].Value!);
                            return $"({str}.IndexOf({pattern}) + 1)";
                        }
                        else
                        {
                            string str     = EnsureStringExpr(c.Arguments[0].Value!);
                            string pattern = EnsureStringExpr(c.Arguments[1].Value!);
                            return $"({str}.IndexOf({pattern}) + 1)";
                        }
                    }
                    case "instrrev" when c.Arguments.Count >= 2:
                    {
                        string str     = EnsureStringExpr(c.Arguments[0].Value!);
                        string pattern = EnsureStringExpr(c.Arguments[1].Value!);
                        return $"({str}.LastIndexOf({pattern}) + 1)";
                    }
                }
            }

            string[] argArr = c.Arguments.Select(a => a.IsMissing ? "/* missing */" : Expr(a.Value!)).ToArray();
            if (BuiltInMap.TryGetFunction(id.Name, argArr, out var mapped))
            {
                // If a built-in that coerces to string/double receives an object returned from
                // a default-member expansion, append a TODO: the caller may need .PropertyName.
                // Strip any inner TODO from nested calls first, then re-append at the outermost level.
                const string todoMarker = " /* TODO: default-member object — may need .Value */";
                bool innerTodo = argArr.Length > 0 && argArr[0].EndsWith(todoMarker);
                if (innerTodo)
                {
                    argArr[0] = argArr[0][..^todoMarker.Length];
                    BuiltInMap.TryGetFunction(id.Name, argArr, out mapped);
                    mapped += todoMarker;
                }
                else if (c.Arguments.Count > 0 && IsDefaultMemberCallArg(c.Arguments[0].Value))
                    mapped += todoMarker;
                return mapped;
            }
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
                    "Clear"       => "/* Err.Clear() */",   // caught by CallStatementNode → commented out
                    _             => $"Err.{ma.MemberName} /* TODO */"
                };
            }
        }

        // Default member expansion: obj(arg) → obj.DefaultMember(arg)
        // Applies when obj's declared type has a VB_UserMemId = 0 property.
        if (c.Target is IdentifierNode idTarget && c.Arguments.Count > 0 && _defaultMemberMap != null)
        {
            string? varType = _procTypes.TryGetValue(idTarget.Name, out var t1) ? t1
                            : _moduleFieldTypes.TryGetValue(idTarget.Name, out var t2) ? t2 : null;
            if (varType != null && _defaultMemberMap.TryGetValue(varType, out var defMember))
                return $"{idTarget.Name}.{defMember}({Args(c.Arguments)})";
        }

        // COM indexed-member check: obj.Member(i) → obj.Member[i]
        // Applies when the receiver's declared type and member name are in ComTypeMap
        // (e.g. ADODB.Recordset.Fields, DAO.Recordset.Fields, …).
        if (c.Target is MemberAccessNode comMa && c.Arguments.Count > 0)
        {
            string? receiverType = TryGetSimpleType(comMa.Object);
            if (receiverType != null && ComTypeMap.IsIndexedMember(receiverType, comMa.MemberName))
                return $"{Expr(c.Target)}[{Args(c.Arguments)}]";
        }

        // General call
        string target = Expr(c.Target);
        string args   = c.Arguments.Count > 0
            ? $"({ArgsWithRef(c.Arguments, TryResolveProcParamModes(c.Target))})"
            : "()";
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

    // C# value type names produced by the Transformer's type mapping.
    private static readonly HashSet<string> NonNullableValueTypes = new(StringComparer.OrdinalIgnoreCase)
        { "int", "long", "double", "float", "bool", "byte", "short", "DateTime", "decimal" };

    private static bool IsNullableValueType(string? csType) =>
        csType != null && csType.EndsWith('?') &&
        NonNullableValueTypes.Contains(csType[..^1]);

    /// <summary>
    /// Returns true when the expression is a bare identifier that resolves to
    /// an enum member in the cross-module enum map.
    /// </summary>
    private bool IsEnumMember(ExpressionNode e) =>
        _enumMemberMap != null &&
        e is IdentifierNode id &&
        _enumMemberMap.ContainsKey(id.Name);

    /// <summary>
    /// Returns true when an expression argument is a CallOrIndexNode whose target is
    /// a variable with a declared type that has a default member.  Used to emit a
    /// TODO annotation when a built-in coercion function (Val, Trim, …) receives an
    /// object returned from a default-member expansion, since .ToString() on that
    /// object may not yield the expected value.
    /// </summary>
    private bool IsDefaultMemberCallArg(ExpressionNode? e)
    {
        if (e is not CallOrIndexNode c) return false;
        if (c.Target is not IdentifierNode idTarget) return false;
        if (_defaultMemberMap == null) return false;
        string? varType = _procTypes.TryGetValue(idTarget.Name, out var t1) ? t1
                        : _moduleFieldTypes.TryGetValue(idTarget.Name, out var t2) ? t2 : null;
        return varType != null && _defaultMemberMap.ContainsKey(varType);
    }

    // ── Scope helpers ─────────────────────────────────────────────────────────

    /// <summary>Resets procedure-level scope and populates it from the parameter list.</summary>
    private void EnterProcScope(IReadOnlyList<ParameterNode> parameters)
    {
        _procTypes.Clear();
        foreach (var p in parameters)
            if (p.TypeRef != null)
                _procTypes[p.Name] = p.TypeRef.TypeName;
    }

    /// <summary>
    /// Attempts to infer a more specific return type for a function declared as
    /// <c>object</c> (was <c>Variant</c> in VB6) by examining what is actually
    /// returned via <see cref="FunctionReturnNode"/> assignments in the body.
    ///
    /// Works by pre-building a full locals map (parameters + all <c>Dim</c>
    /// declarations in all scopes), then calling <see cref="TypeStr"/> on the
    /// TypeRef of each returned identifier.  Returns the declared type unchanged
    /// when inference is ambiguous or yields no result.
    /// </summary>
    private string InferActualReturnType(
        IReadOnlyList<ParameterNode> parameters,
        IReadOnlyList<AstNode> body,
        string declaredReturnType)
    {
        if (!declaredReturnType.Equals("object", StringComparison.OrdinalIgnoreCase))
            return declaredReturnType;

        // Build a full type-ref map: parameters + all locals in the body.
        var typeRefs = new Dictionary<string, TypeRefNode>(StringComparer.OrdinalIgnoreCase);
        foreach (var p in parameters)
            if (p.TypeRef != null) typeRefs[p.Name] = p.TypeRef;
        CollectBodyTypeRefs(body, typeRefs);

        // Collect all FunctionReturnNode values throughout the body.
        var returnValues = new List<ExpressionNode>();
        CollectReturnValues(body, returnValues);
        if (returnValues.Count == 0) return declaredReturnType;

        // Evaluate each returned expression — only identifier nodes are resolvable here.
        string? inferred = null;
        foreach (var rv in returnValues)
        {
            if (rv is not IdentifierNode rvId) return declaredReturnType; // complex expression
            if (!typeRefs.TryGetValue(rvId.Name, out var typeRef)) return declaredReturnType;
            string csType = TypeStr(typeRef, rvId.Name);
            if (csType == "object") return declaredReturnType; // still unresolved
            if (inferred == null)
                inferred = csType;
            else if (!inferred.Equals(csType, StringComparison.Ordinal))
                return declaredReturnType; // conflicting types
        }
        return inferred ?? declaredReturnType;
    }

    /// <summary>
    /// Recursively collects TypeRef nodes for all local variable declarations
    /// in every nested scope of <paramref name="body"/>.
    /// </summary>
    private static void CollectBodyTypeRefs(
        IReadOnlyList<AstNode> body,
        Dictionary<string, TypeRefNode> result)
    {
        foreach (var node in body)
        {
            if (node is LocalDimNode dim)
                foreach (var d in dim.Declarators)
                    if (d.TypeRef != null) result[d.Name] = d.TypeRef;

            // Recurse into every statement that has a child body.
            IEnumerable<IReadOnlyList<AstNode>> children = node switch
            {
                IfNode i         => i.ElseBody != null
                                    ? i.ElseIfClauses.Select(e => e.Body)
                                        .Append(i.ThenBody).Append(i.ElseBody)
                                    : i.ElseIfClauses.Select(e => e.Body).Append(i.ThenBody),
                SelectCaseNode s => s.Cases.Select(c => c.Body),
                ForNextNode f    => [f.Body],
                ForEachNode f    => [f.Body],
                WhileNode w      => [w.Body],
                DoLoopNode d     => [d.Body],
                WithNode w       => [w.Body],
                TryCatchNode tc  => [tc.TryBody, tc.CatchBody],
                _                => []
            };
            foreach (var child in children)
                CollectBodyTypeRefs(child, result);
        }
    }

    /// <summary>
    /// Recursively collects the value-expression of every
    /// <see cref="FunctionReturnNode"/> in all nested scopes of <paramref name="body"/>.
    /// </summary>
    private static void CollectReturnValues(
        IReadOnlyList<AstNode> body,
        List<ExpressionNode> result)
    {
        foreach (var node in body)
        {
            if (node is FunctionReturnNode ret)
                result.Add(ret.Value);

            IEnumerable<IReadOnlyList<AstNode>> children = node switch
            {
                IfNode i         => i.ElseBody != null
                                    ? i.ElseIfClauses.Select(e => e.Body)
                                        .Append(i.ThenBody).Append(i.ElseBody)
                                    : i.ElseIfClauses.Select(e => e.Body).Append(i.ThenBody),
                SelectCaseNode s => s.Cases.Select(c => c.Body),
                ForNextNode f    => [f.Body],
                ForEachNode f    => [f.Body],
                WhileNode w      => [w.Body],
                DoLoopNode d     => [d.Body],
                WithNode w       => [w.Body],
                TryCatchNode tc  => [tc.TryBody, tc.CatchBody],
                _                => []
            };
            foreach (var child in children)
                CollectReturnValues(child, result);
        }
    }

    /// <summary>
    /// Returns the C# type name for simple expressions whose type can be determined
    /// without cross-module resolution: literals and identifiers in the current scope.
    /// Returns null when the type is unknown or complex (e.g. member-access chains).
    /// </summary>
    private string? TryGetSimpleType(ExpressionNode e) => e switch
    {
        StringLiteralNode    => "string",
        IntegerLiteralNode   => "int",
        DoubleLiteralNode    => "double",
        BoolLiteralNode      => "bool",
        IdentifierNode id    => _procTypes.TryGetValue(id.Name, out var t) ? t
                              : _moduleFieldTypes.TryGetValue(id.Name, out t) ? t
                              : null,
        _                    => null,
    };

    private static bool IsNumericCsType(string? type) =>
        type is "int" or "long" or "double" or "float" or "decimal" or "byte";

    /// <summary>
    /// Returns a C# expression for <paramref name="node"/> guaranteed to be a string.
    /// String types are returned as-is; nullable value types get .GetValueOrDefault().ToString();
    /// everything else gets .ToString().
    /// </summary>
    private string EnsureStringExpr(ExpressionNode node)
    {
        string expr  = Expr(node);
        string? type = TryGetSimpleType(node);
        if (type == "string") return expr;
        if (IsNullableValueType(type)) return $"{expr}.GetValueOrDefault().ToString()";
        return $"{expr}.ToString()";
    }

    // VB6 date/time system identifiers (no arguments) that resolve to DateTime values.
    private static readonly HashSet<string> Vb6DateTimeNames =
        new(StringComparer.OrdinalIgnoreCase) { "Date", "Now", "Time" };

    /// <summary>
    /// Returns true when the expression is a VB6 date/time system identifier (Date, Now, Time)
    /// or a variable/field whose declared type is DateTime.
    /// </summary>
    private bool IsDateTimeExpression(ExpressionNode e)
    {
        if (e is IdentifierNode id && Vb6DateTimeNames.Contains(id.Name))
            return true;
        return TryGetSimpleType(e) == "DateTime";
    }

    /// <summary>
    /// Converts a numeric expression string to its string equivalent for use in
    /// string comparisons. Literals are wrapped in quotes; variables get .ToString().
    /// </summary>
    private static string NumericToString(ExpressionNode node, string emitted) => node switch
    {
        IntegerLiteralNode i => $"\"{i.Value}\"",
        DoubleLiteralNode  d => $"\"{d.RawText}\"",
        _                    => $"{emitted}.ToString()",
    };

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

            // VB6 also coerces between string and numeric types implicitly.
            // In C# this is a compile error; convert the numeric side to string
            // (safer than parsing the string, which may throw on non-numeric values).
            string? leftType  = TryGetSimpleType(b.Left);
            string? rightType = TryGetSimpleType(b.Right);
            if (leftType == "string" && IsNumericCsType(rightType))
                right = NumericToString(b.Right, right);
            else if (rightType == "string" && IsNumericCsType(leftType))
                left  = NumericToString(b.Left, left);
        }

        // VB6 date arithmetic: Date ± n means add/subtract days.
        // In C#, DateTime does not support + / - with a bare integer; use .AddDays().
        if ((b.Operator == TokenKind.Plus || b.Operator == TokenKind.Minus) &&
            IsDateTimeExpression(b.Left) &&
            (IsNumericCsType(TryGetSimpleType(b.Right)) ||
             b.Right is IntegerLiteralNode || b.Right is DoubleLiteralNode))
        {
            string daysArg = b.Operator == TokenKind.Minus
                ? (b.Right is IntegerLiteralNode il ? $"-{il.Value}"
                   : b.Right is DoubleLiteralNode  dl ? $"-{dl.RawText}"
                   : $"-({right})")
                : right;
            return $"{left}.AddDays({daysArg})";
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

    /// <summary>
    /// Like Args(), but prepends `ref ` to positional arguments whose corresponding
    /// parameter is declared ByRef, using the supplied mode list.
    /// </summary>
    private string ArgsWithRef(IReadOnlyList<ArgumentNode> args,
        IReadOnlyList<Parsing.Nodes.ParameterMode>? modes)
    {
        if (modes == null) return Args(args);
        var parts = new List<string>(args.Count);
        for (int i = 0; i < args.Count; i++)
        {
            var a = args[i];
            if (a.IsMissing) { parts.Add("/* missing */"); continue; }
            string expr = a.Name != null ? $"/* {a.Name}: */ {Expr(a.Value!)}" : Expr(a.Value!);
            bool isByRef = i < modes.Count && modes[i] == Parsing.Nodes.ParameterMode.ByRef;
            parts.Add(isByRef ? $"ref {expr}" : expr);
        }
        return string.Join(", ", parts);
    }

    /// <summary>
    /// Tries to resolve the parameter mode list for a call target by looking up the
    /// method name and (optionally) the qualifying module name in _methodParamMap.
    /// Returns null when the target cannot be resolved or the map is absent.
    /// </summary>
    private IReadOnlyList<Parsing.Nodes.ParameterMode>? TryResolveProcParamModes(ExpressionNode target)
    {
        if (_methodParamMap == null) return null;

        // Simple call: Foo(...)
        if (target is IdentifierNode id)
        {
            // Check current module first, then all modules (for same-module calls without qualifier).
            if (_methodParamMap.TryGetValue(_currentModuleName, out var selfMethods) &&
                selfMethods.TryGetValue(id.Name, out var modes0))
                return modes0;
            foreach (var module in _methodParamMap.Values)
                if (module.TryGetValue(id.Name, out var modes1))
                    return modes1;
            return null;
        }

        // Qualified call: Obj.Method(...) or ModuleName.Method(...)
        if (target is MemberAccessNode ma)
        {
            // Direct lookup by qualifier string (works when qualifier is the class/module name itself).
            string qualifier = ma.Object is IdentifierNode qid ? qid.Name : Expr(ma.Object);
            if (_methodParamMap.TryGetValue(qualifier, out var methods1) &&
                methods1.TryGetValue(ma.MemberName, out var modes2))
                return modes2;

            // The qualifier is likely a variable name (e.g. `M46V999` of type `clsM46V999`).
            // Look up its declared type (local, module field, or global) and retry.
            if (ma.Object is IdentifierNode varId)
            {
                string? varType = _procTypes.TryGetValue(varId.Name, out var t1) ? t1
                                : _moduleFieldTypes.TryGetValue(varId.Name, out var t2) ? t2
                                : (_globalVarMap != null && _globalVarMap.TryGetValue(varId.Name, out var t3)) ? t3
                                : null;
                if (varType != null &&
                    _methodParamMap.TryGetValue(varType, out var methods2) &&
                    methods2.TryGetValue(ma.MemberName, out var modes3))
                    return modes3;
            }
        }

        return null;
    }

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

    /// <summary>
    /// Returns the appropriate default initializer expression for a declared local.
    /// String → "" (VB6 strings default to empty, not null).
    /// Everything else → default (0 for value types, null for reference types).
    /// </summary>
    private static string DefaultInit(string typeName) =>
        typeName == "string" ? "\"\"" : "default";

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
