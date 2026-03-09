using VB6toCS.Core.Parsing.Nodes;

namespace VB6toCS.Core.Parsing;

/// <summary>
/// Prints an AST as an indented tree to a <see cref="TextWriter"/>.
/// </summary>
public static class AstPrinter
{
    public static void Print(ModuleNode module, TextWriter writer)
    {
        string defMember = module.DefaultMemberName != null ? $" [Default: {module.DefaultMemberName}]" : "";
        writer.WriteLine($"ModuleNode '{module.Name}' [{module.Kind}]{defMember}");
        foreach (var impl in module.Implements)
            writer.WriteLine($"  Implements {impl}");
        PrintNodes(module.Members, writer, "  ");
    }

    private static void PrintNodes(IReadOnlyList<AstNode> nodes, TextWriter writer, string indent)
    {
        foreach (var node in nodes)
            PrintNode(node, writer, indent);
    }

    private static void PrintNode(AstNode node, TextWriter writer, string indent)
    {
        switch (node)
        {
            case CommentNode c:
                writer.WriteLine($"{indent}' {c.Text.TrimStart('\'').Trim()}");
                break;

            case FieldNode f:
                writer.WriteLine($"{indent}FieldNode [{f.Access}]");
                foreach (var d in f.Declarators) PrintDeclarator(d, writer, indent + "  ");
                break;

            case ConstDeclarationNode c:
                writer.WriteLine($"{indent}ConstDeclarationNode [{c.Access}]");
                foreach (var d in c.Declarators) PrintDeclarator(d, writer, indent + "  ");
                break;

            case LocalDimNode d:
                writer.WriteLine($"{indent}LocalDimNode{(d.IsStatic ? " [Static]" : "")}");
                foreach (var dec in d.Declarators) PrintDeclarator(dec, writer, indent + "  ");
                break;

            case ReDimNode r:
                writer.WriteLine($"{indent}ReDimNode{(r.IsPreserve ? " [Preserve]" : "")}");
                foreach (var dec in r.Declarators) PrintDeclarator(dec, writer, indent + "  ");
                break;

            case EnumNode e:
                writer.WriteLine($"{indent}EnumNode '{e.Name}' [{e.Access}]");
                foreach (var m in e.Members)
                    writer.WriteLine($"{indent}  EnumMemberNode '{m.Name}'" +
                        (m.Value != null ? $" = {ExprStr(m.Value)}" : ""));
                break;

            case UdtNode u:
                writer.WriteLine($"{indent}UdtNode '{u.Name}' [{u.Access}]");
                foreach (var f in u.Fields)
                    writer.WriteLine($"{indent}  UdtFieldNode '{f.Name}' As {f.TypeRef.TypeName}");
                break;

            case DeclareNode d:
                var dKind = d.IsSub ? "Sub" : "Function";
                var dAlias = d.AliasName != null ? $" Alias \"{d.AliasName}\"" : "";
                var dRet = d.ReturnType != null ? $" As {d.ReturnType.TypeName}" : "";
                writer.WriteLine($"{indent}DeclareNode [{d.Access}] {dKind} '{d.Name}' Lib \"{d.LibName}\"{dAlias}{dRet}");
                break;

            case SubNode s:
                writer.WriteLine($"{indent}SubNode '{s.Name}' [{s.Access}]{(s.IsStatic ? " [Static]" : "")}");
                foreach (var p in s.Parameters) PrintParam(p, writer, indent + "  ");
                PrintNodes(s.Body, writer, indent + "  ");
                break;

            case FunctionNode f:
                var retType = f.ReturnType != null ? $" As {f.ReturnType.TypeName}" : "";
                writer.WriteLine($"{indent}FunctionNode '{f.Name}' [{f.Access}]{(f.IsStatic ? " [Static]" : "")}{retType}");
                foreach (var p in f.Parameters) PrintParam(p, writer, indent + "  ");
                PrintNodes(f.Body, writer, indent + "  ");
                break;

            case PropertyNode p:
                var pRetType = p.ReturnType != null ? $" As {p.ReturnType.TypeName}" : "";
                writer.WriteLine($"{indent}PropertyNode '{p.Name}' [{p.Kind}] [{p.Access}]{pRetType}");
                foreach (var pm in p.Parameters) PrintParam(pm, writer, indent + "  ");
                PrintNodes(p.Body, writer, indent + "  ");
                break;

            case CsPropertyNode cp:
                var cpType = cp.Type != null ? $" As {cp.Type.TypeName}" : "";
                writer.WriteLine($"{indent}CsPropertyNode '{cp.Name}' [{cp.Access}]{(cp.IsStatic ? " [Static]" : "")}{cpType}");
                if (cp.GetBody != null)
                {
                    writer.WriteLine($"{indent}  Get:");
                    foreach (var pm in cp.GetParameters) PrintParam(pm, writer, indent + "    ");
                    PrintNodes(cp.GetBody, writer, indent + "    ");
                }
                if (cp.LetBody != null)
                {
                    writer.WriteLine($"{indent}  Let:");
                    foreach (var pm in cp.LetParameters) PrintParam(pm, writer, indent + "    ");
                    PrintNodes(cp.LetBody, writer, indent + "    ");
                }
                if (cp.SetBody != null)
                {
                    writer.WriteLine($"{indent}  Set:");
                    foreach (var pm in cp.SetParameters) PrintParam(pm, writer, indent + "    ");
                    PrintNodes(cp.SetBody, writer, indent + "    ");
                }
                break;

            case AssignmentNode a:
                var prefix = a.IsSet ? "Set " : a.IsLet ? "Let " : "";
                writer.WriteLine($"{indent}AssignmentNode {prefix}{ExprStr(a.Target)} = {ExprStr(a.Value)}");
                break;

            case CallStatementNode c:
                writer.WriteLine($"{indent}CallStatementNode {ExprStr(c.Target)}" +
                    (c.Arguments.Count > 0 ? $"({string.Join(", ", c.Arguments.Select(ArgStr))})" : ""));
                break;

            case IfNode i:
                writer.WriteLine($"{indent}IfNode {ExprStr(i.Condition)}");
                writer.WriteLine($"{indent}  ThenBody:");
                PrintNodes(i.ThenBody, writer, indent + "    ");
                foreach (var ei in i.ElseIfClauses)
                {
                    writer.WriteLine($"{indent}  ElseIf {ExprStr(ei.Condition)}");
                    PrintNodes(ei.Body, writer, indent + "    ");
                }
                if (i.ElseBody != null)
                {
                    writer.WriteLine($"{indent}  ElseBody:");
                    PrintNodes(i.ElseBody, writer, indent + "    ");
                }
                break;

            case SingleLineIfNode si:
                writer.WriteLine($"{indent}SingleLineIfNode {ExprStr(si.Condition)}");
                PrintNode(si.ThenStatement, writer, indent + "  Then: ");
                if (si.ElseStatement != null)
                    PrintNode(si.ElseStatement, writer, indent + "  Else: ");
                break;

            case SelectCaseNode s:
                writer.WriteLine($"{indent}SelectCaseNode {ExprStr(s.TestExpression)}");
                foreach (var c in s.Cases)
                {
                    if (c.IsElse)
                        writer.WriteLine($"{indent}  Case Else");
                    else
                        writer.WriteLine($"{indent}  Case {string.Join(", ", c.Patterns.Select(PatStr))}");
                    PrintNodes(c.Body, writer, indent + "    ");
                }
                break;

            case ForNextNode f:
                writer.WriteLine($"{indent}ForNextNode '{f.VariableName}' = {ExprStr(f.Start)} To {ExprStr(f.End)}" +
                    (f.Step != null ? $" Step {ExprStr(f.Step)}" : ""));
                PrintNodes(f.Body, writer, indent + "  ");
                break;

            case ForEachNode fe:
                writer.WriteLine($"{indent}ForEachNode '{fe.VariableName}' In {ExprStr(fe.Collection)}");
                PrintNodes(fe.Body, writer, indent + "  ");
                break;

            case WhileNode w:
                writer.WriteLine($"{indent}WhileNode {ExprStr(w.Condition)}");
                PrintNodes(w.Body, writer, indent + "  ");
                break;

            case DoLoopNode d:
                writer.WriteLine($"{indent}DoLoopNode [{d.Kind}]" +
                    (d.Condition != null ? $" {ExprStr(d.Condition)}" : ""));
                PrintNodes(d.Body, writer, indent + "  ");
                break;

            case WithNode w:
                writer.WriteLine($"{indent}WithNode {ExprStr(w.Object)}");
                PrintNodes(w.Body, writer, indent + "  ");
                break;

            case OnErrorNode o:
                writer.WriteLine($"{indent}OnErrorNode [{o.Kind}]" + (o.LabelName != null ? $" '{o.LabelName}'" : ""));
                break;

            case ResumeNode r:
                writer.WriteLine($"{indent}ResumeNode{(r.IsNext ? " Next" : r.LabelName != null ? $" '{r.LabelName}'" : "")}");
                break;

            case GoToNode g:
                writer.WriteLine($"{indent}GoToNode '{g.Label}'");
                break;

            case GoSubNode g:
                writer.WriteLine($"{indent}GoSubNode '{g.Label}'");
                break;

            case ReturnNode:
                writer.WriteLine($"{indent}ReturnNode");
                break;

            case ExitNode e:
                writer.WriteLine($"{indent}ExitNode {e.What}");
                break;

            case LabelNode l:
                writer.WriteLine($"{indent}LabelNode '{l.Name}':");
                break;

            case EndStatementNode:
                writer.WriteLine($"{indent}EndStatementNode");
                break;

            case TryCatchNode tc:
                writer.WriteLine($"{indent}TryCatchNode");
                writer.WriteLine($"{indent}  Try:");
                PrintNodes(tc.TryBody, writer, indent + "    ");
                writer.WriteLine($"{indent}  Catch ({tc.CatchVariable}):");
                PrintNodes(tc.CatchBody, writer, indent + "    ");
                if (tc.FinallyBody is { Count: > 0 })
                {
                    writer.WriteLine($"{indent}  Finally:");
                    PrintNodes(tc.FinallyBody, writer, indent + "    ");
                }
                break;

            case ErrorStatementNode e:
                writer.WriteLine($"{indent}ErrorStatementNode Error {ExprStr(e.ErrorNumber)}");
                break;

            case FunctionReturnNode r:
                writer.WriteLine($"{indent}FunctionReturnNode = {ExprStr(r.Value)}");
                break;

            default:
                writer.WriteLine($"{indent}{node.GetType().Name}");
                break;
        }
    }

    private static void PrintDeclarator(VariableDeclaratorNode d, TextWriter writer, string indent)
    {
        var typeStr = d.TypeRef != null ? $" As {d.TypeRef.TypeName}{(d.TypeRef.IsArray ? "()" : "")}" : "";
        var valStr = d.DefaultValue != null ? $" = {ExprStr(d.DefaultValue)}" : "";
        writer.WriteLine($"{indent}VariableDeclaratorNode '{d.Name}'{typeStr}{valStr}");
    }

    private static void PrintParam(ParameterNode p, TextWriter writer, string indent)
    {
        var parts = new List<string>();
        if (p.IsOptional) parts.Add("Optional");
        if (p.IsParamArray) parts.Add("ParamArray");
        if (p.Mode != ParameterMode.Unspecified) parts.Add(p.Mode.ToString());
        if (p.TypeRef != null) parts.Add($"As {p.TypeRef.TypeName}");
        var suffix = parts.Count > 0 ? " " + string.Join(" ", parts) : "";
        writer.WriteLine($"{indent}ParameterNode '{p.Name}'{suffix}");
    }

    private static string ExprStr(ExpressionNode e) => e switch
    {
        IntegerLiteralNode i => i.RawText,
        DoubleLiteralNode d => d.RawText,
        StringLiteralNode s => s.RawText,
        BoolLiteralNode b => b.Value.ToString(),
        DateLiteralNode d => d.RawText,
        NothingNode => "Nothing",
        IdentifierNode id => id.Name,
        MeNode => "Me",
        BinaryExpressionNode b => $"({ExprStr(b.Left)} {b.Operator} {ExprStr(b.Right)})",
        UnaryExpressionNode u => $"({u.Operator} {ExprStr(u.Operand)})",
        MemberAccessNode m => $"{ExprStr(m.Object)}.{m.MemberName}",
        BangAccessNode b => $"{ExprStr(b.Object)}!{b.MemberName}",
        WithMemberAccessNode w => $".{w.MemberName}",
        CallOrIndexNode c => $"{ExprStr(c.Target)}({string.Join(", ", c.Arguments.Select(ArgStr))})",
        IndexNode ix => $"{ExprStr(ix.Target)}[{string.Join(", ", ix.Arguments.Select(ArgStr))}]",
        NewObjectNode n => $"New {n.TypeName}",
        TypeOfIsNode t => $"TypeOf {ExprStr(t.Operand)} Is {t.TypeName}",
        _ => e.GetType().Name
    };

    private static string ArgStr(ArgumentNode a)
    {
        if (a.IsMissing) return "<missing>";
        if (a.Name != null) return $"{a.Name}:={ExprStr(a.Value!)}";
        return ExprStr(a.Value!);
    }

    private static string PatStr(CasePatternNode p) => p switch
    {
        CaseValuePattern v => ExprStr(v.Value),
        CaseRangePattern r => $"{ExprStr(r.Low)} To {ExprStr(r.High)}",
        CaseIsPattern i => $"Is {i.Operator} {ExprStr(i.Value)}",
        _ => p.GetType().Name
    };
}
