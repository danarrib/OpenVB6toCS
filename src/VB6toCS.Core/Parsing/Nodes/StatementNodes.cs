namespace VB6toCS.Core.Parsing.Nodes;

public enum DoLoopKind { DoWhileTop, DoUntilTop, DoWhileBottom, DoUntilBottom, DoForever }
public enum OnErrorKind { GoTo, ResumeNext, GoToZero }

public sealed record LocalDimNode(
    int Line,
    int Column,
    bool IsStatic,
    IReadOnlyList<VariableDeclaratorNode> Declarators) : AstNode(Line, Column);

public sealed record ReDimNode(
    int Line,
    int Column,
    bool IsPreserve,
    IReadOnlyList<VariableDeclaratorNode> Declarators) : AstNode(Line, Column)
{
    /// <summary>
    /// Dimension expressions per declarator (same order and count as <see cref="Declarators"/>).
    /// Each inner list contains the upper-bound expressions for that array variable.
    /// Empty outer list means no dimension information was available (legacy / fallback).
    /// </summary>
    public IReadOnlyList<IReadOnlyList<ExpressionNode>> DimensionLists { get; init; } =
        Array.Empty<IReadOnlyList<ExpressionNode>>();
}

/// <summary>VB6 'Error errornumber' statement — raises a runtime error.</summary>
public sealed record ErrorStatementNode(
    int Line,
    int Column,
    ExpressionNode ErrorNumber) : AstNode(Line, Column);

/// <summary>
/// Return-value assignment inside a Function or Property Get body:
/// 'FunctionName = expr'. Produced by Stage 3 semantic analysis.
/// The code generator emits this as 'return expr;' (simple case) or assigns
/// to a '_result' local and returns it at the end of the method.
/// </summary>
public sealed record FunctionReturnNode(
    int Line,
    int Column,
    ExpressionNode Value) : AstNode(Line, Column);

public sealed record AssignmentNode(
    int Line,
    int Column,
    bool IsSet,
    bool IsLet,
    ExpressionNode Target,
    ExpressionNode Value) : AstNode(Line, Column);

public sealed record CallStatementNode(
    int Line,
    int Column,
    ExpressionNode Target,
    IReadOnlyList<ArgumentNode> Arguments,
    bool ExplicitCall) : AstNode(Line, Column);

public sealed record IfNode(
    int Line,
    int Column,
    ExpressionNode Condition,
    IReadOnlyList<AstNode> ThenBody,
    IReadOnlyList<ElseIfClauseNode> ElseIfClauses,
    IReadOnlyList<AstNode>? ElseBody,
    CommentNode? ThenComment = null) : AstNode(Line, Column);

public sealed record ElseIfClauseNode(
    int Line,
    int Column,
    ExpressionNode Condition,
    IReadOnlyList<AstNode> Body) : AstNode(Line, Column);

public sealed record SingleLineIfNode(
    int Line,
    int Column,
    ExpressionNode Condition,
    AstNode ThenStatement,
    AstNode? ElseStatement) : AstNode(Line, Column);

public sealed record SelectCaseNode(
    int Line,
    int Column,
    ExpressionNode TestExpression,
    IReadOnlyList<CaseClauseNode> Cases) : AstNode(Line, Column);

public sealed record CaseClauseNode(
    int Line,
    int Column,
    IReadOnlyList<CasePatternNode> Patterns,
    bool IsElse,
    IReadOnlyList<AstNode> Body) : AstNode(Line, Column);

public sealed record CaseValuePattern(
    int Line,
    int Column,
    ExpressionNode Value) : CasePatternNode(Line, Column);

public sealed record CaseRangePattern(
    int Line,
    int Column,
    ExpressionNode Low,
    ExpressionNode High) : CasePatternNode(Line, Column);

public sealed record CaseIsPattern(
    int Line,
    int Column,
    VB6toCS.Core.Lexing.TokenKind Operator,
    ExpressionNode Value) : CasePatternNode(Line, Column);

public sealed record ForNextNode(
    int Line,
    int Column,
    string VariableName,
    ExpressionNode Start,
    ExpressionNode End,
    ExpressionNode? Step,
    IReadOnlyList<AstNode> Body) : AstNode(Line, Column);

public sealed record ForEachNode(
    int Line,
    int Column,
    string VariableName,
    ExpressionNode Collection,
    IReadOnlyList<AstNode> Body) : AstNode(Line, Column);

public sealed record WhileNode(
    int Line,
    int Column,
    ExpressionNode Condition,
    IReadOnlyList<AstNode> Body) : AstNode(Line, Column);

public sealed record DoLoopNode(
    int Line,
    int Column,
    DoLoopKind Kind,
    ExpressionNode? Condition,
    IReadOnlyList<AstNode> Body) : AstNode(Line, Column);

public sealed record WithNode(
    int Line,
    int Column,
    ExpressionNode Object,
    IReadOnlyList<AstNode> Body) : AstNode(Line, Column);

public sealed record OnErrorNode(
    int Line,
    int Column,
    OnErrorKind Kind,
    string? LabelName) : AstNode(Line, Column);

public sealed record ResumeNode(
    int Line,
    int Column,
    bool IsNext,
    string? LabelName) : AstNode(Line, Column);

public sealed record GoToNode(
    int Line,
    int Column,
    string Label) : AstNode(Line, Column);

public sealed record GoSubNode(
    int Line,
    int Column,
    string Label) : AstNode(Line, Column);

public sealed record ReturnNode(int Line, int Column) : AstNode(Line, Column);

public sealed record ExitNode(
    int Line,
    int Column,
    string What) : AstNode(Line, Column);

public sealed record LabelNode(
    int Line,
    int Column,
    string Name) : AstNode(Line, Column);

public sealed record CommentNode(
    int Line,
    int Column,
    string Text) : AstNode(Line, Column);

public sealed record EndStatementNode(int Line, int Column) : AstNode(Line, Column);

/// <summary>
/// Structured try/catch block produced by Stage 4 from the VB6 pattern:
///   On Error GoTo label  /  [guarded body]  /  Exit Sub|Function  /  label:  /  [handler body]
/// The CatchVariable name ("_ex") is available to the code generator for the catch parameter.
/// </summary>
public sealed record TryCatchNode(
    int Line,
    int Column,
    IReadOnlyList<AstNode> TryBody,
    string CatchVariable,
    IReadOnlyList<AstNode> CatchBody) : AstNode(Line, Column);
