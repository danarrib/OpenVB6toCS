using VB6toCS.Core.Lexing;

namespace VB6toCS.Core.Parsing.Nodes;

public sealed record IntegerLiteralNode(
    int Line,
    int Column,
    string RawText,
    long Value) : ExpressionNode(Line, Column);

public sealed record DoubleLiteralNode(
    int Line,
    int Column,
    string RawText,
    double Value) : ExpressionNode(Line, Column);

public sealed record StringLiteralNode(
    int Line,
    int Column,
    string RawText,
    string Value) : ExpressionNode(Line, Column);

public sealed record BoolLiteralNode(
    int Line,
    int Column,
    bool Value) : ExpressionNode(Line, Column);

public sealed record DateLiteralNode(
    int Line,
    int Column,
    string RawText) : ExpressionNode(Line, Column);

public sealed record NothingNode(int Line, int Column) : ExpressionNode(Line, Column);

public sealed record IdentifierNode(
    int Line,
    int Column,
    string Name) : ExpressionNode(Line, Column);

public sealed record MeNode(int Line, int Column) : ExpressionNode(Line, Column);

public sealed record BinaryExpressionNode(
    int Line,
    int Column,
    ExpressionNode Left,
    TokenKind Operator,
    ExpressionNode Right) : ExpressionNode(Line, Column);

public sealed record UnaryExpressionNode(
    int Line,
    int Column,
    TokenKind Operator,
    ExpressionNode Operand) : ExpressionNode(Line, Column);

public sealed record MemberAccessNode(
    int Line,
    int Column,
    ExpressionNode Object,
    string MemberName) : ExpressionNode(Line, Column);

public sealed record BangAccessNode(
    int Line,
    int Column,
    ExpressionNode Object,
    string MemberName) : ExpressionNode(Line, Column);

/// <summary>.Member inside a With block — no object prefix.</summary>
public sealed record WithMemberAccessNode(
    int Line,
    int Column,
    string MemberName) : ExpressionNode(Line, Column);

/// <summary>
/// Array element access: arr(i) or arr(i, j).
/// Produced by Stage 3 when a CallOrIndexNode target is resolved as an array variable.
/// Code generator emits arr[i] / arr[i, j].
/// </summary>
public sealed record IndexNode(
    int Line,
    int Column,
    ExpressionNode Target,
    IReadOnlyList<ArgumentNode> Arguments) : ExpressionNode(Line, Column);

/// <summary>foo(x, y) — ambiguous call vs. array index; resolved in Stage 3.</summary>
public sealed record CallOrIndexNode(
    int Line,
    int Column,
    ExpressionNode Target,
    IReadOnlyList<ArgumentNode> Arguments) : ExpressionNode(Line, Column);

public sealed record ArgumentNode(
    int Line,
    int Column,
    string? Name,
    bool IsMissing,
    ExpressionNode? Value) : AstNode(Line, Column);

public sealed record NewObjectNode(
    int Line,
    int Column,
    string TypeName) : ExpressionNode(Line, Column);

public sealed record TypeOfIsNode(
    int Line,
    int Column,
    ExpressionNode Operand,
    string TypeName) : ExpressionNode(Line, Column);
