namespace VB6toCS.Core.Parsing.Nodes;

public sealed record FieldNode(
    int Line,
    int Column,
    AccessModifier Access,
    IReadOnlyList<VariableDeclaratorNode> Declarators) : AstNode(Line, Column);

public sealed record VariableDeclaratorNode(
    int Line,
    int Column,
    string Name,
    TypeRefNode? TypeRef,
    ExpressionNode? DefaultValue) : AstNode(Line, Column);

public sealed record ConstDeclarationNode(
    int Line,
    int Column,
    AccessModifier Access,
    IReadOnlyList<VariableDeclaratorNode> Declarators) : AstNode(Line, Column);

public sealed record EnumNode(
    int Line,
    int Column,
    AccessModifier Access,
    string Name,
    IReadOnlyList<EnumMemberNode> Members) : AstNode(Line, Column);

public sealed record EnumMemberNode(
    int Line,
    int Column,
    string Name,
    ExpressionNode? Value,
    IReadOnlyList<CommentNode> TrailingComments) : AstNode(Line, Column);

public sealed record UdtNode(
    int Line,
    int Column,
    AccessModifier Access,
    string Name,
    IReadOnlyList<UdtFieldNode> Fields) : AstNode(Line, Column);

public sealed record UdtFieldNode(
    int Line,
    int Column,
    string Name,
    TypeRefNode TypeRef) : AstNode(Line, Column);

public sealed record ImplementsNode(
    int Line,
    int Column,
    string InterfaceName) : AstNode(Line, Column);

/// <summary>
/// VB6 Declare statement — external DLL function (e.g. Win32 API).
/// Preserved in the AST for reference; code generator will emit a P/Invoke stub
/// or a // TODO comment.
/// </summary>
public sealed record DeclareNode(
    int Line,
    int Column,
    AccessModifier Access,
    bool IsSub,
    string Name,
    string LibName,
    string? AliasName,
    IReadOnlyList<ParameterNode> Parameters,
    TypeRefNode? ReturnType) : AstNode(Line, Column);
