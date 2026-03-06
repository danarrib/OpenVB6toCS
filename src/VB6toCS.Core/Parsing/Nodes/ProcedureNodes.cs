namespace VB6toCS.Core.Parsing.Nodes;

public enum PropertyKind { Get, Let, Set }

public sealed record SubNode(
    int Line,
    int Column,
    AccessModifier Access,
    bool IsStatic,
    string Name,
    IReadOnlyList<ParameterNode> Parameters,
    IReadOnlyList<AstNode> Body) : AstNode(Line, Column);

public sealed record FunctionNode(
    int Line,
    int Column,
    AccessModifier Access,
    bool IsStatic,
    string Name,
    IReadOnlyList<ParameterNode> Parameters,
    TypeRefNode? ReturnType,
    IReadOnlyList<AstNode> Body) : AstNode(Line, Column);

public sealed record PropertyNode(
    int Line,
    int Column,
    AccessModifier Access,
    bool IsStatic,
    PropertyKind Kind,
    string Name,
    IReadOnlyList<ParameterNode> Parameters,
    TypeRefNode? ReturnType,
    IReadOnlyList<AstNode> Body) : AstNode(Line, Column);
