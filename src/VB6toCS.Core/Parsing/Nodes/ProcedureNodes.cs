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

/// <summary>
/// A C#-style property grouping one or more of: Property Get / Let / Set.
/// Produced by Stage 3 semantic analysis; replaces the separate PropertyNode members.
/// Type is taken from the Get return type, or from the value parameter of Let/Set.
/// </summary>
public sealed record CsPropertyNode(
    int Line,
    int Column,
    AccessModifier Access,
    bool IsStatic,
    string Name,
    TypeRefNode? Type,
    // Get accessor — null if no Property Get declared
    IReadOnlyList<ParameterNode> GetParameters,
    IReadOnlyList<AstNode>? GetBody,
    // Let accessor — null if no Property Let declared
    IReadOnlyList<ParameterNode> LetParameters,
    IReadOnlyList<AstNode>? LetBody,
    // Set accessor — null if no Property Set declared
    IReadOnlyList<ParameterNode> SetParameters,
    IReadOnlyList<AstNode>? SetBody) : AstNode(Line, Column);
