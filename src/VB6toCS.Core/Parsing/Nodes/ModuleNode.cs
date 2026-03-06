namespace VB6toCS.Core.Parsing.Nodes;

public sealed record ModuleNode(
    int Line,
    int Column,
    ModuleKind Kind,
    string Name,
    IReadOnlyList<string> Implements,
    IReadOnlyList<AstNode> Members) : AstNode(Line, Column);
