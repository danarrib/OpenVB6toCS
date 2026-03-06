namespace VB6toCS.Core.Parsing.Nodes;

/// <summary>Reference to a VB6 type (As Double, As String, As MyClass, etc.).</summary>
public sealed record TypeRefNode(
    string TypeName,
    bool IsArray,
    bool IsNew,
    int ArrayRank) : AstNode(0, 0);
