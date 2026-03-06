namespace VB6toCS.Core.Parsing.Nodes;

public enum ParameterMode { ByVal, ByRef, Unspecified }
public enum AccessModifier { Public, Private, Friend, Global }

public sealed record ParameterNode(
    int Line,
    int Column,
    ParameterMode Mode,
    bool IsOptional,
    bool IsParamArray,
    string Name,
    TypeRefNode? TypeRef,
    ExpressionNode? DefaultValue) : AstNode(Line, Column);
