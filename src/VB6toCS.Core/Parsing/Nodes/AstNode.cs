namespace VB6toCS.Core.Parsing.Nodes;

public abstract record AstNode(int Line, int Column);
public abstract record ExpressionNode(int Line, int Column) : AstNode(Line, Column);
public abstract record CasePatternNode(int Line, int Column) : AstNode(Line, Column);
