namespace VB6toCS.Core.Lexing;

/// <summary>
/// A single lexical token produced by the VB6 lexer.
/// </summary>
public readonly struct Token
{
    public TokenKind Kind { get; }
    public string Text { get; }       // raw source text of this token
    public int Line { get; }          // 1-based line number
    public int Column { get; }        // 1-based column number

    public Token(TokenKind kind, string text, int line, int column)
    {
        Kind = kind;
        Text = text;
        Line = line;
        Column = column;
    }

    public bool IsKeyword => Kind >= TokenKind.KwPublic && Kind <= TokenKind.KwMultiUse;

    public override string ToString() => $"[{Kind} \"{Text}\" {Line}:{Column}]";
}
