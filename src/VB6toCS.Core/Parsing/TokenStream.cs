using VB6toCS.Core.Lexing;

namespace VB6toCS.Core.Parsing;

public sealed class TokenStream
{
    private readonly IReadOnlyList<Token> _tokens;
    private int _pos;

    public TokenStream(IReadOnlyList<Token> tokens)
    {
        _tokens = tokens;
        _pos = 0;
    }

    /// <summary>Current token without consuming.</summary>
    public Token Peek() => _pos < _tokens.Count ? _tokens[_pos] : Eof();

    /// <summary>Look ahead by <paramref name="offset"/> positions without consuming.</summary>
    public Token PeekAt(int offset)
    {
        int idx = _pos + offset;
        return idx < _tokens.Count ? _tokens[idx] : Eof();
    }

    /// <summary>Advance and return the current token.</summary>
    public Token Consume()
    {
        Token t = Peek();
        if (_pos < _tokens.Count) _pos++;
        return t;
    }

    /// <summary>Consume if the current token matches <paramref name="kind"/>. Returns true if consumed.</summary>
    public bool Match(TokenKind kind)
    {
        if (Peek().Kind != kind) return false;
        _pos++;
        return true;
    }

    /// <summary>Check current token kind without consuming.</summary>
    public bool Check(TokenKind kind) => Peek().Kind == kind;

    /// <summary>Consume the current token and return it; throw if kind doesn't match.</summary>
    public Token Expect(TokenKind kind)
    {
        Token t = Peek();
        if (t.Kind != kind)
            throw new ParseException(
                $"Expected {kind} but found {t.Kind} ('{t.Text}')", t.Line, t.Column);
        _pos++;
        return t;
    }

    /// <summary>Skip Newline and Colon tokens.</summary>
    public void SkipNewlines()
    {
        while (Peek().Kind is TokenKind.Newline or TokenKind.Colon)
            _pos++;
    }

    /// <summary>Consume all remaining tokens on this line (used for attribute/header lines).</summary>
    public void SkipToEndOfLine()
    {
        while (Peek().Kind is not (TokenKind.Newline or TokenKind.EndOfFile))
            _pos++;
        Match(TokenKind.Newline);
    }

    /// <summary>
    /// When true, <see cref="AtEndOfStatement"/> and <see cref="ExpectEndOfStatement"/>
    /// also treat <c>KwElse</c> as an end of statement (for single-line If bodies).
    /// </summary>
    public bool SingleLineIfMode { get; set; }

    /// <summary>True if we are at the end of a logical statement (ignoring trailing comments).</summary>
    public bool AtEndOfStatement()
    {
        int i = 0;
        while (PeekAt(i).Kind == TokenKind.Comment) i++;
        var k = PeekAt(i).Kind;
        return k is TokenKind.Newline or TokenKind.Colon or TokenKind.EndOfFile
               || (SingleLineIfMode && k == TokenKind.KwElse);
    }

    /// <summary>
    /// If the current token is a comment, consume and return it as a <see cref="Nodes.CommentNode"/>.
    /// Returns null if there is no trailing comment.
    /// </summary>
    public Nodes.CommentNode? ConsumeTrailingComment()
    {
        if (Peek().Kind != TokenKind.Comment) return null;
        var t = Consume();
        return new Nodes.CommentNode(t.Line, t.Column, t.Text);
    }

    /// <summary>Consume trailing comments then one end-of-statement token or throw.</summary>
    public void ExpectEndOfStatement()
    {
        // Skip any trailing comment on the same line
        while (Peek().Kind == TokenKind.Comment) _pos++;

        Token t = Peek();
        if (t.Kind is TokenKind.Newline or TokenKind.Colon)
            _pos++;
        else if (t.Kind == TokenKind.EndOfFile)
            return;
        else if (SingleLineIfMode && t.Kind == TokenKind.KwElse)
            return; // don't consume Else — caller will check for it
        else
            throw new ParseException(
                $"Expected end of statement but found {t.Kind} ('{t.Text}')", t.Line, t.Column);
    }

    private Token Eof()
    {
        // Return a synthetic EOF token positioned after the last real token
        if (_tokens.Count > 0)
        {
            var last = _tokens[_tokens.Count - 1];
            return new Token(TokenKind.EndOfFile, "", last.Line, last.Column + 1);
        }
        return new Token(TokenKind.EndOfFile, "", 1, 1);
    }
}
