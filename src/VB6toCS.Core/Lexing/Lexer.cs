using System.Text;

namespace VB6toCS.Core.Lexing;

/// <summary>
/// Hand-written lexer for Visual Basic 6 source files.
/// Produces a flat list of tokens; line-continuation sequences (" _\r\n") are
/// silently consumed so the parser never sees them.
/// </summary>
public sealed class Lexer
{
    private readonly string _src;
    private int _pos;
    private int _line;
    private int _col;

    // ── Keyword table (case-insensitive) ──────────────────────────────────
    private static readonly Dictionary<string, TokenKind> Keywords =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ["public"]      = TokenKind.KwPublic,
            ["private"]     = TokenKind.KwPrivate,
            ["friend"]      = TokenKind.KwFriend,
            ["global"]      = TokenKind.KwGlobal,
            ["dim"]         = TokenKind.KwDim,
            ["const"]       = TokenKind.KwConst,
            ["static"]      = TokenKind.KwStatic,
            ["as"]          = TokenKind.KwAs,
            ["new"]         = TokenKind.KwNew,
            ["set"]         = TokenKind.KwSet,
            ["let"]         = TokenKind.KwLet,
            ["call"]        = TokenKind.KwCall,
            ["option"]      = TokenKind.KwOption,
            ["explicit"]    = TokenKind.KwExplicit,
            ["attribute"]   = TokenKind.KwAttribute,
            ["enum"]        = TokenKind.KwEnum,
            ["type"]        = TokenKind.KwType,
            ["implements"]  = TokenKind.KwImplements,
            ["function"]    = TokenKind.KwFunction,
            ["sub"]         = TokenKind.KwSub,
            ["property"]    = TokenKind.KwProperty,
            ["get"]         = TokenKind.KwGet,
            ["end"]         = TokenKind.KwEnd,
            ["exit"]        = TokenKind.KwExit,
            ["byval"]       = TokenKind.KwByVal,
            ["byref"]       = TokenKind.KwByRef,
            ["optional"]    = TokenKind.KwOptional,
            ["paramarray"]  = TokenKind.KwParamArray,
            ["if"]          = TokenKind.KwIf,
            ["then"]        = TokenKind.KwThen,
            ["else"]        = TokenKind.KwElse,
            ["elseif"]      = TokenKind.KwElseIf,
            ["select"]      = TokenKind.KwSelect,
            ["case"]        = TokenKind.KwCase,
            ["for"]         = TokenKind.KwFor,
            ["each"]        = TokenKind.KwEach,
            ["in"]          = TokenKind.KwIn,
            ["to"]          = TokenKind.KwTo,
            ["step"]        = TokenKind.KwStep,
            ["next"]        = TokenKind.KwNext,
            ["while"]       = TokenKind.KwWhile,
            ["wend"]        = TokenKind.KwWend,
            ["do"]          = TokenKind.KwDo,
            ["loop"]        = TokenKind.KwLoop,
            ["until"]       = TokenKind.KwUntil,
            ["with"]        = TokenKind.KwWith,
            ["goto"]        = TokenKind.KwGoTo,
            ["gosub"]       = TokenKind.KwGoSub,
            ["return"]      = TokenKind.KwReturn,
            ["on"]          = TokenKind.KwOn,
            ["error"]       = TokenKind.KwError,
            ["resume"]      = TokenKind.KwResume,
            ["nothing"]     = TokenKind.KwNothing,
            ["me"]          = TokenKind.KwMe,
            ["is"]          = TokenKind.KwIs,
            ["like"]        = TokenKind.KwLike,
            ["typeof"]      = TokenKind.KwTypeof,
            ["boolean"]     = TokenKind.KwBoolean,
            ["byte"]        = TokenKind.KwByte,
            ["integer"]     = TokenKind.KwInteger,
            ["long"]        = TokenKind.KwLong,
            ["single"]      = TokenKind.KwSingle,
            ["double"]      = TokenKind.KwDouble,
            ["currency"]    = TokenKind.KwCurrency,
            ["decimal"]     = TokenKind.KwDecimal,
            ["date"]        = TokenKind.KwDate,
            ["string"]      = TokenKind.KwString,
            ["object"]      = TokenKind.KwObject,
            ["variant"]     = TokenKind.KwVariant,
            ["and"]         = TokenKind.KwAnd,
            ["or"]          = TokenKind.KwOr,
            ["not"]         = TokenKind.KwNot,
            ["xor"]         = TokenKind.KwXor,
            ["eqv"]         = TokenKind.KwEqv,
            ["imp"]         = TokenKind.KwImp,
            ["mod"]         = TokenKind.KwMod,
            ["version"]     = TokenKind.KwVersion,
            ["begin"]       = TokenKind.KwBegin,
            ["multiuse"]    = TokenKind.KwMultiUse,
            ["true"]        = TokenKind.BoolLiteral,
            ["false"]       = TokenKind.BoolLiteral,
        };

    public Lexer(string source)
    {
        _src = source;
        _pos = 0;
        _line = 1;
        _col = 1;
    }

    // ── Public API ────────────────────────────────────────────────────────

    public List<Token> Tokenize()
    {
        var tokens = new List<Token>();
        while (true)
        {
            var tok = NextToken();
            tokens.Add(tok);
            if (tok.Kind == TokenKind.EndOfFile) break;
        }
        return tokens;
    }

    // ── Core scanning ─────────────────────────────────────────────────────

    private Token NextToken()
    {
        SkipHorizontalWhitespace();

        if (_pos >= _src.Length)
            return Make(TokenKind.EndOfFile, "");

        int startLine = _line, startCol = _col;
        char c = _src[_pos];

        // Line continuation:  <space>_<newline>  — consume and restart
        if (c == '_' && IsLineContinuation())
        {
            ConsumLineContinuation();
            return NextToken();
        }

        // Newline
        if (c == '\r' || c == '\n')
            return ScanNewline(startLine, startCol);

        // Comment
        if (c == '\'')
            return ScanComment(startLine, startCol);

        // REM comment (only valid as first token on logical line — close enough)
        if (MatchKeywordAhead("rem") && (PeekAt(3) == ' ' || PeekAt(3) == '\t' || PeekAt(3) == '\0'))
            return ScanComment(startLine, startCol);

        // String literal
        if (c == '"')
            return ScanString(startLine, startCol);

        // Date literal
        if (c == '#')
            return ScanDateOrHash(startLine, startCol);

        // Hex / Octal literal  &H...  &O...
        if (c == '&' && _pos + 1 < _src.Length && (_src[_pos + 1] == 'H' || _src[_pos + 1] == 'h' ||
                                                     _src[_pos + 1] == 'O' || _src[_pos + 1] == 'o'))
            return ScanHexOctal(startLine, startCol);

        // Number literal
        if (char.IsDigit(c) || (c == '.' && _pos + 1 < _src.Length && char.IsDigit(_src[_pos + 1])))
            return ScanNumber(startLine, startCol);

        // Identifier or keyword
        if (char.IsLetter(c) || c == '_')
            return ScanIdentifierOrKeyword(startLine, startCol);

        // Operators and punctuation
        return ScanSymbol(startLine, startCol);
    }

    // ── Whitespace / continuation ─────────────────────────────────────────

    private void SkipHorizontalWhitespace()
    {
        while (_pos < _src.Length && (_src[_pos] == ' ' || _src[_pos] == '\t'))
            Advance();
    }

    private bool IsLineContinuation()
    {
        // Pattern: optional spaces, then '_', then optional spaces, then newline
        int i = _pos;
        if (i >= _src.Length || _src[i] != '_') return false;
        i++;
        while (i < _src.Length && (_src[i] == ' ' || _src[i] == '\t')) i++;
        return i < _src.Length && (_src[i] == '\r' || _src[i] == '\n');
    }

    private void ConsumLineContinuation()
    {
        Advance(); // _
        while (_pos < _src.Length && (_src[_pos] == ' ' || _src[_pos] == '\t')) Advance();
        if (_pos < _src.Length && _src[_pos] == '\r') Advance();
        if (_pos < _src.Length && _src[_pos] == '\n') Advance();
        _line++;
        _col = 1;
        SkipHorizontalWhitespace();
    }

    // ── Scanners ──────────────────────────────────────────────────────────

    private Token ScanNewline(int line, int col)
    {
        var sb = new StringBuilder();
        if (_src[_pos] == '\r') { sb.Append(Advance()); }
        if (_pos < _src.Length && _src[_pos] == '\n') { sb.Append(Advance()); }
        _line++;
        _col = 1;
        return new Token(TokenKind.Newline, sb.ToString(), line, col);
    }

    private Token ScanComment(int line, int col)
    {
        var sb = new StringBuilder();
        while (true)
        {
            // Read to end of line
            while (_pos < _src.Length && _src[_pos] != '\r' && _src[_pos] != '\n')
                sb.Append(Advance());

            // VB6 quirk: a comment ending with space+underscore continues on the next line
            string text = sb.ToString();
            if (text.Length >= 2 && text[^1] == '_' && text[^2] == ' '
                && _pos < _src.Length)
            {
                // Consume the newline and continue reading as part of this comment
                if (_src[_pos] == '\r') _pos++;
                if (_pos < _src.Length && _src[_pos] == '\n') { _pos++; _line++; _col = 1; }
                continue;
            }
            break;
        }
        return new Token(TokenKind.Comment, sb.ToString(), line, col);
    }

    private Token ScanString(int line, int col)
    {
        var sb = new StringBuilder();
        sb.Append(Advance()); // opening "
        while (_pos < _src.Length)
        {
            char c = _src[_pos];
            if (c == '"')
            {
                sb.Append(Advance());
                // "" is an escaped quote inside the string
                if (_pos < _src.Length && _src[_pos] == '"')
                    sb.Append(Advance());
                else
                    break;
            }
            else if (c == '\r' || c == '\n')
            {
                break; // unterminated string — let parser report the error
            }
            else
            {
                sb.Append(Advance());
            }
        }
        return new Token(TokenKind.StringLiteral, sb.ToString(), line, col);
    }

    private Token ScanDateOrHash(int line, int col)
    {
        // Try to scan a date literal: #...#
        int save = _pos;
        var sb = new StringBuilder();
        sb.Append(Advance()); // #
        while (_pos < _src.Length && _src[_pos] != '#' && _src[_pos] != '\r' && _src[_pos] != '\n')
            sb.Append(Advance());

        if (_pos < _src.Length && _src[_pos] == '#')
        {
            sb.Append(Advance()); // closing #
            return new Token(TokenKind.DateLiteral, sb.ToString(), line, col);
        }

        // Not a date — treat as bare #
        _pos = save + 1;
        _col = col + 1;
        return new Token(TokenKind.Hash, "#", line, col);
    }

    private Token ScanHexOctal(int line, int col)
    {
        var sb = new StringBuilder();
        sb.Append(Advance()); // &
        sb.Append(Advance()); // H or O
        while (_pos < _src.Length && IsHexChar(_src[_pos]))
            sb.Append(Advance());
        // optional type suffix: &, %
        if (_pos < _src.Length && (_src[_pos] == '&' || _src[_pos] == '%'))
            sb.Append(Advance());
        return new Token(TokenKind.IntegerLiteral, sb.ToString(), line, col);
    }

    private Token ScanNumber(int line, int col)
    {
        var sb = new StringBuilder();
        bool isDouble = false;

        while (_pos < _src.Length && char.IsDigit(_src[_pos]))
            sb.Append(Advance());

        if (_pos < _src.Length && _src[_pos] == '.')
        {
            isDouble = true;
            sb.Append(Advance());
            while (_pos < _src.Length && char.IsDigit(_src[_pos]))
                sb.Append(Advance());
        }

        // Exponent
        if (_pos < _src.Length && (_src[_pos] == 'E' || _src[_pos] == 'e'))
        {
            isDouble = true;
            sb.Append(Advance());
            if (_pos < _src.Length && (_src[_pos] == '+' || _src[_pos] == '-'))
                sb.Append(Advance());
            while (_pos < _src.Length && char.IsDigit(_src[_pos]))
                sb.Append(Advance());
        }

        // Type suffixes: %, &, !, #, @
        if (_pos < _src.Length && "%&!#@".Contains(_src[_pos]))
        {
            char suffix = _src[_pos];
            sb.Append(Advance());
            if (suffix == '!' || suffix == '#' || suffix == '@') isDouble = true;
        }

        return new Token(isDouble ? TokenKind.DoubleLiteral : TokenKind.IntegerLiteral,
                         sb.ToString(), line, col);
    }

    private Token ScanIdentifierOrKeyword(int line, int col)
    {
        var sb = new StringBuilder();
        while (_pos < _src.Length && (char.IsLetterOrDigit(_src[_pos]) || _src[_pos] == '_'))
            sb.Append(Advance());

        // Optional type-declaration character suffix: %, &, !, #, @, $
        // Exception: '!' followed by a letter is bang (default-member) access, e.g. col!Key.
        // In that case we leave the '!' for ScanSymbol to emit as a Bang token.
        if (_pos < _src.Length)
        {
            char suffix = _src[_pos];
            bool isBangAccess = suffix == '!' &&
                                _pos + 1 < _src.Length &&
                                char.IsLetter(_src[_pos + 1]);
            if (!isBangAccess && "%&!#@$".Contains(suffix))
                sb.Append(Advance());
        }

        string text = sb.ToString();

        if (Keywords.TryGetValue(text, out var kw))
            return new Token(kw, text, line, col);

        return new Token(TokenKind.Identifier, text, line, col);
    }

    private Token ScanSymbol(int line, int col)
    {
        char c = Advance();
        switch (c)
        {
            case '+': return Make2(TokenKind.Plus, "+", line, col);
            case '-': return Make2(TokenKind.Minus, "-", line, col);
            case '*': return Make2(TokenKind.Star, "*", line, col);
            case '/': return Make2(TokenKind.Slash, "/", line, col);
            case '\\': return Make2(TokenKind.Backslash, "\\", line, col);
            case '^': return Make2(TokenKind.Caret, "^", line, col);
            case '&': return Make2(TokenKind.Ampersand, "&", line, col);
            case '(': return Make2(TokenKind.LParen, "(", line, col);
            case ')': return Make2(TokenKind.RParen, ")", line, col);
            case ',': return Make2(TokenKind.Comma, ",", line, col);
            case '.': return Make2(TokenKind.Dot, ".", line, col);
            case '!': return Make2(TokenKind.Bang, "!", line, col);
            case '@': return Make2(TokenKind.At, "@", line, col);
            case ':': return Make2(TokenKind.Colon, ":", line, col);
            case '_': return Make2(TokenKind.Underscore, "_", line, col);

            case '=': return Make2(TokenKind.Equals, "=", line, col);

            case '<':
                if (_pos < _src.Length && _src[_pos] == '>')
                { Advance(); return Make2(TokenKind.NotEqual, "<>", line, col); }
                if (_pos < _src.Length && _src[_pos] == '=')
                { Advance(); return Make2(TokenKind.LessEqual, "<=", line, col); }
                return Make2(TokenKind.LessThan, "<", line, col);

            case '>':
                if (_pos < _src.Length && _src[_pos] == '=')
                { Advance(); return Make2(TokenKind.GreaterEqual, ">=", line, col); }
                return Make2(TokenKind.GreaterThan, ">", line, col);

            default:
                // Unknown character — treat as identifier text so the parser can report a useful error
                return Make2(TokenKind.Identifier, c.ToString(), line, col);
        }
    }

    // ── Helpers ───────────────────────────────────────────────────────────

    private char Advance()
    {
        char c = _src[_pos++];
        _col++;
        return c;
    }

    private Token Make(TokenKind kind, string text) =>
        new(kind, text, _line, _col);

    private static Token Make2(TokenKind kind, string text, int line, int col) =>
        new(kind, text, line, col);

    private char PeekAt(int offset)
    {
        int i = _pos + offset - 1; // _pos is already past first char when called from some paths
        return i < _src.Length ? _src[i] : '\0';
    }

    private bool MatchKeywordAhead(string kw)
    {
        if (_pos + kw.Length > _src.Length) return false;
        return string.Compare(_src, _pos, kw, 0, kw.Length, StringComparison.OrdinalIgnoreCase) == 0;
    }

    private static bool IsHexChar(char c) =>
        char.IsDigit(c) || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f');
}
