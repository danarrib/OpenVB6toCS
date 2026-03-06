using VB6toCS.Core.Lexing;
using VB6toCS.Core.Parsing.Nodes;

namespace VB6toCS.Core.Parsing;

/// <summary>
/// Recursive descent parser for VB6 ActiveX DLL source files.
/// Consumes a token list from the Lexer and produces a <see cref="ModuleNode"/> root.
/// </summary>
public sealed class Parser
{
    private readonly TokenStream _ts;
    private readonly string _filePath;

    public Parser(IReadOnlyList<Token> tokens, string filePath)
    {
        _ts = new TokenStream(tokens);
        _filePath = filePath;
    }

    // ── Entry point ────────────────────────────────────────────────────────

    public ModuleNode Parse()
    {
        _ts.SkipNewlines();
        ModuleKind kind = DetectModuleKind();
        SkipClassHeader();
        string name = ParseModuleName();
        var (implements, members) = ParseModuleMembers();
        return new ModuleNode(1, 1, kind, name, implements, members);
    }

    // ── Module kind detection ───────────────────────────────────────────────

    private ModuleKind DetectModuleKind()
    {
        // .cls files start with VERSION; .bas files start with Attribute or a declaration
        // We can also detect by checking if KwVersion appears near the top
        for (int i = 0; i < 5; i++)
        {
            var t = _ts.PeekAt(i);
            if (t.Kind == TokenKind.KwVersion) return ModuleKind.Class;
            if (t.Kind == TokenKind.EndOfFile) break;
        }
        return ModuleKind.StaticModule;
    }

    // ── Class header skipping ───────────────────────────────────────────────

    /// <summary>
    /// Skips the VB6 .cls file boilerplate:
    ///   VERSION 1.0 CLASS
    ///   BEGIN ... END
    ///   Attribute VB_Name = "Foo"
    ///   Attribute ...
    ///
    /// Stops when it sees Option, a declaration keyword, a procedure keyword,
    /// or EOF.
    /// </summary>
    private void SkipClassHeader()
    {
        while (true)
        {
            _ts.SkipNewlines();
            var t = _ts.Peek();

            if (t.Kind == TokenKind.EndOfFile) break;

            // Stop when we hit real module content
            if (IsModuleMemberStart(t.Kind)) break;

            // VERSION line
            if (t.Kind == TokenKind.KwVersion) { _ts.SkipToEndOfLine(); continue; }

            // BEGIN ... END block
            if (t.Kind == TokenKind.KwBegin)
            {
                _ts.Consume();
                SkipBeginBlock();
                continue;
            }

            // Attribute line (including VB_Name)
            if (t.Kind == TokenKind.KwAttribute) { _ts.SkipToEndOfLine(); continue; }

            // Option Explicit line
            if (t.Kind == TokenKind.KwOption) { _ts.SkipToEndOfLine(); continue; }

            // Comment — skip
            if (t.Kind == TokenKind.Comment) { _ts.Consume(); continue; }

            break;
        }
    }

    private void SkipBeginBlock()
    {
        // Consume tokens until matching END (depth-tracked for nested BEGIN)
        int depth = 1;
        while (depth > 0 && _ts.Peek().Kind != TokenKind.EndOfFile)
        {
            var t = _ts.Consume();
            if (t.Kind == TokenKind.KwBegin) depth++;
            else if (t.Kind == TokenKind.KwEnd) depth--;
        }
        _ts.Match(TokenKind.Newline);
    }

    // ── Module name extraction ──────────────────────────────────────────────

    /// <summary>
    /// Extracts the module name from the first  Attribute VB_Name = "Foo"  line.
    /// Falls back to the file name if not found.
    /// </summary>
    private string ParseModuleName()
    {
        // VB_Name attribute was already consumed in SkipClassHeader for .cls files.
        // For .bas files it may appear before any declarations.
        // We scan forward in the remaining tokens without consuming.
        for (int i = 0; i < _ts.PeekAt(0).Line + 20; i++)
        {
            var t = _ts.PeekAt(i);
            if (t.Kind == TokenKind.EndOfFile) break;
            if (t.Kind == TokenKind.KwAttribute)
            {
                // Check for:  Attribute VB_Name = "Foo"
                var id = _ts.PeekAt(i + 1);
                if (id.Kind == TokenKind.Identifier &&
                    id.Text.Equals("VB_Name", StringComparison.OrdinalIgnoreCase))
                {
                    // find the string literal on this line
                    for (int j = i + 2; j < i + 6; j++)
                    {
                        var v = _ts.PeekAt(j);
                        if (v.Kind == TokenKind.StringLiteral)
                            return UnquoteString(v.Text);
                        if (v.Kind == TokenKind.Newline || v.Kind == TokenKind.EndOfFile)
                            break;
                    }
                }
            }
        }
        // Fallback: derive from file path
        return Path.GetFileNameWithoutExtension(_filePath);
    }

    private static string UnquoteString(string raw)
    {
        if (raw.Length >= 2 && raw[0] == '"' && raw[^1] == '"')
            return raw[1..^1].Replace("\"\"", "\"");
        return raw;
    }

    // ── Module members ─────────────────────────────────────────────────────

    private (IReadOnlyList<string> Implements, IReadOnlyList<AstNode> Members) ParseModuleMembers()
    {
        var implements = new List<string>();
        var members = new List<AstNode>();

        while (true)
        {
            _ts.SkipNewlines();
            if (_ts.Peek().Kind == TokenKind.EndOfFile) break;

            // Comments at module level
            if (_ts.Peek().Kind == TokenKind.Comment)
            {
                var c = _ts.Consume();
                members.Add(new CommentNode(c.Line, c.Column, c.Text));
                continue;
            }

            // Option Explicit (already validated, skip)
            if (_ts.Peek().Kind == TokenKind.KwOption)
            {
                _ts.SkipToEndOfLine();
                continue;
            }

            // Attribute lines at module level
            if (_ts.Peek().Kind == TokenKind.KwAttribute)
            {
                _ts.SkipToEndOfLine();
                continue;
            }

            if (!IsModuleMemberStart(_ts.Peek().Kind)) break;

            AstNode member = ParseModuleMember();

            if (member is ImplementsNode impl)
                implements.Add(impl.InterfaceName);
            else
                members.Add(member);
        }

        return (implements, members);
    }

    private bool IsModuleMemberStart(TokenKind k) => k switch
    {
        TokenKind.KwPublic or TokenKind.KwPrivate or TokenKind.KwFriend or TokenKind.KwGlobal => true,
        TokenKind.KwDim or TokenKind.KwConst or TokenKind.KwStatic => true,
        TokenKind.KwEnum or TokenKind.KwType or TokenKind.KwImplements => true,
        TokenKind.KwSub or TokenKind.KwFunction or TokenKind.KwProperty => true,
        TokenKind.KwOption or TokenKind.KwAttribute => true,
        _ => false
    };

    private AstNode ParseModuleMember()
    {
        var start = _ts.Peek();
        AccessModifier access = ParseAccessModifier();
        bool isStatic = _ts.Match(TokenKind.KwStatic);

        var next = _ts.Peek();

        // Declare Sub/Function — external DLL declaration (e.g. Win32 API)
        // "Declare" is not a keyword token; it comes through as Identifier("Declare")
        if (next.Kind == TokenKind.Identifier &&
            next.Text.Equals("Declare", StringComparison.OrdinalIgnoreCase))
            return ParseDeclare(start, access);

        return next.Kind switch
        {
            TokenKind.KwDim => ParseFieldDeclaration(start, access),
            TokenKind.KwConst => ParseConstDeclaration(start, access),
            TokenKind.KwEnum => ParseEnum(start, access),
            TokenKind.KwType => ParseUdt(start, access),
            TokenKind.KwImplements => ParseImplements(),
            TokenKind.KwSub => ParseSub(start, access, isStatic),
            TokenKind.KwFunction => ParseFunction(start, access, isStatic),
            TokenKind.KwProperty => ParseProperty(start, access, isStatic),
            // A bare identifier or type keyword at module level → treat as field (Dim-less)
            _ when next.Kind == TokenKind.Identifier || IsTypeName(next.Kind) =>
                ParseFieldDeclaration(start, access),
            _ => throw new ParseException(
                $"Unexpected token '{next.Text}' at module level", next.Line, next.Column)
        };
    }

    private AccessModifier ParseAccessModifier()
    {
        return _ts.Peek().Kind switch
        {
            TokenKind.KwPublic => Consume(AccessModifier.Public),
            TokenKind.KwPrivate => Consume(AccessModifier.Private),
            TokenKind.KwFriend => Consume(AccessModifier.Friend),
            TokenKind.KwGlobal => Consume(AccessModifier.Global),
            _ => AccessModifier.Public  // default in .bas modules
        };

        AccessModifier Consume(AccessModifier m) { _ts.Consume(); return m; }
    }

    // ── Field / Const declarations ─────────────────────────────────────────

    private FieldNode ParseFieldDeclaration(Token start, AccessModifier access)
    {
        _ts.Match(TokenKind.KwDim); // optional, might not be present
        var declarators = ParseVariableDeclaratorList();
        _ts.ExpectEndOfStatement();
        return new FieldNode(start.Line, start.Column, access, declarators);
    }

    private ConstDeclarationNode ParseConstDeclaration(Token start, AccessModifier access)
    {
        _ts.Expect(TokenKind.KwConst);
        var declarators = ParseConstDeclaratorList();
        _ts.ExpectEndOfStatement();
        return new ConstDeclarationNode(start.Line, start.Column, access, declarators);
    }

    private List<VariableDeclaratorNode> ParseVariableDeclaratorList()
    {
        var list = new List<VariableDeclaratorNode>();
        list.Add(ParseVariableDeclarator());
        while (_ts.Match(TokenKind.Comma))
            list.Add(ParseVariableDeclarator());
        return list;
    }

    private List<VariableDeclaratorNode> ParseConstDeclaratorList()
    {
        var list = new List<VariableDeclaratorNode>();
        list.Add(ParseConstDeclarator());
        while (_ts.Match(TokenKind.Comma))
            list.Add(ParseConstDeclarator());
        return list;
    }

    private VariableDeclaratorNode ParseVariableDeclarator()
    {
        var nameTok = ExpectIdentifierOrKeyword();

        // Array syntax on the variable name: Dim arr() As String  or  Dim arr(10) As Integer
        bool nameIsArray = false;
        if (_ts.Match(TokenKind.LParen))
        {
            nameIsArray = true;
            // Consume optional dimension list
            int depth = 1;
            while (depth > 0 && _ts.Peek().Kind != TokenKind.EndOfFile)
            {
                var t = _ts.Consume();
                if (t.Kind == TokenKind.LParen) depth++;
                else if (t.Kind == TokenKind.RParen) depth--;
            }
        }

        TypeRefNode? typeRef = null;
        if (_ts.Match(TokenKind.KwAs))
        {
            typeRef = ParseTypeRef();
            if (nameIsArray && !typeRef.IsArray)
                typeRef = typeRef with { IsArray = true, ArrayRank = 1 };
        }
        else if (nameIsArray)
        {
            // Array with no explicit type — treat as Variant array
            typeRef = new TypeRefNode("Variant", true, false, 1);
        }

        return new VariableDeclaratorNode(nameTok.Line, nameTok.Column, nameTok.Text, typeRef, null);
    }

    private VariableDeclaratorNode ParseConstDeclarator()
    {
        var nameTok = ExpectIdentifierOrKeyword();
        TypeRefNode? typeRef = null;
        if (_ts.Match(TokenKind.KwAs))
            typeRef = ParseTypeRef();
        _ts.Expect(TokenKind.Equals);
        var value = ParseExpression();
        return new VariableDeclaratorNode(nameTok.Line, nameTok.Column, nameTok.Text, typeRef, value);
    }

    // ── Enum ───────────────────────────────────────────────────────────────

    private EnumNode ParseEnum(Token start, AccessModifier access)
    {
        _ts.Expect(TokenKind.KwEnum);
        var nameTok = _ts.Expect(TokenKind.Identifier);
        _ts.ExpectEndOfStatement();

        var members = new List<EnumMemberNode>();
        while (true)
        {
            _ts.SkipNewlines();
            if (_ts.Check(TokenKind.KwEnd)) break;
            if (_ts.Peek().Kind == TokenKind.EndOfFile) break;

            if (_ts.Peek().Kind == TokenKind.Comment)
            {
                _ts.Consume(); // skip comments inside enum for now
                continue;
            }

            var mName = ExpectIdentifierOrKeyword();
            ExpressionNode? value = null;
            if (_ts.Match(TokenKind.Equals))
                value = ParseExpression();

            // collect trailing comments on same line
            var trailing = new List<CommentNode>();
            while (_ts.Peek().Kind == TokenKind.Comment)
            {
                var c = _ts.Consume();
                trailing.Add(new CommentNode(c.Line, c.Column, c.Text));
            }
            _ts.ExpectEndOfStatement();
            members.Add(new EnumMemberNode(mName.Line, mName.Column, mName.Text, value, trailing));
        }

        _ts.Expect(TokenKind.KwEnd);
        _ts.Expect(TokenKind.KwEnum);
        _ts.ExpectEndOfStatement();

        return new EnumNode(start.Line, start.Column, access, nameTok.Text, members);
    }

    // ── UDT (Type) ─────────────────────────────────────────────────────────

    private UdtNode ParseUdt(Token start, AccessModifier access)
    {
        _ts.Expect(TokenKind.KwType);
        var nameTok = _ts.Expect(TokenKind.Identifier);
        _ts.ExpectEndOfStatement();

        var fields = new List<UdtFieldNode>();
        while (true)
        {
            _ts.SkipNewlines();
            if (_ts.Check(TokenKind.KwEnd)) break;
            if (_ts.Peek().Kind == TokenKind.EndOfFile) break;
            if (_ts.Peek().Kind == TokenKind.Comment) { _ts.Consume(); continue; }

            var fName = ExpectIdentifierOrKeyword();
            _ts.Expect(TokenKind.KwAs);
            var fType = ParseTypeRef();
            _ts.ExpectEndOfStatement();
            fields.Add(new UdtFieldNode(fName.Line, fName.Column, fName.Text, fType));
        }

        _ts.Expect(TokenKind.KwEnd);
        _ts.Expect(TokenKind.KwType);
        _ts.ExpectEndOfStatement();

        return new UdtNode(start.Line, start.Column, access, nameTok.Text, fields);
    }

    // ── Implements ─────────────────────────────────────────────────────────

    private ImplementsNode ParseImplements()
    {
        var tok = _ts.Expect(TokenKind.KwImplements);
        var name = ExpectIdentifierOrKeyword();
        _ts.ExpectEndOfStatement();
        return new ImplementsNode(tok.Line, tok.Column, name.Text);
    }

    // ── Declare ────────────────────────────────────────────────────────────

    private DeclareNode ParseDeclare(Token start, AccessModifier access)
    {
        _ts.Consume(); // "Declare" identifier

        bool isSub = _ts.Peek().Kind == TokenKind.KwSub;
        if (isSub) _ts.Consume();
        else _ts.Expect(TokenKind.KwFunction);

        var nameTok = ExpectIdentifierOrKeyword();

        // Lib "libname"
        if (!(_ts.Peek().Kind == TokenKind.Identifier &&
              _ts.Peek().Text.Equals("Lib", StringComparison.OrdinalIgnoreCase)))
            throw new ParseException("Expected 'Lib' after Declare name", _ts.Peek().Line, _ts.Peek().Column);
        _ts.Consume();
        var libTok = _ts.Expect(TokenKind.StringLiteral);
        string libName = UnquoteString(libTok.Text);

        // Optional: Alias "aliasname"
        string? aliasName = null;
        if (_ts.Peek().Kind == TokenKind.Identifier &&
            _ts.Peek().Text.Equals("Alias", StringComparison.OrdinalIgnoreCase))
        {
            _ts.Consume();
            var aliasTok = _ts.Expect(TokenKind.StringLiteral);
            aliasName = UnquoteString(aliasTok.Text);
        }

        var parameters = ParseParameterList();

        TypeRefNode? returnType = null;
        if (!isSub && _ts.Match(TokenKind.KwAs))
            returnType = ParseTypeRef();

        _ts.ExpectEndOfStatement();
        return new DeclareNode(start.Line, start.Column, access, isSub,
            nameTok.Text, libName, aliasName, parameters, returnType);
    }

    // ── Sub / Function / Property ──────────────────────────────────────────

    // True for any keyword that can follow "End" to close a procedure body.
    // Used as a broad terminator so mismatched End Sub/Function/Property are tolerated.
    private static bool IsProcedureEndKeyword(TokenKind k) =>
        k is TokenKind.KwSub or TokenKind.KwFunction or TokenKind.KwProperty;

    private SubNode ParseSub(Token start, AccessModifier access, bool isStatic)
    {
        _ts.Expect(TokenKind.KwSub);
        var nameTok = ExpectIdentifierOrKeyword();
        var parameters = ParseParameterList();
        _ts.ExpectEndOfStatement();

        var body = ParseBody(isEndOf: s => s.Check(TokenKind.KwEnd) && IsProcedureEndKeyword(s.PeekAt(1).Kind));

        _ts.Expect(TokenKind.KwEnd);
        if (IsProcedureEndKeyword(_ts.Peek().Kind)) _ts.Consume(); // accept any; mismatch tolerated
        _ts.ExpectEndOfStatement();

        return new SubNode(start.Line, start.Column, access, isStatic, nameTok.Text, parameters, body);
    }

    private FunctionNode ParseFunction(Token start, AccessModifier access, bool isStatic)
    {
        _ts.Expect(TokenKind.KwFunction);
        var nameTok = ExpectIdentifierOrKeyword();
        var parameters = ParseParameterList();
        TypeRefNode? returnType = null;
        if (_ts.Match(TokenKind.KwAs))
            returnType = ParseTypeRef();
        _ts.ExpectEndOfStatement();

        var body = ParseBody(isEndOf: s => s.Check(TokenKind.KwEnd) && IsProcedureEndKeyword(s.PeekAt(1).Kind));

        _ts.Expect(TokenKind.KwEnd);
        if (IsProcedureEndKeyword(_ts.Peek().Kind)) _ts.Consume(); // accept any; mismatch tolerated
        _ts.ExpectEndOfStatement();

        return new FunctionNode(start.Line, start.Column, access, isStatic, nameTok.Text, parameters, returnType, body);
    }

    private PropertyNode ParseProperty(Token start, AccessModifier access, bool isStatic)
    {
        _ts.Expect(TokenKind.KwProperty);
        var kindTok = _ts.Peek();
        PropertyKind propKind = kindTok.Kind switch
        {
            TokenKind.KwGet => PropertyKind.Get,
            TokenKind.KwSet => PropertyKind.Set,
            TokenKind.KwLet => PropertyKind.Let,
            _ => throw new ParseException(
                $"Expected Get, Set, or Let after Property, found '{kindTok.Text}'",
                kindTok.Line, kindTok.Column)
        };
        _ts.Consume();

        var nameTok = ExpectIdentifierOrKeyword();
        var parameters = ParseParameterList();
        TypeRefNode? returnType = null;
        if (_ts.Match(TokenKind.KwAs))
            returnType = ParseTypeRef();
        _ts.ExpectEndOfStatement();

        var body = ParseBody(isEndOf: s => s.Check(TokenKind.KwEnd) && IsProcedureEndKeyword(s.PeekAt(1).Kind));

        _ts.Expect(TokenKind.KwEnd);
        if (IsProcedureEndKeyword(_ts.Peek().Kind)) _ts.Consume(); // accept any; mismatch tolerated
        _ts.ExpectEndOfStatement();

        return new PropertyNode(start.Line, start.Column, access, isStatic, propKind, nameTok.Text, parameters, returnType, body);
    }

    // ── Parameter list ─────────────────────────────────────────────────────

    private IReadOnlyList<ParameterNode> ParseParameterList()
    {
        var list = new List<ParameterNode>();
        if (!_ts.Match(TokenKind.LParen)) return list;
        if (_ts.Match(TokenKind.RParen)) return list;

        list.Add(ParseParameter());
        while (_ts.Match(TokenKind.Comma))
            list.Add(ParseParameter());

        _ts.Expect(TokenKind.RParen);
        return list;
    }

    private ParameterNode ParseParameter()
    {
        var start = _ts.Peek();
        bool isOptional = _ts.Match(TokenKind.KwOptional);
        bool isParamArray = _ts.Match(TokenKind.KwParamArray);

        ParameterMode mode = ParameterMode.Unspecified;
        if (_ts.Match(TokenKind.KwByVal)) mode = ParameterMode.ByVal;
        else if (_ts.Match(TokenKind.KwByRef)) mode = ParameterMode.ByRef;

        var nameTok = ExpectIdentifierOrKeyword();

        // ParamArray values() — consume the () that declares the array nature of the parameter
        bool nameHasArrayParens = false;
        if (_ts.Match(TokenKind.LParen))
        {
            nameHasArrayParens = true;
            _ts.Expect(TokenKind.RParen);
        }

        TypeRefNode? typeRef = null;
        if (_ts.Match(TokenKind.KwAs))
        {
            typeRef = ParseTypeRef();
            // If the name had () and the type doesn't, mark it as array
            if (nameHasArrayParens && !typeRef.IsArray)
                typeRef = typeRef with { IsArray = true, ArrayRank = 1 };
        }

        ExpressionNode? defaultValue = null;
        if (_ts.Match(TokenKind.Equals))
            defaultValue = ParseExpression();

        return new ParameterNode(start.Line, start.Column, mode, isOptional, isParamArray,
            nameTok.Text, typeRef, defaultValue);
    }

    // ── TypeRef ────────────────────────────────────────────────────────────

    private TypeRefNode ParseTypeRef()
    {
        bool isNew = _ts.Match(TokenKind.KwNew);

        // Collect type name (may be keyword like Integer, String, Object, or identifier)
        var nameTok = _ts.Peek();
        if (!IsTypeName(nameTok.Kind) && nameTok.Kind != TokenKind.Identifier)
            throw new ParseException($"Expected type name, found '{nameTok.Text}'", nameTok.Line, nameTok.Column);
        _ts.Consume();

        // Handle dotted type names: ADODB.Connection, Scripting.Dictionary, etc.
        string typeName = nameTok.Text;
        while (_ts.Peek().Kind == TokenKind.Dot)
        {
            _ts.Consume(); // dot
            var part = ExpectIdentifierOrKeyword();
            typeName += "." + part.Text;
        }

        // Fixed-length string: String * n  (e.g. As String * 80)
        // Translate as plain String — the fixed-length constraint is dropped.
        if (typeName.Equals("String", StringComparison.OrdinalIgnoreCase) &&
            _ts.Check(TokenKind.Star))
        {
            _ts.Consume(); // *
            _ts.Consume(); // the length literal
        }
        bool isArray = false;
        int arrayRank = 0;

        // Optional array dimensions: (n) or (n, m) or ()
        if (_ts.Match(TokenKind.LParen))
        {
            isArray = true;
            arrayRank = 1;
            int depth = 1;
            while (depth > 0 && _ts.Peek().Kind != TokenKind.EndOfFile)
            {
                var t = _ts.Consume();
                if (t.Kind == TokenKind.Comma) arrayRank++;
                else if (t.Kind == TokenKind.RParen) depth--;
            }
        }

        return new TypeRefNode(typeName, isArray, isNew, arrayRank);
    }

    private bool IsTypeName(TokenKind k) => k switch
    {
        TokenKind.KwBoolean or TokenKind.KwByte or TokenKind.KwInteger or TokenKind.KwLong
        or TokenKind.KwSingle or TokenKind.KwDouble or TokenKind.KwCurrency or TokenKind.KwDecimal
        or TokenKind.KwDate or TokenKind.KwString or TokenKind.KwObject or TokenKind.KwVariant => true,
        _ => false
    };

    // ── Body parsing ────────────────────────────────────────────────────────

    private IReadOnlyList<AstNode> ParseBody(Func<TokenStream, bool> isEndOf)
    {
        var stmts = new List<AstNode>();

        while (true)
        {
            _ts.SkipNewlines();
            if (_ts.Peek().Kind == TokenKind.EndOfFile) break;
            if (isEndOf(_ts)) break;

            var stmt = ParseStatement();
            if (stmt != null) stmts.Add(stmt);
        }

        return stmts;
    }

    private AstNode? ParseStatement()
    {
        _ts.SkipNewlines();
        var t = _ts.Peek();

        // Label: any token (identifier OR keyword) immediately followed by ':' on the same line.
        // e.g. "MyLabel:", "SAIDA:", "ERROR:" (keyword used as label name)
        if (t.Kind == TokenKind.Identifier || t.Kind >= TokenKind.KwPublic)
        {
            var next = _ts.PeekAt(1);
            if (next.Kind == TokenKind.Colon && next.Line == t.Line)
            {
                _ts.Consume(); // label name
                _ts.Consume(); // colon
                return new LabelNode(t.Line, t.Column, t.Text);
            }
        }

        return t.Kind switch
        {
            TokenKind.Comment => ParseComment(),
            TokenKind.KwDim => ParseDim(),
            TokenKind.KwStatic => ParseStaticDim(),
            TokenKind.KwConst => ParseLocalConst(),
            TokenKind.KwIf => ParseIf(),
            TokenKind.KwSelect => ParseSelectCase(),
            TokenKind.KwFor => ParseFor(),
            TokenKind.KwWhile => ParseWhile(),
            TokenKind.KwDo => ParseDoLoop(),
            TokenKind.KwWith => ParseWith(),
            TokenKind.KwOn => ParseOnError(),
            TokenKind.KwResume => ParseResume(),
            TokenKind.KwGoTo => ParseGoTo(),
            TokenKind.KwGoSub => ParseGoSub(),
            TokenKind.KwReturn => ParseReturn(),
            TokenKind.KwExit => ParseExit(),
            TokenKind.KwEnd => ParseEndStatement(),
            TokenKind.KwError => ParseErrorStatement(),
            // Numeric line label: "10 Statement" — consume the number, emit a LabelNode.
            // The body loop will parse the following statement on the next iteration.
            TokenKind.IntegerLiteral => ParseNumericLabel(),
            TokenKind.KwSet => ParseSetAssignment(),
            TokenKind.KwLet => ParseLetAssignment(),
            TokenKind.KwCall => ParseCallStatement(),
            TokenKind.KwAttribute => ParseAttributeLine(),
            // Conditional compilation directives: #If, #ElseIf, #Else, #End If
            // Skip the directive line; include code from all branches.
            TokenKind.Hash when IsConditionalCompilationDirective() => SkipConditionalDirectiveLine(t),
            // Identifier or keyword used as call/assignment
            _ when t.Kind == TokenKind.Identifier &&
                   t.Text.Equals("ReDim", StringComparison.OrdinalIgnoreCase) => ParseReDim(),
            _ when IsFileIOStatement(t) => SkipFileIOStatement(t),
            _ when IsStatementStart(t.Kind) => ParseAssignmentOrCall(),
            _ => throw new ParseException($"Unexpected token '{t.Text}' in statement", t.Line, t.Column)
        };
    }

    private bool IsStatementStart(TokenKind k) =>
        k == TokenKind.Identifier || k == TokenKind.KwMe ||
        IsTypeName(k) || k == TokenKind.Dot;

    // ── Simple statements ───────────────────────────────────────────────────

    private CommentNode ParseComment()
    {
        var t = _ts.Consume();
        return new CommentNode(t.Line, t.Column, t.Text);
    }

    private AstNode? ParseAttributeLine()
    {
        _ts.SkipToEndOfLine();
        return null;
    }

    private LocalDimNode ParseDim()
    {
        var start = _ts.Expect(TokenKind.KwDim);
        var declarators = ParseVariableDeclaratorList();
        _ts.ExpectEndOfStatement();
        return new LocalDimNode(start.Line, start.Column, false, declarators);
    }

    private LocalDimNode ParseStaticDim()
    {
        var start = _ts.Expect(TokenKind.KwStatic);
        var declarators = ParseVariableDeclaratorList();
        _ts.ExpectEndOfStatement();
        return new LocalDimNode(start.Line, start.Column, true, declarators);
    }

    private ConstDeclarationNode ParseLocalConst()
    {
        var start = _ts.Peek();
        _ts.Expect(TokenKind.KwConst);
        var declarators = ParseConstDeclaratorList();
        _ts.ExpectEndOfStatement();
        return new ConstDeclarationNode(start.Line, start.Column, AccessModifier.Private, declarators);
    }

    private ReDimNode ParseReDim()
    {
        var start = _ts.Consume(); // "ReDim" identifier
        bool isPreserve = _ts.Peek().Kind == TokenKind.Identifier &&
                          _ts.Peek().Text.Equals("Preserve", StringComparison.OrdinalIgnoreCase);
        if (isPreserve) _ts.Consume();

        var declarators = new List<VariableDeclaratorNode>();
        declarators.Add(ParseVariableDeclarator());
        while (_ts.Match(TokenKind.Comma))
            declarators.Add(ParseVariableDeclarator());

        _ts.ExpectEndOfStatement();
        return new ReDimNode(start.Line, start.Column, isPreserve, declarators);
    }

    // VB6 file I/O statements — out of scope; skip the line and emit a comment placeholder.
    private bool IsFileIOStatement(Token t)
    {
        if (t.Kind != TokenKind.Identifier) return false;
        var text = t.Text;
        if (text.Equals("Open",  StringComparison.OrdinalIgnoreCase)) return true;
        if (text.Equals("Close", StringComparison.OrdinalIgnoreCase)) return true;
        if (text.Equals("Print", StringComparison.OrdinalIgnoreCase) &&
            _ts.PeekAt(1).Kind == TokenKind.Hash) return true;
        if (text.Equals("Line",  StringComparison.OrdinalIgnoreCase) &&
            _ts.PeekAt(1).Kind == TokenKind.Identifier &&
            _ts.PeekAt(1).Text.Equals("Input", StringComparison.OrdinalIgnoreCase)) return true;
        if (text.Equals("Get",   StringComparison.OrdinalIgnoreCase) &&
            _ts.PeekAt(1).Kind == TokenKind.Hash) return true;
        if (text.Equals("Put",   StringComparison.OrdinalIgnoreCase) &&
            _ts.PeekAt(1).Kind == TokenKind.Hash) return true;
        return false;
    }

    private CommentNode SkipFileIOStatement(Token t)
    {
        _ts.SkipToEndOfLine();
        return new CommentNode(t.Line, t.Column, $"' TODO: File I/O not supported — skipped: {t.Text}");
    }

    // Returns true when the current Hash token is the start of a #If/#ElseIf/#Else/#End directive.
    private bool IsConditionalCompilationDirective()
    {
        // Hash must be followed by If, ElseIf, Else, or End (possibly with If after)
        var next = _ts.PeekAt(1);
        return next.Kind is TokenKind.KwIf or TokenKind.KwElseIf or TokenKind.KwElse or TokenKind.KwEnd;
    }

    private CommentNode? SkipConditionalDirectiveLine(Token t)
    {
        _ts.SkipToEndOfLine();
        return null; // directive line is noise; code inside the block is parsed normally
    }

    private AstNode ParseEndStatement()
    {
        var start = _ts.Peek(); // KwEnd
        var next = _ts.PeekAt(1);

        if (next.Kind is TokenKind.Newline or TokenKind.EndOfFile or TokenKind.Colon)
        {
            _ts.Consume();
            _ts.ExpectEndOfStatement();
            return new EndStatementNode(start.Line, start.Column);
        }

        // Mismatched procedure terminator (e.g. "End Function" inside a Property Get).
        // VB6 is lenient about these. Consume both tokens so parsing can continue.
        if (next.Kind is TokenKind.KwSub or TokenKind.KwFunction or TokenKind.KwProperty)
        {
            _ts.Consume(); // End
            _ts.Consume(); // Sub/Function/Property
            _ts.ExpectEndOfStatement();
            return new EndStatementNode(start.Line, start.Column);
        }

        // Bare End or other block terminator — let parent handle it via isEndOf.
        _ts.Consume();
        return new EndStatementNode(start.Line, start.Column);
    }

    private ExitNode ParseExit()
    {
        var start = _ts.Expect(TokenKind.KwExit);
        var what = _ts.Peek();
        string whatStr = what.Kind switch
        {
            TokenKind.KwSub => "Sub",
            TokenKind.KwFunction => "Function",
            TokenKind.KwFor => "For",
            TokenKind.KwDo => "Do",
            TokenKind.KwWhile => "While",
            TokenKind.KwProperty => "Property",
            _ => throw new ParseException($"Expected Sub/Function/For/Do/While after Exit", what.Line, what.Column)
        };
        _ts.Consume();
        _ts.ExpectEndOfStatement();
        return new ExitNode(start.Line, start.Column, whatStr);
    }

    private ReturnNode ParseReturn()
    {
        var t = _ts.Expect(TokenKind.KwReturn);
        _ts.ExpectEndOfStatement();
        return new ReturnNode(t.Line, t.Column);
    }

    private ErrorStatementNode ParseErrorStatement()
    {
        var start = _ts.Expect(TokenKind.KwError);
        var num = ParseExpression();
        _ts.ExpectEndOfStatement();
        return new ErrorStatementNode(start.Line, start.Column, num);
    }

    private LabelNode ParseNumericLabel()
    {
        var t = _ts.Consume(); // integer literal used as a line label
        // Do NOT call ExpectEndOfStatement — the actual statement follows on the same line.
        return new LabelNode(t.Line, t.Column, t.Text);
    }

    private GoToNode ParseGoTo()
    {
        var start = _ts.Expect(TokenKind.KwGoTo);
        var label = ExpectIdentifierOrKeyword();
        _ts.ExpectEndOfStatement();
        return new GoToNode(start.Line, start.Column, label.Text);
    }

    private GoSubNode ParseGoSub()
    {
        var start = _ts.Expect(TokenKind.KwGoSub);
        var label = ExpectIdentifierOrKeyword();
        _ts.ExpectEndOfStatement();
        return new GoSubNode(start.Line, start.Column, label.Text);
    }

    // ── On Error / Resume ───────────────────────────────────────────────────

    private OnErrorNode ParseOnError()
    {
        var start = _ts.Expect(TokenKind.KwOn);

        // "On Local Error" is a VB6 variant meaning the same as "On Error" but
        // scoped to the current procedure. "Local" is not a keyword — consume it as identifier.
        if (_ts.Peek().Kind == TokenKind.Identifier &&
            _ts.Peek().Text.Equals("Local", StringComparison.OrdinalIgnoreCase))
            _ts.Consume();

        _ts.Expect(TokenKind.KwError);

        if (_ts.Match(TokenKind.KwResume))
        {
            _ts.Expect(TokenKind.KwNext);
            _ts.ExpectEndOfStatement();
            return new OnErrorNode(start.Line, start.Column, OnErrorKind.ResumeNext, null);
        }

        if (_ts.Match(TokenKind.KwGoTo))
        {
            var target = _ts.Peek();
            if (target.Kind == TokenKind.IntegerLiteral && target.Text == "0")
            {
                _ts.Consume();
                _ts.ExpectEndOfStatement();
                return new OnErrorNode(start.Line, start.Column, OnErrorKind.GoToZero, null);
            }
            var label = ExpectIdentifierOrKeyword();
            _ts.ExpectEndOfStatement();
            return new OnErrorNode(start.Line, start.Column, OnErrorKind.GoTo, label.Text);
        }

        throw new ParseException("Expected GoTo or Resume Next after On Error", start.Line, start.Column);
    }

    private ResumeNode ParseResume()
    {
        var start = _ts.Expect(TokenKind.KwResume);
        if (_ts.Match(TokenKind.KwNext))
        {
            _ts.ExpectEndOfStatement();
            return new ResumeNode(start.Line, start.Column, true, null);
        }
        if (_ts.AtEndOfStatement())
        {
            _ts.ExpectEndOfStatement();
            return new ResumeNode(start.Line, start.Column, false, null);
        }
        var label = ExpectIdentifierOrKeyword();
        _ts.ExpectEndOfStatement();
        return new ResumeNode(start.Line, start.Column, false, label.Text);
    }

    // ── If statement ───────────────────────────────────────────────────────

    private AstNode ParseIf()
    {
        var start = _ts.Expect(TokenKind.KwIf);
        var condition = ParseExpression();
        _ts.Expect(TokenKind.KwThen);

        // Single-line If: If cond Then stmt [Else stmt]
        if (!_ts.AtEndOfStatement())
        {
            _ts.SingleLineIfMode = true;
            try
            {
                var thenStmt = ParseStatement()!;
                AstNode? elseStmt = null;
                if (_ts.Match(TokenKind.KwElse))
                    elseStmt = ParseStatement();
                // Each inner statement already consumed EOS; nothing left to consume here.
                return new SingleLineIfNode(start.Line, start.Column, condition, thenStmt, elseStmt);
            }
            finally
            {
                _ts.SingleLineIfMode = false;
            }
        }

        _ts.ExpectEndOfStatement();

        // Block If
        var thenBody = ParseBody(isEndOf: s =>
            s.Check(TokenKind.KwEnd) ||
            s.Check(TokenKind.KwElse) ||
            s.Check(TokenKind.KwElseIf));

        var elseIfClauses = new List<ElseIfClauseNode>();
        while (_ts.Check(TokenKind.KwElseIf) ||
               (_ts.Check(TokenKind.KwElse) && _ts.PeekAt(1).Kind == TokenKind.KwIf))
        {
            var eiStart = _ts.Peek();
            if (_ts.Check(TokenKind.KwElseIf))
                _ts.Consume();
            else
            {
                _ts.Consume(); // KwElse
                _ts.Consume(); // KwIf
            }
            var eiCond = ParseExpression();
            _ts.Expect(TokenKind.KwThen);
            _ts.ExpectEndOfStatement();
            var eiBody = ParseBody(isEndOf: s =>
                s.Check(TokenKind.KwEnd) ||
                s.Check(TokenKind.KwElse) ||
                s.Check(TokenKind.KwElseIf));
            elseIfClauses.Add(new ElseIfClauseNode(eiStart.Line, eiStart.Column, eiCond, eiBody));
        }

        IReadOnlyList<AstNode>? elseBody = null;
        if (_ts.Match(TokenKind.KwElse))
        {
            _ts.ExpectEndOfStatement();
            elseBody = ParseBody(isEndOf: s => s.Check(TokenKind.KwEnd));
        }

        _ts.Expect(TokenKind.KwEnd);
        _ts.Expect(TokenKind.KwIf);
        _ts.ExpectEndOfStatement();

        return new IfNode(start.Line, start.Column, condition, thenBody, elseIfClauses, elseBody);
    }

    // ── Select Case ────────────────────────────────────────────────────────

    private SelectCaseNode ParseSelectCase()
    {
        var start = _ts.Expect(TokenKind.KwSelect);
        _ts.Expect(TokenKind.KwCase);
        var testExpr = ParseExpression();
        _ts.ExpectEndOfStatement();

        var cases = new List<CaseClauseNode>();
        while (true)
        {
            _ts.SkipNewlines();
            if (_ts.Check(TokenKind.KwEnd)) break;
            if (_ts.Peek().Kind == TokenKind.EndOfFile) break;
            if (_ts.Peek().Kind == TokenKind.Comment) { _ts.Consume(); continue; }
            if (!_ts.Check(TokenKind.KwCase)) break;

            cases.Add(ParseCaseClause());
        }

        _ts.Expect(TokenKind.KwEnd);
        _ts.Expect(TokenKind.KwSelect);
        _ts.ExpectEndOfStatement();

        return new SelectCaseNode(start.Line, start.Column, testExpr, cases);
    }

    private CaseClauseNode ParseCaseClause()
    {
        var start = _ts.Expect(TokenKind.KwCase);
        bool isElse = false;
        var patterns = new List<CasePatternNode>();

        if (_ts.Match(TokenKind.KwElse))
        {
            isElse = true;
        }
        else
        {
            patterns.Add(ParseCasePattern());
            while (_ts.Match(TokenKind.Comma))
                patterns.Add(ParseCasePattern());
        }
        _ts.ExpectEndOfStatement();

        var body = ParseBody(isEndOf: s =>
            s.Check(TokenKind.KwEnd) || s.Check(TokenKind.KwCase));

        return new CaseClauseNode(start.Line, start.Column, patterns, isElse, body);
    }

    private CasePatternNode ParseCasePattern()
    {
        var t = _ts.Peek();

        // Case Is <op> <expr>
        if (_ts.Match(TokenKind.KwIs))
        {
            var op = _ts.Peek().Kind;
            if (op is not (TokenKind.Equals or TokenKind.NotEqual or TokenKind.LessThan
                or TokenKind.GreaterThan or TokenKind.LessEqual or TokenKind.GreaterEqual))
                throw new ParseException("Expected comparison operator after Case Is", t.Line, t.Column);
            _ts.Consume();
            var val = ParseExpression();
            return new CaseIsPattern(t.Line, t.Column, op, val);
        }

        // Case <expr> [To <expr>]
        var low = ParseExpression();
        if (_ts.Match(TokenKind.KwTo))
        {
            var high = ParseExpression();
            return new CaseRangePattern(t.Line, t.Column, low, high);
        }
        return new CaseValuePattern(t.Line, t.Column, low);
    }

    // ── Loops ──────────────────────────────────────────────────────────────

    private AstNode ParseFor()
    {
        var start = _ts.Expect(TokenKind.KwFor);

        // For Each
        if (_ts.Match(TokenKind.KwEach))
        {
            var varTok = ExpectIdentifierOrKeyword();
            _ts.Expect(TokenKind.KwIn);
            var collection = ParseExpression();
            _ts.ExpectEndOfStatement();

            var body = ParseBody(isEndOf: s => s.Check(TokenKind.KwNext));
            _ts.Expect(TokenKind.KwNext);
            if (!_ts.AtEndOfStatement()) _ts.Consume(); // optional variable name
            _ts.ExpectEndOfStatement();

            return new ForEachNode(start.Line, start.Column, varTok.Text, collection, body);
        }

        // For counter = start To end [Step step]
        var counterTok = ExpectIdentifierOrKeyword();
        _ts.Expect(TokenKind.Equals);
        var fromExpr = ParseExpression();
        _ts.Expect(TokenKind.KwTo);
        var toExpr = ParseExpression();
        ExpressionNode? stepExpr = null;
        if (_ts.Match(TokenKind.KwStep))
            stepExpr = ParseExpression();
        _ts.ExpectEndOfStatement();

        var forBody = ParseBody(isEndOf: s => s.Check(TokenKind.KwNext));
        _ts.Expect(TokenKind.KwNext);
        if (!_ts.AtEndOfStatement()) _ts.Consume(); // optional counter name
        _ts.ExpectEndOfStatement();

        return new ForNextNode(start.Line, start.Column, counterTok.Text, fromExpr, toExpr, stepExpr, forBody);
    }

    private WhileNode ParseWhile()
    {
        var start = _ts.Expect(TokenKind.KwWhile);
        var condition = ParseExpression();
        _ts.ExpectEndOfStatement();

        var body = ParseBody(isEndOf: s => s.Check(TokenKind.KwWend));
        _ts.Expect(TokenKind.KwWend);
        _ts.ExpectEndOfStatement();

        return new WhileNode(start.Line, start.Column, condition, body);
    }

    private DoLoopNode ParseDoLoop()
    {
        var start = _ts.Expect(TokenKind.KwDo);

        // Do While / Do Until (condition at top)
        if (_ts.Peek().Kind is TokenKind.KwWhile or TokenKind.KwUntil)
        {
            bool isUntil = _ts.Consume().Kind == TokenKind.KwUntil;
            var condition = ParseExpression();
            _ts.ExpectEndOfStatement();

            var body = ParseBody(isEndOf: s => s.Check(TokenKind.KwLoop));
            _ts.Expect(TokenKind.KwLoop);
            _ts.ExpectEndOfStatement();

            var kind = isUntil ? DoLoopKind.DoUntilTop : DoLoopKind.DoWhileTop;
            return new DoLoopNode(start.Line, start.Column, kind, condition, body);
        }

        // Do ... Loop [While/Until condition] (condition at bottom or forever)
        _ts.ExpectEndOfStatement();
        var loopBody = ParseBody(isEndOf: s => s.Check(TokenKind.KwLoop));
        _ts.Expect(TokenKind.KwLoop);

        if (_ts.Peek().Kind is TokenKind.KwWhile or TokenKind.KwUntil)
        {
            bool isUntil = _ts.Consume().Kind == TokenKind.KwUntil;
            var condition = ParseExpression();
            _ts.ExpectEndOfStatement();
            var kind = isUntil ? DoLoopKind.DoUntilBottom : DoLoopKind.DoWhileBottom;
            return new DoLoopNode(start.Line, start.Column, kind, condition, loopBody);
        }

        _ts.ExpectEndOfStatement();
        return new DoLoopNode(start.Line, start.Column, DoLoopKind.DoForever, null, loopBody);
    }

    // ── With ───────────────────────────────────────────────────────────────

    private WithNode ParseWith()
    {
        var start = _ts.Expect(TokenKind.KwWith);
        var obj = ParseExpression();
        _ts.ExpectEndOfStatement();

        var body = ParseBody(isEndOf: s => s.Check(TokenKind.KwEnd) && s.PeekAt(1).Kind == TokenKind.KwWith);
        _ts.Expect(TokenKind.KwEnd);
        _ts.Expect(TokenKind.KwWith);
        _ts.ExpectEndOfStatement();

        return new WithNode(start.Line, start.Column, obj, body);
    }

    // ── Assignment / Call ──────────────────────────────────────────────────

    private AstNode ParseSetAssignment()
    {
        var start = _ts.Expect(TokenKind.KwSet);
        var target = ParsePostfix();
        _ts.Expect(TokenKind.Equals);
        var value = ParseExpression();
        _ts.ExpectEndOfStatement();
        return new AssignmentNode(start.Line, start.Column, true, false, target, value);
    }

    private AstNode ParseLetAssignment()
    {
        var start = _ts.Expect(TokenKind.KwLet);
        var target = ParsePostfix();
        _ts.Expect(TokenKind.Equals);
        var value = ParseExpression();
        _ts.ExpectEndOfStatement();
        return new AssignmentNode(start.Line, start.Column, false, true, target, value);
    }

    private AstNode ParseCallStatement()
    {
        var start = _ts.Expect(TokenKind.KwCall);
        var target = ParsePostfix();
        List<ArgumentNode> args;

        // Call foo(a, b)  — parenthesized
        if (_ts.Check(TokenKind.LParen))
            args = ParseArgumentListParenthesized();
        else
            args = [];

        _ts.ExpectEndOfStatement();
        return new CallStatementNode(start.Line, start.Column, target, args, true);
    }

    private AstNode ParseAssignmentOrCall()
    {
        var target = ParsePostfix();

        // Assignment: target = expr
        if (_ts.Match(TokenKind.Equals))
        {
            var value = ParseExpression();
            _ts.ExpectEndOfStatement();
            return new AssignmentNode(target.Line, target.Column, false, false, target, value);
        }

        // Call without parentheses: foo a, b
        var args = new List<ArgumentNode>();
        if (!_ts.AtEndOfStatement())
            args = ParseArgumentListUnparenthesized();

        _ts.ExpectEndOfStatement();
        return new CallStatementNode(target.Line, target.Column, target, args, false);
    }

    // ── Argument lists ─────────────────────────────────────────────────────

    private List<ArgumentNode> ParseArgumentListParenthesized()
    {
        _ts.Expect(TokenKind.LParen);
        var args = new List<ArgumentNode>();
        if (_ts.Match(TokenKind.RParen)) return args;

        args.Add(ParseArgument());
        while (_ts.Match(TokenKind.Comma))
        {
            if (_ts.Check(TokenKind.RParen))
            {
                // trailing comma before ) — treated as missing last arg
                args.Add(new ArgumentNode(_ts.Peek().Line, _ts.Peek().Column, null, true, null));
                break;
            }
            args.Add(ParseArgument());
        }

        _ts.Expect(TokenKind.RParen);
        return args;
    }

    private List<ArgumentNode> ParseArgumentListUnparenthesized()
    {
        var args = new List<ArgumentNode>();
        args.Add(ParseArgument());
        while (_ts.Match(TokenKind.Comma))
            args.Add(ParseArgument());
        return args;
    }

    private ArgumentNode ParseArgument()
    {
        var start = _ts.Peek();

        // Missing argument: just a comma follows
        if (start.Kind == TokenKind.Comma)
            return new ArgumentNode(start.Line, start.Column, null, true, null);

        // ByVal/ByRef in call argument list (e.g. CopyMemory ByVal ptr, ByVal src, len).
        // This is a runtime hint — discard the keyword and parse the expression.
        if (start.Kind is TokenKind.KwByVal or TokenKind.KwByRef)
        {
            _ts.Consume();
            start = _ts.Peek();
        }

        // Named argument: name := expr
        if (start.Kind == TokenKind.Identifier && _ts.PeekAt(1).Kind == TokenKind.Colon
            && _ts.PeekAt(2).Kind == TokenKind.Equals)
        {
            var name = _ts.Consume().Text;
            _ts.Consume(); // :
            _ts.Consume(); // =
            var val = ParseExpression();
            return new ArgumentNode(start.Line, start.Column, name, false, val);
        }

        var expr = ParseExpression();
        return new ArgumentNode(start.Line, start.Column, null, false, expr);
    }

    // ── Expressions ────────────────────────────────────────────────────────

    private ExpressionNode ParseExpression() => ParseImp();

    private ExpressionNode ParseImp()
    {
        var left = ParseEqv();
        while (_ts.Check(TokenKind.KwImp))
        {
            var op = _ts.Consume();
            var right = ParseEqv();
            left = new BinaryExpressionNode(op.Line, op.Column, left, op.Kind, right);
        }
        return left;
    }

    private ExpressionNode ParseEqv()
    {
        var left = ParseOr();
        while (_ts.Check(TokenKind.KwEqv))
        {
            var op = _ts.Consume();
            var right = ParseOr();
            left = new BinaryExpressionNode(op.Line, op.Column, left, op.Kind, right);
        }
        return left;
    }

    private ExpressionNode ParseOr()
    {
        var left = ParseAnd();
        while (_ts.Peek().Kind is TokenKind.KwOr or TokenKind.KwXor)
        {
            var op = _ts.Consume();
            var right = ParseAnd();
            left = new BinaryExpressionNode(op.Line, op.Column, left, op.Kind, right);
        }
        return left;
    }

    private ExpressionNode ParseAnd()
    {
        var left = ParseNot();
        while (_ts.Check(TokenKind.KwAnd))
        {
            var op = _ts.Consume();
            var right = ParseNot();
            left = new BinaryExpressionNode(op.Line, op.Column, left, op.Kind, right);
        }
        return left;
    }

    private ExpressionNode ParseNot()
    {
        if (_ts.Check(TokenKind.KwNot))
        {
            var op = _ts.Consume();
            return new UnaryExpressionNode(op.Line, op.Column, op.Kind, ParseNot());
        }
        return ParseComparison();
    }

    private ExpressionNode ParseComparison()
    {
        var left = ParseStringConcat();
        while (_ts.Peek().Kind is TokenKind.Equals or TokenKind.NotEqual
               or TokenKind.LessThan or TokenKind.GreaterThan
               or TokenKind.LessEqual or TokenKind.GreaterEqual
               or TokenKind.KwIs or TokenKind.KwLike)
        {
            var op = _ts.Consume();
            // TypeOf x Is T
            if (op.Kind == TokenKind.KwIs && left is TypeOfIsNode)
            {
                // already handled in ParsePrimary
                break;
            }
            var right = ParseStringConcat();
            left = new BinaryExpressionNode(op.Line, op.Column, left, op.Kind, right);
        }
        return left;
    }

    private ExpressionNode ParseStringConcat()
    {
        var left = ParseAddSub();
        while (_ts.Check(TokenKind.Ampersand))
        {
            var op = _ts.Consume();
            var right = ParseAddSub();
            left = new BinaryExpressionNode(op.Line, op.Column, left, op.Kind, right);
        }
        return left;
    }

    private ExpressionNode ParseAddSub()
    {
        var left = ParseMod();
        while (_ts.Peek().Kind is TokenKind.Plus or TokenKind.Minus)
        {
            var op = _ts.Consume();
            var right = ParseMod();
            left = new BinaryExpressionNode(op.Line, op.Column, left, op.Kind, right);
        }
        return left;
    }

    private ExpressionNode ParseMod()
    {
        var left = ParseIntDiv();
        while (_ts.Check(TokenKind.KwMod))
        {
            var op = _ts.Consume();
            var right = ParseIntDiv();
            left = new BinaryExpressionNode(op.Line, op.Column, left, op.Kind, right);
        }
        return left;
    }

    private ExpressionNode ParseIntDiv()
    {
        var left = ParseMulDiv();
        while (_ts.Check(TokenKind.Backslash))
        {
            var op = _ts.Consume();
            var right = ParseMulDiv();
            left = new BinaryExpressionNode(op.Line, op.Column, left, op.Kind, right);
        }
        return left;
    }

    private ExpressionNode ParseMulDiv()
    {
        var left = ParseUnary();
        while (_ts.Peek().Kind is TokenKind.Star or TokenKind.Slash)
        {
            var op = _ts.Consume();
            var right = ParseUnary();
            left = new BinaryExpressionNode(op.Line, op.Column, left, op.Kind, right);
        }
        return left;
    }

    private ExpressionNode ParseUnary()
    {
        if (_ts.Peek().Kind is TokenKind.Minus or TokenKind.Plus)
        {
            var op = _ts.Consume();
            return new UnaryExpressionNode(op.Line, op.Column, op.Kind, ParseUnary());
        }
        return ParseExponentiation();
    }

    private ExpressionNode ParseExponentiation()
    {
        var left = ParsePostfix();
        if (_ts.Check(TokenKind.Caret))
        {
            var op = _ts.Consume();
            var right = ParseUnary(); // right-associative
            return new BinaryExpressionNode(op.Line, op.Column, left, op.Kind, right);
        }
        return left;
    }

    private ExpressionNode ParsePostfix()
    {
        var expr = ParsePrimary();

        while (true)
        {
            if (_ts.Match(TokenKind.Dot))
            {
                var memberTok = ExpectIdentifierOrKeyword();
                expr = new MemberAccessNode(expr.Line, expr.Column, expr, memberTok.Text);
            }
            else if (_ts.Match(TokenKind.Bang))
            {
                var memberTok = _ts.Expect(TokenKind.Identifier);
                expr = new BangAccessNode(expr.Line, expr.Column, expr, memberTok.Text);
            }
            else if (_ts.Check(TokenKind.LParen))
            {
                var args = ParseArgumentListParenthesized();
                expr = new CallOrIndexNode(expr.Line, expr.Column, expr, args);
            }
            else
            {
                break;
            }
        }

        return expr;
    }

    private ExpressionNode ParsePrimary()
    {
        var t = _ts.Peek();

        switch (t.Kind)
        {
            case TokenKind.IntegerLiteral:
                _ts.Consume();
                return new IntegerLiteralNode(t.Line, t.Column, t.Text, ParseIntegerValue(t.Text));

            case TokenKind.DoubleLiteral:
                _ts.Consume();
                return new DoubleLiteralNode(t.Line, t.Column, t.Text,
                    double.Parse(t.Text.TrimEnd('%', '&', '!', '#', '@'), System.Globalization.CultureInfo.InvariantCulture));

            case TokenKind.StringLiteral:
                _ts.Consume();
                return new StringLiteralNode(t.Line, t.Column, t.Text, UnquoteString(t.Text));

            case TokenKind.BoolLiteral:
                _ts.Consume();
                return new BoolLiteralNode(t.Line, t.Column,
                    t.Text.Equals("True", StringComparison.OrdinalIgnoreCase));

            case TokenKind.DateLiteral:
                _ts.Consume();
                return new DateLiteralNode(t.Line, t.Column, t.Text);

            case TokenKind.KwNothing:
                _ts.Consume();
                return new NothingNode(t.Line, t.Column);

            case TokenKind.KwMe:
                _ts.Consume();
                return new MeNode(t.Line, t.Column);

            case TokenKind.KwNew:
            {
                _ts.Consume();
                var typeTok = ExpectIdentifierOrKeyword();
                return new NewObjectNode(t.Line, t.Column, typeTok.Text);
            }

            case TokenKind.KwTypeof:
            {
                _ts.Consume();
                var operand = ParsePostfix();
                _ts.Expect(TokenKind.KwIs);
                var typeTok = ExpectIdentifierOrKeyword();
                return new TypeOfIsNode(t.Line, t.Column, operand, typeTok.Text);
            }

            case TokenKind.LParen:
            {
                _ts.Consume();
                var inner = ParseExpression();
                _ts.Expect(TokenKind.RParen);
                return inner;
            }

            case TokenKind.Dot:
            {
                // .Member inside a With block
                _ts.Consume();
                var memberTok = ExpectIdentifierOrKeyword();
                return new WithMemberAccessNode(t.Line, t.Column, memberTok.Text);
            }

            case TokenKind.Identifier:
            case TokenKind.KwString: // String can appear as identifier in some contexts
                _ts.Consume();
                return new IdentifierNode(t.Line, t.Column, t.Text);

            default:
                // Some keywords can appear as identifiers (e.g. built-in function names)
                if (IsTypeName(t.Kind) || IsKeywordUsableAsIdentifier(t.Kind))
                {
                    _ts.Consume();
                    return new IdentifierNode(t.Line, t.Column, t.Text);
                }
                throw new ParseException($"Unexpected token '{t.Text}' in expression", t.Line, t.Column);
        }
    }

    private bool IsKeywordUsableAsIdentifier(TokenKind k) => k switch
    {
        // VB6 allows some keywords as identifiers (built-in functions etc.)
        TokenKind.KwDate or TokenKind.KwError or TokenKind.KwGet or TokenKind.KwLet
        or TokenKind.KwSet => true,
        _ => false
    };

    // ── Helpers ────────────────────────────────────────────────────────────

    /// <summary>
    /// Consume and return the current token, accepting either an Identifier or any
    /// keyword that commonly appears in an identifier position (procedure/variable names).
    /// </summary>
    private Token ExpectIdentifierOrKeyword()
    {
        var t = _ts.Peek();
        if (t.Kind == TokenKind.Identifier || IsTypeName(t.Kind) || IsKeywordUsableAsIdentifier(t.Kind)
            || t.Kind >= TokenKind.KwPublic) // any keyword can be used as identifier in VB6
        {
            return _ts.Consume();
        }
        throw new ParseException($"Expected identifier but found '{t.Text}'", t.Line, t.Column);
    }

    private static long ParseIntegerValue(string text)
    {
        if (text.StartsWith("&H", StringComparison.OrdinalIgnoreCase))
            return Convert.ToInt64(text.TrimEnd('%', '&', '!', '#', '@')[2..], 16);
        if (text.StartsWith("&O", StringComparison.OrdinalIgnoreCase))
            return Convert.ToInt64(text.TrimEnd('%', '&', '!', '#', '@')[2..], 8);
        return long.Parse(text.TrimEnd('%', '&', '!', '#', '@'));
    }
}
