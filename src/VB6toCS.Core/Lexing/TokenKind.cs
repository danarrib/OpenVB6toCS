namespace VB6toCS.Core.Lexing;

public enum TokenKind
{
    // ── Special ──────────────────────────────────────────────────────────
    EndOfFile,
    Newline,        // \n or \r\n (statement separator)
    Colon,          // : (inline statement separator)

    // ── Literals ─────────────────────────────────────────────────────────
    StringLiteral,  // "hello"
    IntegerLiteral, // 42 or &H1F (hex) or &O17 (octal)
    DoubleLiteral,  // 3.14
    DateLiteral,    // #2024-01-01#
    BoolLiteral,    // True / False

    // ── Identifier / Unknown name ─────────────────────────────────────────
    Identifier,

    // ── Comments ──────────────────────────────────────────────────────────
    Comment,        // ' comment text  (retained for output)

    // ── Declaration keywords ───────────────────────────────────────────────
    KwPublic,
    KwPrivate,
    KwFriend,
    KwGlobal,
    KwDim,
    KwConst,
    KwStatic,
    KwAs,
    KwNew,
    KwSet,
    KwLet,
    KwCall,
    KwOption,
    KwExplicit,
    KwAttribute,

    // ── Type declaration keywords ──────────────────────────────────────────
    KwEnum,
    KwType,
    KwImplements,

    // ── Procedure keywords ─────────────────────────────────────────────────
    KwFunction,
    KwSub,
    KwProperty,
    KwGet,
    KwEnd,        // "End" — followed by Sub/Function/Property/Enum/Type/If/Select/With
    KwExit,       // Exit Function / Exit Sub

    // ── Parameter modifiers ────────────────────────────────────────────────
    KwByVal,
    KwByRef,
    KwOptional,
    KwParamArray,

    // ── Control flow ──────────────────────────────────────────────────────
    KwIf,
    KwThen,
    KwElse,
    KwElseIf,
    KwEndIf,     // "End If" — lexed as single token for convenience
    KwSelect,
    KwCase,
    KwFor,
    KwEach,
    KwIn,
    KwTo,
    KwStep,
    KwNext,
    KwWhile,
    KwWend,
    KwDo,
    KwLoop,
    KwUntil,
    KwWith,
    KwGoTo,
    KwGoSub,
    KwReturn,

    // ── Error handling ─────────────────────────────────────────────────────
    KwOn,
    KwError,
    KwResume,

    // ── Object keywords ────────────────────────────────────────────────────
    KwNothing,
    KwMe,
    KwIs,
    KwLike,
    KwTypeof,

    // ── Built-in data type names ───────────────────────────────────────────
    KwBoolean,
    KwByte,
    KwInteger,
    KwLong,
    KwSingle,
    KwDouble,
    KwCurrency,
    KwDecimal,
    KwDate,
    KwString,
    KwObject,
    KwVariant,

    // ── Logical / bitwise operators (keywords) ─────────────────────────────
    KwAnd,
    KwOr,
    KwNot,
    KwXor,
    KwEqv,
    KwImp,
    KwMod,

    // ── Module header keywords (VB6 class file boilerplate) ────────────────
    KwVersion,
    KwBegin,
    KwMultiUse,

    // ── Arithmetic / comparison operators (symbols) ────────────────────────
    Plus,           // +
    Minus,          // -
    Star,           // *
    Slash,          // /
    Backslash,      // \ (integer division)
    Caret,          // ^ (exponentiation)
    Ampersand,      // & (string concat)
    Equals,         // =
    NotEqual,       // <>
    LessThan,       // <
    GreaterThan,    // >
    LessEqual,      // <=
    GreaterEqual,   // >=

    // ── Punctuation ───────────────────────────────────────────────────────
    LParen,         // (
    RParen,         // )
    Comma,          // ,
    Dot,            // .
    Bang,           // !  (default member access)
    Hash,           // #  (file number prefix / date delimiter)
    At,             // @
    Underscore,     // _ (only meaningful mid-line; line-continuation is consumed by lexer)
}
