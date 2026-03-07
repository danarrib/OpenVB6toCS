using VB6toCS.Core.Parsing.Nodes;

namespace VB6toCS.Core.Analysis;

public enum SymbolKind
{
    Field,
    Constant,
    Enum,
    EnumMember,
    Udt,
    Sub,
    Function,
    Property,
    Parameter,
    LocalVar,
}

public sealed record Symbol(
    string Name,        // canonical casing — taken from the declaration site
    SymbolKind Kind,
    TypeRefNode? Type,
    bool IsArray);
