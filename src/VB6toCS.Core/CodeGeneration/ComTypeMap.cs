namespace VB6toCS.Core.CodeGeneration;

/// <summary>
/// Lookup table for COM library member access patterns that should be emitted
/// as C# indexer syntax <c>obj.Member[arg]</c> rather than function call syntax
/// <c>obj.Member(arg)</c>.
///
/// VB6 uses identical syntax for function calls and indexed collection access.
/// For project-internal types this is resolved by the Stage-3 analyser (which
/// has the symbol table).  For external COM types there is no symbol table, so
/// the code generator consults this static map instead.
///
/// Keys are matched case-insensitively.  The owner type name is tried both in
/// its fully-qualified form (e.g. "ADODB.Recordset") and its unqualified form
/// (e.g. "Recordset"), so declarations like <c>As Recordset</c> and
/// <c>As ADODB.Recordset</c> are both handled.
/// </summary>
internal static class ComTypeMap
{
    // (ownerTypeName, memberName) → true if member access with args should use [] not ()
    private static readonly HashSet<(string, string)> IndexedMembers =
        new(TupleComparer.OrdinalIgnoreCase)
        {
            // ── ADODB ────────────────────────────────────────────────────────────
            // ADODB.Recordset
            ("ADODB.Recordset",   "Fields"),
            ("Recordset",         "Fields"),
            ("ADODB.Recordset",   "Properties"),
            ("Recordset",         "Properties"),
            // ADODB.Fields (the collection itself — also accessed as rs.Fields("name") etc.)
            ("ADODB.Fields",      "Item"),
            ("Fields",            "Item"),
            // ADODB.Properties collection
            ("ADODB.Properties",  "Item"),
            ("Properties",        "Item"),
            // ADODB.Command
            ("ADODB.Command",     "Parameters"),
            ("Command",           "Parameters"),
            // ADODB.Parameters collection
            ("ADODB.Parameters",  "Item"),
            ("Parameters",        "Item"),
            // ADODB.Errors collection
            ("ADODB.Errors",      "Item"),
            ("Errors",            "Item"),

            // ── DAO 3.x ──────────────────────────────────────────────────────────
            ("DAO.Recordset",     "Fields"),
            ("DAO.Recordset",     "Properties"),
            ("DAO.Fields",        "Item"),
            ("DAO.Database",      "TableDefs"),
            ("DAO.Database",      "QueryDefs"),
            ("DAO.Database",      "Relations"),
            ("DAO.TableDef",      "Fields"),
            ("DAO.QueryDef",      "Parameters"),

            // ── MSXML ────────────────────────────────────────────────────────────
            ("MSXML2.IXMLDOMNodeList",  "item"),
            ("IXMLDOMNodeList",         "item"),
            ("MSXML2.IXMLDOMNamedNodeMap", "item"),
            ("IXMLDOMNamedNodeMap",     "item"),
        };

    /// <summary>
    /// Returns true when <paramref name="memberName"/> on a receiver of type
    /// <paramref name="ownerType"/> should be emitted with indexer brackets.
    /// </summary>
    public static bool IsIndexedMember(string ownerType, string memberName)
    {
        if (IndexedMembers.Contains((ownerType, memberName)))
            return true;

        // Also try the unqualified type name (strip leading namespace prefix).
        int dot = ownerType.LastIndexOf('.');
        if (dot >= 0)
        {
            string unqualified = ownerType[(dot + 1)..];
            if (IndexedMembers.Contains((unqualified, memberName)))
                return true;
        }

        return false;
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private sealed class TupleComparer : IEqualityComparer<(string, string)>
    {
        public static readonly TupleComparer OrdinalIgnoreCase = new();

        public bool Equals((string, string) x, (string, string) y) =>
            StringComparer.OrdinalIgnoreCase.Equals(x.Item1, y.Item1) &&
            StringComparer.OrdinalIgnoreCase.Equals(x.Item2, y.Item2);

        public int GetHashCode((string, string) obj) =>
            HashCode.Combine(
                StringComparer.OrdinalIgnoreCase.GetHashCode(obj.Item1),
                StringComparer.OrdinalIgnoreCase.GetHashCode(obj.Item2));
    }
}
