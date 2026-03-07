namespace VB6toCS.Core.Analysis;

/// <summary>
/// Two-level symbol table: module scope (fields, methods, enums, etc.) and
/// procedure scope (parameters, local Dim variables).
/// </summary>
public sealed class SymbolTable
{
    private readonly Dictionary<string, Symbol> _module =
        new(StringComparer.OrdinalIgnoreCase);

    private Dictionary<string, Symbol>? _local;

    // ── Procedure context ───────────────────────────────────────────────────

    /// <summary>Name of the currently-open procedure (canonical casing).</summary>
    public string? CurrentProcedureName { get; private set; }

    /// <summary>
    /// Name to detect as a return-value assignment, or null if the current
    /// procedure has no return value (Sub, Property Let/Set).
    /// Set to the function/property name for Function and Property Get bodies.
    /// </summary>
    public string? CurrentReturnName { get; private set; }

    // ── Module scope ────────────────────────────────────────────────────────

    public void AddModule(Symbol s) => _module[s.Name] = s;

    // ── Procedure scope ─────────────────────────────────────────────────────

    public void EnterProcedure(string procedureName, string? returnName)
    {
        CurrentProcedureName = procedureName;
        CurrentReturnName = returnName;
        _local = new Dictionary<string, Symbol>(StringComparer.OrdinalIgnoreCase);
    }

    public void ExitProcedure()
    {
        CurrentProcedureName = null;
        CurrentReturnName = null;
        _local = null;
    }

    public void AddLocal(Symbol s) => _local![s.Name] = s;

    // ── Lookup ──────────────────────────────────────────────────────────────

    /// <summary>
    /// Looks up a name in the innermost scope first (local, then module).
    /// Returns null if not found.
    /// </summary>
    public Symbol? Lookup(string name)
    {
        if (_local != null && _local.TryGetValue(name, out var s)) return s;
        _module.TryGetValue(name, out s);
        return s;
    }
}
