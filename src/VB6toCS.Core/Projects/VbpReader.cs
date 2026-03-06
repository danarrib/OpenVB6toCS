namespace VB6toCS.Core.Projects;

/// <summary>
/// Parses a VB6 project file (.vbp) into a <see cref="VbProject"/>.
/// </summary>
public static class VbpReader
{
    public static VbProject Read(string vbpPath)
    {
        vbpPath = Path.GetFullPath(vbpPath);
        string dir = Path.GetDirectoryName(vbpPath)!;

        var sourceFiles = new List<VbSourceFile>();
        var comRefs = new List<VbComReference>();
        var skipped = new List<VbSkippedFile>();

        string? nameEntry = null;
        string? exeName = null;

        foreach (string raw in File.ReadLines(vbpPath))
        {
            string line = raw.Trim();
            if (line.Length == 0) continue;

            int eq = line.IndexOf('=');
            if (eq < 0) continue;

            string key = line[..eq].Trim();
            string value = line[(eq + 1)..].Trim();

            switch (key.ToUpperInvariant())
            {
                case "CLASS":
                    if (TryParseSourceEntry(value, dir, out var name, out var path))
                        sourceFiles.Add(new VbSourceFile(name, path, VbSourceKind.Class));
                    break;

                case "MODULE":
                    if (TryParseSourceEntry(value, dir, out name, out path))
                        sourceFiles.Add(new VbSourceFile(name, path, VbSourceKind.StaticModule));
                    break;

                case "FORM":
                    if (TryParseSourceEntry(value, dir, out name, out path))
                        skipped.Add(new VbSkippedFile(name, path, "Form (.frm) — UI component, out of scope"));
                    break;

                case "USERCONTROL":
                    if (TryParseSourceEntry(value, dir, out name, out path))
                        skipped.Add(new VbSkippedFile(name, path, "UserControl (.ctl) — UI component, out of scope"));
                    break;

                case "PROPERTYPAGE":
                    if (TryParseSourceEntry(value, dir, out name, out path))
                        skipped.Add(new VbSkippedFile(name, path, "PropertyPage — UI component, out of scope"));
                    break;

                case "REFERENCE":
                    if (TryParseReference(value, out var comRef))
                        comRefs.Add(comRef);
                    break;

                case "NAME":
                    nameEntry = Unquote(value);
                    break;

                case "EXENAME32":
                    // "MyDLL.dll" or "MyDLL" — strip extension
                    exeName = Path.GetFileNameWithoutExtension(Unquote(value));
                    break;
            }
        }

        // Project name: Name= preferred, ExeName32= fallback, filename stem last resort
        string projectName =
            nameEntry ??
            exeName ??
            Path.GetFileNameWithoutExtension(vbpPath);

        return new VbProject(projectName, vbpPath, sourceFiles, comRefs, skipped);
    }

    // ── Helpers ────────────────────────────────────────────────────────────

    /// <summary>
    /// Parses  Name; relative\path\to\File.cls  into a resolved absolute path.
    /// Both  "Name; path"  and  "Name;path"  (no space) are accepted.
    /// </summary>
    private static bool TryParseSourceEntry(
        string value, string baseDir,
        out string name, out string fullPath)
    {
        name = "";
        fullPath = "";

        int semi = value.IndexOf(';');
        if (semi < 0) return false;

        name = value[..semi].Trim();
        string rel = value[(semi + 1)..].Trim();
        fullPath = Path.GetFullPath(Path.Combine(baseDir, rel));
        return true;
    }

    /// <summary>
    /// Parses  *\G{GUID}#major.minor#lcid#path#description
    /// </summary>
    private static bool TryParseReference(string value, out VbComReference result)
    {
        result = null!;

        // Strip leading  *\G  or  *\G  prefix
        if (value.StartsWith("*\\G", StringComparison.OrdinalIgnoreCase))
            value = value[3..];
        else if (value.StartsWith("*\\", StringComparison.OrdinalIgnoreCase))
            value = value[2..];

        string[] parts = value.Split('#');
        // parts: {GUID}  major.minor  lcid  path  description
        if (parts.Length < 5) return false;

        string guidStr = parts[0].Trim('{', '}');
        if (!Guid.TryParse(guidStr, out Guid guid)) return false;

        string[] ver = parts[1].Split('.');
        int major = ver.Length > 0 && int.TryParse(ver[0], out int mj) ? mj : 0;
        int minor = ver.Length > 1 && int.TryParse(ver[1], out int mn) ? mn : 0;
        int lcid  = int.TryParse(parts[2], out int lc) ? lc : 0;

        string description = parts[4].Trim();

        result = new VbComReference(description, guid, major, minor, lcid);
        return true;
    }

    private static string Unquote(string s) =>
        s.Length >= 2 && s[0] == '"' && s[^1] == '"' ? s[1..^1] : s;
}
