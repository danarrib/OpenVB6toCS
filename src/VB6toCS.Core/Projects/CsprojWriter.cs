namespace VB6toCS.Core.Projects;

/// <summary>
/// Generates a SDK-style .csproj for a translated VB6 project.
/// </summary>
public static class CsprojWriter
{
    public static void Write(VbProject project, string outputDir, bool noInterop = false)
    {
        // Determine which COM libraries need <COMReference> entries.
        // Known full-mapping libraries are excluded — their types are translated
        // to native .NET and no interop assembly is needed.
        // When noInterop is set, all <COMReference> entries are suppressed entirely.
        var interopRefs = noInterop
            ? []
            : project.ComReferences.Where(r => !IsFullyMapped(r)).ToList();

        string csprojPath = Path.Combine(outputDir, project.Name + ".csproj");
        using var w = new StreamWriter(csprojPath);

        w.WriteLine("<Project Sdk=\"Microsoft.NET.Sdk\">");
        w.WriteLine();
        w.WriteLine("  <PropertyGroup>");
        w.WriteLine("    <TargetFramework>net8.0</TargetFramework>");
        w.WriteLine("    <OutputType>Library</OutputType>");
        w.WriteLine($"    <RootNamespace>{XmlEscape(project.Name)}</RootNamespace>");
        w.WriteLine($"    <AssemblyName>{XmlEscape(project.Name)}</AssemblyName>");
        w.WriteLine("    <Nullable>enable</Nullable>");
        w.WriteLine("    <ImplicitUsings>enable</ImplicitUsings>");
        w.WriteLine("  </PropertyGroup>");

        if (interopRefs.Count > 0)
        {
            w.WriteLine();
            w.WriteLine("  <!-- COM Interop references (no native .NET mapping found) -->");
            w.WriteLine("  <!-- Review each one and replace with a .NET equivalent when possible -->");
            w.WriteLine("  <ItemGroup>");
            foreach (var r in interopRefs)
            {
                string name = SanitizeName(r.Description);
                w.WriteLine($"    <!-- {XmlEscape(r.Description)} -->");
                w.WriteLine($"    <COMReference Include=\"{XmlEscape(name)}\">");
                w.WriteLine($"      <Guid>{{{r.Guid}}}</Guid>");
                w.WriteLine($"      <VersionMajor>{r.VersionMajor}</VersionMajor>");
                w.WriteLine($"      <VersionMinor>{r.VersionMinor}</VersionMinor>");
                w.WriteLine($"      <Lcid>{r.Lcid}</Lcid>");
                w.WriteLine("      <WrapperTool>tlbimp</WrapperTool>");
                w.WriteLine("      <Isolated>false</Isolated>");
                w.WriteLine("    </COMReference>");
            }
            w.WriteLine("  </ItemGroup>");
        }

        w.WriteLine();
        w.WriteLine("</Project>");
    }

    // ── COM mapping knowledge ───────────────────────────────────────────────

    /// <summary>
    /// Returns true when every type from this COM library has a native .NET
    /// equivalent and no interop assembly is needed in the output project.
    /// </summary>
    private static bool IsFullyMapped(VbComReference r)
    {
        // OLE Automation / stdole2 — basic COM plumbing; not referenced in generated C#
        if (r.Guid == new Guid("00020430-0000-0000-C000-000000000046")) return true;

        // Scripting Runtime — Dictionary maps to Dictionary<K,V>; FSO maps to System.IO
        if (r.Guid == new Guid("420B2830-E718-11CF-893D-00A0C9054228")) return true;

        return false;
    }

    // ── Helpers ────────────────────────────────────────────────────────────

    private static string XmlEscape(string s) =>
        s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");

    /// <summary>
    /// Derives a short identifier from a COM library description.
    /// e.g. "Microsoft ActiveX Data Objects 2.0 Library" → "ADODB"
    /// Falls back to the raw description with spaces removed.
    /// </summary>
    private static string SanitizeName(string description)
    {
        // Well-known COM library short names
        if (description.Contains("ActiveX Data Objects", StringComparison.OrdinalIgnoreCase))
            return "ADODB";
        if (description.Contains("DAO", StringComparison.OrdinalIgnoreCase))
            return "DAO";
        if (description.Contains("Soap", StringComparison.OrdinalIgnoreCase))
            return "MSSOAPLib";
        if (description.Contains("XML", StringComparison.OrdinalIgnoreCase))
            return "MSXML2";

        // Generic fallback: keep only alphanumeric characters
        return new string(description.Where(char.IsLetterOrDigit).ToArray());
    }
}
