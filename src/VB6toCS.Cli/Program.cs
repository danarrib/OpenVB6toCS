using VB6toCS.Core.Lexing;
using VB6toCS.Core.Parsing;
using VB6toCS.Core.Parsing.Nodes;
using VB6toCS.Core.Projects;

if (args.Length == 0)
{
    Console.Error.WriteLine("Usage: vb6tocs <file.vbp|file.cls|file.bas>");
    return 1;
}

var path = Path.GetFullPath(args[0]);
if (!File.Exists(path))
{
    Console.Error.WriteLine($"File not found: {path}");
    return 1;
}

string ext = Path.GetExtension(path).ToLowerInvariant();

return ext switch
{
    ".vbp" => RunProject(path),
    ".cls" or ".bas" => RunSingleFile(path),
    _ => Error($"Unsupported file type '{ext}'. Pass a .vbp, .cls, or .bas file.")
};

// ── Project mode ───────────────────────────────────────────────────────────

static int RunProject(string vbpPath)
{
    VbProject project;
    try
    {
        project = VbpReader.Read(vbpPath);
    }
    catch (Exception ex)
    {
        return Error($"Failed to read project file: {ex.Message}");
    }

    string outputDir = Path.Combine(Path.GetDirectoryName(vbpPath)!, project.Name);

    Console.WriteLine($"Project : {project.Name}");
    Console.WriteLine($"Source  : {vbpPath}");
    Console.WriteLine($"Output  : {outputDir}");
    Console.WriteLine();

    // Report COM references
    if (project.ComReferences.Count > 0)
    {
        Console.WriteLine($"COM references ({project.ComReferences.Count}):");
        foreach (var r in project.ComReferences)
            Console.WriteLine($"  {r.Description}  [v{r.VersionMajor}.{r.VersionMinor}]");
        Console.WriteLine();
    }

    // Report skipped files
    if (project.SkippedFiles.Count > 0)
    {
        Console.WriteLine($"Skipped ({project.SkippedFiles.Count} — out of scope):");
        foreach (var s in project.SkippedFiles)
            Console.WriteLine($"  {s.Name}  — {s.Reason}");
        Console.WriteLine();
    }

    // Parse each in-scope source file
    Console.WriteLine($"Source files ({project.SourceFiles.Count}):");
    int errors = 0;
    var parsed = new List<(VbSourceFile File, ModuleNode Module)>();

    foreach (var src in project.SourceFiles)
    {
        if (!File.Exists(src.FullPath))
        {
            Console.WriteLine($"  [MISSING ] {src.Name}  ({Path.GetFileName(src.FullPath)})");
            errors++;
            continue;
        }

        try
        {
            var tokens = new Lexer(File.ReadAllText(src.FullPath)).Tokenize();
            TokenListValidator.RequireOptionExplicit(tokens, src.FullPath);
            var module = new Parser(tokens, src.FullPath).Parse();
            parsed.Add((src, module));
            Console.WriteLine($"  [OK      ] {src.Name}  ({Path.GetFileName(src.FullPath)})");
        }
        catch (VB6SourceException ex)
        {
            Console.WriteLine($"  [REJECTED] {src.Name}  — {ex.Message}");
            errors++;
        }
        catch (ParseException ex)
        {
            Console.WriteLine($"  [ERROR   ] {src.Name}  — {ex.Message}");
            errors++;
        }
    }

    Console.WriteLine();

    if (errors > 0)
    {
        Console.WriteLine($"Completed with {errors} error(s). Output not written.");
        return 1;
    }

    // Write output
    Directory.CreateDirectory(outputDir);

    // Write .csproj
    CsprojWriter.Write(project, outputDir);
    Console.WriteLine($"Written  : {project.Name}.csproj");

    // Placeholder .cs files (code generation not yet implemented)
    foreach (var (src, module) in parsed)
    {
        string csFile = Path.Combine(outputDir, module.Name + ".cs");
        File.WriteAllText(csFile,
            $"// TODO: code generation not yet implemented\n" +
            $"// Source: {Path.GetFileName(src.FullPath)}\n" +
            $"// Module: {module.Name} [{module.Kind}]\n");
        Console.WriteLine($"Written  : {module.Name}.cs  (placeholder)");
    }

    Console.WriteLine();
    Console.WriteLine($"Done. {parsed.Count} file(s) parsed successfully.");
    Console.WriteLine($"Note: C# code generation is not yet implemented — .cs files are placeholders.");
    return 0;
}

// ── Single-file diagnostic mode ────────────────────────────────────────────

static int RunSingleFile(string path)
{
    var source = File.ReadAllText(path);
    var tokens = new Lexer(source).Tokenize();

    try
    {
        TokenListValidator.RequireOptionExplicit(tokens, path);
    }
    catch (VB6SourceException ex)
    {
        return Error(ex.Message);
    }

    try
    {
        var parser = new Parser(tokens, path);
        var module = parser.Parse();

        Console.WriteLine($"Parsing: {path}");
        Console.WriteLine(new string('-', 60));
        AstPrinter.Print(module, Console.Out);
        Console.WriteLine(new string('-', 60));
        Console.WriteLine($"Parsed successfully: module '{module.Name}' [{module.Kind}], {module.Members.Count} top-level members.");
    }
    catch (ParseException ex)
    {
        return Error($"Parse error: {ex.Message}");
    }

    return 0;
}

// ── Helpers ────────────────────────────────────────────────────────────────

static int Error(string message)
{
    Console.Error.WriteLine($"Error: {message}");
    return 1;
}
