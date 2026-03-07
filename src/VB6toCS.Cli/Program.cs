using VB6toCS.Core.Analysis;
using VB6toCS.Core.Lexing;
using VB6toCS.Core.Parsing;
using VB6toCS.Core.Parsing.Nodes;
using VB6toCS.Core.Projects;

// ── Argument parsing ────────────────────────────────────────────────────────

if (args.Length == 0)
{
    Console.Error.WriteLine(CliOptions.UsageText);
    return 1;
}

var (options, parseError) = CliOptions.Parse(args);
if (options == null)
{
    Console.Error.WriteLine($"Error: {parseError}");
    Console.Error.WriteLine();
    Console.Error.WriteLine(CliOptions.UsageText);
    return 1;
}

var path = Path.GetFullPath(options.InputPath);
if (!File.Exists(path))
{
    Console.Error.WriteLine($"Error: File not found: {path}");
    return 1;
}

string ext = Path.GetExtension(path).ToLowerInvariant();
return ext switch
{
    ".vbp"        => RunProject(path, options),
    ".cls" or ".bas" => RunSingleFile(path, options),
    _ => PrintError($"Unsupported file type '{ext}'. Pass a .vbp, .cls, or .bas file.")
};

// ── Project mode ────────────────────────────────────────────────────────────

static int RunProject(string vbpPath, CliOptions options)
{
    // Stage 0: read VBP project file
    VbProject project;
    try
    {
        project = VbpReader.Read(vbpPath);
    }
    catch (Exception ex)
    {
        return PrintError($"Failed to read project file: {ex.Message}");
    }

    PrintProjectHeader(project, options);

    // Stage 0: list files and stop
    if (options.UpToStage == 0)
    {
        Console.WriteLine($"Source files ({project.SourceFiles.Count}):");
        foreach (var src in project.SourceFiles)
            Console.WriteLine($"  {src.Name}  ({Path.GetFileName(src.FullPath)})  [{src.Kind}]");
        Console.WriteLine();
        Console.WriteLine($"{project.SourceFiles.Count} file(s) listed. Stage 0 — no further processing.");
        return 0;
    }

    // Stages 1+: process each source file through the pipeline
    int errors = 0;
    int ok = 0;
    var parsed = new List<(VbSourceFile Src, ModuleNode Module)>();

    Console.WriteLine($"Source files ({project.SourceFiles.Count}):");

    foreach (var src in project.SourceFiles)
    {
        if (!File.Exists(src.FullPath))
        {
            Console.WriteLine($"  [MISSING ] {src.Name}  — {src.FullPath}");
            errors++;
            continue;
        }

        string source = File.ReadAllText(src.FullPath);

        // Stage 1: tokenize
        List<Token> tokens;
        try
        {
            tokens = new Lexer(source).Tokenize();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  [ERROR   ] {src.Name}  — Lex error: {ex.Message}");
            errors++;
            continue;
        }

        if (options.UpToStage == 1)
        {
            Console.WriteLine($"  [OK      ] {src.Name}  ({Path.GetFileName(src.FullPath)}) — {tokens.Count} tokens");
            ok++;
            continue;
        }

        // Stage 2+: validate Option Explicit and parse
        ModuleNode module;
        try
        {
            TokenListValidator.RequireOptionExplicit(tokens, src.FullPath);
            module = new Parser(tokens, src.FullPath).Parse();
        }
        catch (VB6SourceException ex)
        {
            Console.WriteLine($"  [REJECTED] {src.Name}  — {ex.Message}");
            errors++;
            continue;
        }
        catch (ParseException ex)
        {
            Console.WriteLine($"  [ERROR   ] {src.Name}  — {ex.Message}");
            errors++;
            continue;
        }

        // Stage 3+: semantic analysis
        if (options.UpToStage >= 3)
        {
            var analyser = new Analyser();
            module = analyser.Analyse(module);
            foreach (var d in analyser.Diagnostics)
                Console.WriteLine($"  [WARN    ] {src.Name}  — {d}");
        }

        parsed.Add((src, module));
        ok++;
        Console.WriteLine($"  [OK      ] {src.Name}  ({Path.GetFileName(src.FullPath)})");
    }

    Console.WriteLine();

    if (errors > 0)
    {
        Console.WriteLine($"Completed with {errors} error(s). Output not written.");
        return 1;
    }

    // Stage 1–2: diagnostic only, no output files
    if (options.UpToStage <= 2)
    {
        Console.WriteLine($"Stage {options.UpToStage} complete. {ok} file(s) processed. No output written.");
        return 0;
    }

    // Stage 3: diagnostic only — no output files
    if (options.UpToStage == 3)
    {
        Console.WriteLine($"Stage 3 complete. {ok} file(s) analysed. No output written.");
        return 0;
    }

    // Stage 4: not yet implemented
    if (options.UpToStage == 4)
    {
        Console.WriteLine("Stage 4 (IR transformation) is not yet implemented.");
        Console.WriteLine($"Pipeline stopped after stage 3. {ok} file(s) analysed. No output written.");
        return 0;
    }

    // Stages 5–6: write output
    string outputDir = Path.Combine(Path.GetDirectoryName(vbpPath)!, project.Name);
    Directory.CreateDirectory(outputDir);

    CsprojWriter.Write(project, outputDir, options.NoInterop);
    Console.WriteLine($"Written  : {project.Name}.csproj");

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
    Console.WriteLine("Note: C# code generation is not yet implemented — .cs files are placeholders.");
    if (options.UpToStage == 6)
        Console.WriteLine("Note: Roslyn formatting (stage 6) is not yet implemented.");

    return 0;
}

// ── Single-file diagnostic mode ─────────────────────────────────────────────

static int RunSingleFile(string path, CliOptions options)
{
    string source = File.ReadAllText(path);

    List<Token> tokens;
    try
    {
        tokens = new Lexer(source).Tokenize();
    }
    catch (Exception ex)
    {
        return PrintError($"Lex error: {ex.Message}");
    }

    // Stage 0 or 1: print token table
    if (options.UpToStage <= 1)
    {
        Console.WriteLine($"Tokenizing: {path}");
        Console.WriteLine(new string('-', 72));
        Console.WriteLine($"{"Line",-6} {"Col",-6} {"Kind",-28} Text");
        Console.WriteLine(new string('-', 72));
        foreach (var t in tokens.Where(t => t.Kind != TokenKind.EndOfFile))
        {
            string text = t.Kind == TokenKind.Newline
                ? "<newline>"
                : t.Text.Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");
            Console.WriteLine($"{t.Line,-6} {t.Column,-6} {t.Kind,-28} {text}");
        }
        Console.WriteLine(new string('-', 72));
        Console.WriteLine($"{tokens.Count} tokens.");
        return 0;
    }

    // Stage 2+: validate and parse
    try
    {
        TokenListValidator.RequireOptionExplicit(tokens, path);
    }
    catch (VB6SourceException ex)
    {
        return PrintError(ex.Message);
    }

    try
    {
        var module = new Parser(tokens, path).Parse();

        if (options.UpToStage >= 3)
        {
            var analyser = new Analyser();
            module = analyser.Analyse(module);
            foreach (var d in analyser.Diagnostics)
                Console.Error.WriteLine($"Warning: {d}");
        }

        Console.WriteLine($"Parsing: {path}");
        Console.WriteLine(new string('-', 60));
        AstPrinter.Print(module, Console.Out);
        Console.WriteLine(new string('-', 60));
        Console.WriteLine($"Parsed successfully: module '{module.Name}' [{module.Kind}], {module.Members.Count} top-level members.");
    }
    catch (ParseException ex)
    {
        return PrintError($"Parse error: {ex.Message}");
    }

    return 0;
}

// ── Helpers ─────────────────────────────────────────────────────────────────

static void PrintProjectHeader(VbProject project, CliOptions options)
{
    Console.WriteLine($"Project : {project.Name}");
    Console.WriteLine($"Source  : {project.VbpPath}");
    if (options.UpToStage >= 5)
        Console.WriteLine($"Output  : {Path.Combine(project.VbpDirectory, project.Name)}");
    Console.WriteLine();

    if (options.NoInterop)
        Console.WriteLine("Note: -NoInterop — COM references omitted from .csproj; TODO comments emitted at usage sites.");

    if (project.ComReferences.Count > 0)
    {
        Console.WriteLine($"COM references ({project.ComReferences.Count}):");
        foreach (var r in project.ComReferences)
            Console.WriteLine($"  {r.Description}  [v{r.VersionMajor}.{r.VersionMinor}]");
        Console.WriteLine();
    }

    if (project.SkippedFiles.Count > 0)
    {
        Console.WriteLine($"Skipped ({project.SkippedFiles.Count} — out of scope):");
        foreach (var s in project.SkippedFiles)
            Console.WriteLine($"  {s.Name}  — {s.Reason}");
        Console.WriteLine();
    }
}

static int PrintError(string message)
{
    Console.Error.WriteLine($"Error: {message}");
    return 1;
}
