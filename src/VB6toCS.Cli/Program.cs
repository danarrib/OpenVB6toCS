using VB6toCS.Core.Analysis;
using VB6toCS.Core.CodeGeneration;
using VB6toCS.Core.Lexing;
using VB6toCS.Core.Parsing;
using VB6toCS.Core.Parsing.Nodes;
using VB6toCS.Core.Projects;
using VB6toCS.Core.Transformation;

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
    var sourceEncoding = ResolveEncoding(options.EncodingName);
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

    // ── Loop 1: stages 1–3 ───────────────────────────────────────────────────
    // Process every file through lexing, parsing, and semantic analysis.
    // Files that pass stage 2 are collected into stage3List for the second loop.

    int errors = 0;
    int ok = 0;
    var stage3List = new List<(VbSourceFile Src, ModuleNode Module)>();

    Console.WriteLine($"Source files ({project.SourceFiles.Count}):");

    foreach (var src in project.SourceFiles)
    {
        if (!File.Exists(src.FullPath))
        {
            Console.WriteLine($"  [MISSING ] {src.Name}  — {src.FullPath}");
            errors++;
            continue;
        }

        string source = File.ReadAllText(src.FullPath, sourceEncoding);

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

        // Stage 2: validate Option Explicit and parse
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

        // Stage 3: semantic analysis
        if (options.UpToStage >= 3)
        {
            var analyser = new Analyser();
            module = analyser.Analyse(module);
            foreach (var d in analyser.Diagnostics)
                Console.WriteLine($"  [WARN    ] {src.Name}  — {d}");
        }

        stage3List.Add((src, module));
    }

    Console.WriteLine();

    if (errors > 0)
    {
        Console.WriteLine($"Completed with {errors} error(s). Output not written.");
        return 1;
    }

    // Stage 1–3 diagnostic-only exits
    if (options.UpToStage <= 2)
    {
        Console.WriteLine($"Stage {options.UpToStage} complete. {stage3List.Count} file(s) processed. No output written.");
        return 0;
    }

    if (options.UpToStage == 3)
    {
        Console.WriteLine($"Stage 3 complete. {stage3List.Count} file(s) analysed. No output written.");
        return 0;
    }

    // ── Cross-module default member map (used by code generator to expand default member calls) ─
    var defaultMemberMap = BuildDefaultMemberMap(stage3List.Select(r => r.Module));

    // ── Cross-module method parameter map (used by code generator to emit ref at call sites) ───
    var methodParamMap = BuildMethodParamMap(stage3List.Select(r => r.Module));

    // ── Cross-module method parameter TYPE map (built after Stage 4 for normalised types) ──────
    // Built later (after parsed is populated). Declared here for scope.
    IReadOnlyDictionary<string, IReadOnlyDictionary<string, IReadOnlyList<string?>>> methodParamTypeMap;

    // ── Cross-module global variable map (public fields from all modules → declared type) ──────
    var globalVarMap = BuildGlobalVarMap(stage3List.Select(r => r.Module));

    // ── Cross-module Collection type inference (between stages 3 and 4) ───────
    IReadOnlyDictionary<(string, string), string> collectionTypeMap;
    IReadOnlyDictionary<(string, string), VB6toCS.Core.Analysis.CollectionKind> collectionKindMap;
    {
        var inferrer = new CollectionTypeInferrer();
        inferrer.Analyse(stage3List.Select(r => r.Module));
        collectionTypeMap = inferrer.GetInferredTypes();
        collectionKindMap = inferrer.GetInferredKinds();
    }

    // ── Cross-module enum member map (used by code generator for qualification) ─
    var enumMemberMap = BuildEnumMemberMap(stage3List.Select(r => r.Module));

    // ── Set of field names whose declared type is an enum (suppresses int cast) ─
    var enumTypedFieldNames = BuildEnumTypedFieldNames(stage3List.Select(r => r.Module), enumMemberMap);

    // ── Loop 2: stage 4 — IR transformation with inferred collection types ────
    var parsed = new List<(VbSourceFile Src, ModuleNode Module)>();

    foreach (var (src, module3) in stage3List)
    {
        var transformer = new Transformer(collectionTypeMap, collectionKindMap);
        var module = transformer.Transform(module3);
        foreach (var d in transformer.Diagnostics)
            Console.WriteLine($"  [WARN    ] {src.Name}  — {d}");

        parsed.Add((src, module));
        ok++;
        Console.WriteLine($"  [OK      ] {src.Name}  ({Path.GetFileName(src.FullPath)})");
    }

    Console.WriteLine();

    // Stage 4: diagnostic only — no output files
    if (options.UpToStage == 4)
    {
        Console.WriteLine($"Stage 4 complete. {ok} file(s) transformed. No output written.");
        return 0;
    }

    // ── Method parameter TYPE map — built from Stage-4 normalised modules ───────────
    methodParamTypeMap = BuildMethodParamTypeMap(parsed.Select(p => p.Module));

    // ── Cross-module member type map: fieldName → C# type (all fields + UDT fields) ──
    // Built from Stage-4 modules. Used to resolve member-access types like obj.Val_IS
    // when the receiver's declared type is unknown (cross-module UDT fields).
    var allMemberTypeMap = BuildAllMemberTypeMap(parsed.Select(p => p.Module));

    // ── Set of field names (module fields + UDT fields) whose type is Dictionary<…> ─
    // Built from Stage-4 modules so types are already normalised by the Transformer.
    // Used by the code generator to append .Values when iterating over a Dictionary
    // that is accessed as a member (e.g. obj.ColecaoPedCobert → needs .Values).
    var dictionaryTypedFieldNames = BuildDictionaryTypedFieldNames(parsed.Select(p => p.Module));

    // Stages 5–6: write output
    string outputDir = Path.Combine(Path.GetDirectoryName(vbpPath)!, project.Name);
    Directory.CreateDirectory(outputDir);

    CsprojWriter.Write(project, outputDir, options.NoInterop);
    Console.WriteLine($"Written  : {project.Name}.csproj");

    foreach (var (src, module) in parsed)
    {
        bool isStatic = src.Kind == VbSourceKind.StaticModule;
        string csCode = CodeGenerator.Generate(module, isStatic, enumMemberMap, enumTypedFieldNames, defaultMemberMap, methodParamMap, globalVarMap, dictionaryTypedFieldNames, allMemberTypeMap, methodParamTypeMap);
        string csFile = Path.Combine(outputDir, module.Name + ".cs");
        File.WriteAllText(csFile, csCode, System.Text.Encoding.UTF8);
        Console.WriteLine($"Written  : {module.Name}.cs");
    }

    Console.WriteLine();
    Console.WriteLine($"Done. {parsed.Count} file(s) translated.");
    if (options.UpToStage == 6)
        Console.WriteLine("Note: Roslyn formatting (stage 6) is not yet implemented.");

    return 0;
}

// ── Single-file diagnostic mode ─────────────────────────────────────────────

static int RunSingleFile(string path, CliOptions options)
{
    var sourceEncoding = ResolveEncoding(options.EncodingName);
    string source = File.ReadAllText(path, sourceEncoding);

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

        if (options.UpToStage >= 4)
        {
            var transformer = new Transformer();
            module = transformer.Transform(module);
            foreach (var d in transformer.Diagnostics)
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

/// <summary>
/// Builds a cross-module method parameter map from all Stage-3 ASTs.
/// Maps each module/class name → method name → ordered list of ParameterMode values.
/// Used by the code generator to emit 'ref' at call sites where the callee declares ByRef.
/// </summary>
static IReadOnlyDictionary<string, IReadOnlyDictionary<string, IReadOnlyList<ParameterMode>>>
    BuildMethodParamMap(IEnumerable<ModuleNode> modules)
{
    var outer = new Dictionary<string, IReadOnlyDictionary<string, IReadOnlyList<ParameterMode>>>(
        StringComparer.OrdinalIgnoreCase);

    foreach (var module in modules)
    {
        var inner = new Dictionary<string, IReadOnlyList<ParameterMode>>(StringComparer.OrdinalIgnoreCase);
        foreach (var member in module.Members)
        {
            switch (member)
            {
                case SubNode s:
                    inner[s.Name] = s.Parameters.Select(p => p.Mode).ToArray();
                    break;
                case FunctionNode f:
                    inner[f.Name] = f.Parameters.Select(p => p.Mode).ToArray();
                    break;
                // Parameterized Property Get is emitted as a method — include its params too.
                case CsPropertyNode p when p.GetParameters.Count > 0:
                    inner[p.Name] = p.GetParameters.Select(pr => pr.Mode).ToArray();
                    break;
            }
        }
        outer[module.Name] = inner;
    }
    return outer;
}

/// <summary>
/// Builds a cross-module method parameter TYPE map from Stage-4 (normalised) ASTs.
/// Maps module name → method name → ordered list of C# parameter type strings (null = unknown).
/// Used by the code generator to coerce call-site arguments to the expected parameter types
/// (e.g. double argument where int is expected → (int)arg; enum → (int)arg).
/// </summary>
static IReadOnlyDictionary<string, IReadOnlyDictionary<string, IReadOnlyList<string?>>>
    BuildMethodParamTypeMap(IEnumerable<ModuleNode> modules)
{
    var outer = new Dictionary<string, IReadOnlyDictionary<string, IReadOnlyList<string?>>>(
        StringComparer.OrdinalIgnoreCase);

    foreach (var module in modules)
    {
        var inner = new Dictionary<string, IReadOnlyList<string?>>(StringComparer.OrdinalIgnoreCase);

        IReadOnlyList<string?> ParamTypes(IReadOnlyList<ParameterNode> parms) =>
            parms.Select(p => p.TypeRef?.TypeName).ToArray();

        foreach (var member in module.Members)
        {
            switch (member)
            {
                case SubNode s:      inner[s.Name] = ParamTypes(s.Parameters); break;
                case FunctionNode f: inner[f.Name] = ParamTypes(f.Parameters); break;
                case CsPropertyNode p when p.GetParameters.Count > 0:
                    inner[p.Name] = ParamTypes(p.GetParameters); break;
            }
        }
        outer[module.Name] = inner;
    }
    return outer;
}

/// <summary>
/// Builds a cross-module global variable map from all Stage-3 ASTs.
/// Maps every Public field name (from any module) to its declared VB6 type name.
/// Used to resolve qualified calls like `M46V999.Method()` where M46V999 is a public
/// field declared in another module and not visible in the local/module scope.
/// When the same name is declared in multiple modules with different types, it is excluded
/// (ambiguous — the caller must rely on the local scope lookup instead).
/// </summary>
static IReadOnlyDictionary<string, string> BuildGlobalVarMap(IEnumerable<ModuleNode> modules)
{
    var map       = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    var conflicts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

    foreach (var module in modules)
    {
        foreach (var member in module.Members)
        {
            if (member is not FieldNode f) continue;
            if (f.Access != AccessModifier.Public && f.Access != AccessModifier.Global) continue;
            foreach (var d in f.Declarators)
            {
                if (d.TypeRef == null) continue;
                if (!map.TryAdd(d.Name, d.TypeRef.TypeName))
                {
                    if (!map[d.Name].Equals(d.TypeRef.TypeName, StringComparison.OrdinalIgnoreCase))
                        conflicts.Add(d.Name);
                }
            }
        }
    }

    foreach (var name in conflicts)
        map.Remove(name);

    return map;
}

/// <summary>
/// Builds a cross-module default member map from all Stage-3 ASTs.
/// Maps each module name to the name of its default member (VB_UserMemId = 0).
/// </summary>
static IReadOnlyDictionary<string, string> BuildDefaultMemberMap(IEnumerable<ModuleNode> modules)
{
    var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    foreach (var module in modules)
        if (module.DefaultMemberName != null)
            map[module.Name] = module.DefaultMemberName;
    return map;
}

/// <summary>
/// Builds a set of field/property/UDT-field names (across all modules) whose
/// declared VB6 type is an enum type.  Used by the code generator to suppress
/// the (int) cast when both sides of a comparison are already enum-typed.
/// </summary>
static IReadOnlySet<string> BuildEnumTypedFieldNames(
    IEnumerable<ModuleNode> modules,
    IReadOnlyDictionary<string, (string ModuleName, string EnumName)> enumMemberMap)
{
    // Collect every enum TYPE name (not member name) from the map's values.
    var enumTypeNames = new HashSet<string>(
        enumMemberMap.Values.Select(v => v.EnumName),
        StringComparer.OrdinalIgnoreCase);

    var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

    foreach (var module in modules)
    {
        foreach (var member in module.Members)
        {
            switch (member)
            {
                case FieldNode f:
                    foreach (var d in f.Declarators)
                        if (d.TypeRef != null && enumTypeNames.Contains(d.TypeRef.TypeName))
                            result.Add(d.Name);
                    break;

                case CsPropertyNode p:
                    if (p.Type != null && enumTypeNames.Contains(p.Type.TypeName))
                        result.Add(p.Name);
                    break;

                case UdtNode u:
                    foreach (var f in u.Fields)
                        if (enumTypeNames.Contains(f.TypeRef.TypeName))
                            result.Add(f.Name);
                    break;
            }
        }
    }

    return result;
}

/// <summary>
/// Builds a set of field names (module-level fields, properties, and UDT fields) whose
/// normalised C# type is a <c>Dictionary&lt;…&gt;</c>.
/// Built from Stage-4 modules (after Transformer normalises Collection → Dictionary).
/// Used by the code generator to append <c>.Values</c> when a For Each iterates over a
/// Dictionary accessed as a member (e.g. <c>obj.ColecaoPedCobert</c>).
/// A name is included only if it is NEVER declared with a non-Dictionary type anywhere in
/// the project (avoids false positives for field names shared across different types).
/// </summary>
static IReadOnlySet<string> BuildDictionaryTypedFieldNames(IEnumerable<ModuleNode> modules)
{
    var result  = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    var nonDict = new HashSet<string>(StringComparer.OrdinalIgnoreCase); // conflict tracker

    foreach (var module in modules)
    {
        foreach (var member in module.Members)
        {
            switch (member)
            {
                case FieldNode f:
                    foreach (var d in f.Declarators)
                    {
                        if (IsDictionaryType(d.TypeRef?.TypeName)) result.Add(d.Name);
                        else if (d.TypeRef != null) nonDict.Add(d.Name);
                    }
                    break;

                case CsPropertyNode p:
                    if (IsDictionaryType(p.Type?.TypeName)) result.Add(p.Name);
                    else if (p.Type != null) nonDict.Add(p.Name);
                    break;

                case UdtNode u:
                    foreach (var fld in u.Fields)
                    {
                        if (IsDictionaryType(fld.TypeRef.TypeName)) result.Add(fld.Name);
                        else nonDict.Add(fld.Name);
                    }
                    break;
            }
        }
    }

    result.ExceptWith(nonDict); // remove names that appear with non-Dictionary types anywhere
    return result;
}

static bool IsDictionaryType(string? typeName) =>
    typeName != null && typeName.StartsWith("Dictionary<", StringComparison.OrdinalIgnoreCase);

/// <summary>
/// Builds a cross-module map of member name → C# type for ALL fields, properties,
/// and UDT fields across every module (built from Stage-4 normalised types).
/// Names that appear with conflicting types in different declarations are excluded
/// to avoid false positives. Used by the code generator to resolve the type of
/// member-access expressions like <c>obj.Val_IS</c> when <c>obj</c>'s type is unknown.
/// </summary>
static IReadOnlyDictionary<string, string> BuildAllMemberTypeMap(IEnumerable<ModuleNode> modules)
{
    var map       = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    var conflicts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

    void Add(string name, string? typeName)
    {
        if (string.IsNullOrEmpty(typeName) || conflicts.Contains(name)) return;
        if (!map.TryAdd(name, typeName) &&
            !map[name].Equals(typeName, StringComparison.OrdinalIgnoreCase))
        {
            conflicts.Add(name);
            map.Remove(name);
        }
    }

    foreach (var module in modules)
    {
        foreach (var member in module.Members)
        {
            switch (member)
            {
                case FieldNode f:
                    foreach (var d in f.Declarators)
                        Add(d.Name, d.TypeRef?.TypeName);
                    break;
                case CsPropertyNode p:
                    Add(p.Name, p.Type?.TypeName);
                    break;
                case UdtNode u:
                    foreach (var fld in u.Fields)
                        Add(fld.Name, fld.TypeRef.TypeName);
                    break;
            }
        }
    }

    return map;
}

/// <summary>
/// Builds a cross-module enum member map from all Stage-3 ASTs.
/// Maps each public enum member name to its (moduleName, enumName).
/// Names that appear in more than one module are excluded (ambiguous).
/// </summary>
static IReadOnlyDictionary<string, (string ModuleName, string EnumName)>
    BuildEnumMemberMap(IEnumerable<ModuleNode> modules)
{
    var map       = new Dictionary<string, (string, string)>(StringComparer.OrdinalIgnoreCase);
    var conflicts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

    foreach (var module in modules)
    {
        foreach (var member in module.Members)
        {
            if (member is not EnumNode e) continue;
            foreach (var em in e.Members)
            {
                if (!map.TryAdd(em.Name, (module.Name, e.Name)))
                {
                    // Member name already seen — mark ambiguous if from a different module
                    if (!map[em.Name].Item1.Equals(module.Name, StringComparison.OrdinalIgnoreCase))
                        conflicts.Add(em.Name);
                }
            }
        }
    }

    foreach (var name in conflicts)
        map.Remove(name);

    return map;
}

/// <summary>
/// Resolves the encoding to use for reading VB6 source files.
/// Priority: user-specified → windows-1252 (if available) → Latin-1 fallback.
/// </summary>
static System.Text.Encoding ResolveEncoding(string? name)
{
    if (name != null)
    {
        try
        {
            System.Text.Encoding.RegisterProvider(
                System.Text.CodePagesEncodingProvider.Instance);
            return System.Text.Encoding.GetEncoding(name);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Warning: Unknown encoding '{name}': {ex.Message}. Falling back to Latin-1.");
            return System.Text.Encoding.Latin1;
        }
    }

    // Try windows-1252 (requires the code pages provider on non-Windows platforms)
    try
    {
        System.Text.Encoding.RegisterProvider(
            System.Text.CodePagesEncodingProvider.Instance);
        return System.Text.Encoding.GetEncoding(1252);
    }
    catch
    {
        return System.Text.Encoding.Latin1;
    }
}
