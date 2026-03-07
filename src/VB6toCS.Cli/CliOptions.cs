/// <summary>Parsed command-line options for the vb6tocs tool.</summary>
public sealed record CliOptions(
    string InputPath,
    bool NoInterop,
    int UpToStage)
{
    public const int MaxStage = 6;

    public static (CliOptions? Options, string? Error) Parse(string[] args)
    {
        string? inputPath = null;
        bool noInterop = false;
        int upToStage = MaxStage;

        for (int i = 0; i < args.Length; i++)
        {
            var arg = args[i];

            if (arg.Equals("-NoInterop", StringComparison.OrdinalIgnoreCase))
            {
                noInterop = true;
            }
            else if (arg.Equals("-UpToStage", StringComparison.OrdinalIgnoreCase))
            {
                if (i + 1 >= args.Length)
                    return (null, $"-UpToStage requires a value between 0 and {MaxStage}.");

                if (!int.TryParse(args[++i], out int stage) || stage < 0 || stage > MaxStage)
                    return (null, $"-UpToStage value must be an integer 0–{MaxStage}, got '{args[i]}'.");

                upToStage = stage;
            }
            else if (arg.StartsWith('-'))
            {
                return (null, $"Unknown option '{arg}'.");
            }
            else if (inputPath == null)
            {
                inputPath = arg;
            }
            else
            {
                return (null, $"Unexpected argument '{arg}'.");
            }
        }

        if (inputPath == null)
            return (null, "No input file specified.");

        return (new CliOptions(inputPath, noInterop, upToStage), null);
    }

    public static string UsageText => $"""
        Usage: vb6tocs [options] <file.vbp|file.cls|file.bas>

        Options:
          -NoInterop          Omit <COMReference> entries from the generated .csproj.
                              Use this when targeting Linux / non-Windows environments
                              where COM Interop is unavailable. TODO comments are emitted
                              at COM usage sites instead.

          -UpToStage <0–{MaxStage}>    Run the pipeline up to the specified stage (default: {MaxStage}).
                              0 — Parse .vbp and list source files
                              1 — Tokenize (lex) each file
                              2 — Parse each file into an AST
                              3 — Semantic analysis        (not yet implemented)
                              4 — IR transformation        (not yet implemented)
                              5 — C# code generation       (writes placeholders for now)
                              6 — Roslyn formatting        (not yet implemented)
        """;
}
