using VB6toCS.Core.Lexing;

if (args.Length == 0)
{
    Console.Error.WriteLine("Usage: vb6tocs <file.cls|file.bas>");
    return 1;
}

var path = args[0];
if (!File.Exists(path))
{
    Console.Error.WriteLine($"File not found: {path}");
    return 1;
}

var source = File.ReadAllText(path);
var tokens = new Lexer(source).Tokenize();

try
{
    TokenListValidator.RequireOptionExplicit(tokens, path);
}
catch (VB6SourceException ex)
{
    Console.Error.WriteLine($"Error: {ex.Message}");
    return 1;
}

Console.WriteLine($"Tokenizing: {path}");
Console.WriteLine(new string('-', 60));

foreach (var tok in tokens)
    Console.WriteLine(tok);

Console.WriteLine(new string('-', 60));
Console.WriteLine($"{tokens.Count} tokens.");
return 0;
