namespace VB6toCS.Core.Lexing;

public static class TokenListValidator
{
    /// <summary>
    /// Enforces that the token stream contains "Option Explicit" at module level.
    /// Throws if missing — the user must fix the VB6 source before translating.
    /// </summary>
    public static void RequireOptionExplicit(IReadOnlyList<Token> tokens, string filePath)
    {
        for (int i = 0; i < tokens.Count - 1; i++)
        {
            if (tokens[i].Kind == TokenKind.KwOption &&
                tokens[i + 1].Kind == TokenKind.KwExplicit)
                return;
        }

        throw new VB6SourceException(
            $"{filePath}: 'Option Explicit' is required but was not found. " +
            "Add 'Option Explicit' to the top of the file and try again.");
    }
}

public class VB6SourceException(string message) : Exception(message);
