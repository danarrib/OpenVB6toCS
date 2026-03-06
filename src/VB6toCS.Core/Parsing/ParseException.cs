namespace VB6toCS.Core.Parsing;

public sealed class ParseException : Exception
{
    public int Line { get; }
    public int Column { get; }

    public ParseException(string message, int line, int column)
        : base($"Parse error at {line}:{column}: {message}")
    {
        Line = line;
        Column = column;
    }
}
