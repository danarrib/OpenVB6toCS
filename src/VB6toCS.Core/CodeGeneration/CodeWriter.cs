using System.Text;

namespace VB6toCS.Core.CodeGeneration;

/// <summary>Simple indenting text writer for C# code generation.</summary>
internal sealed class CodeWriter
{
    private readonly StringBuilder _sb = new();
    private int _indent;
    private const string IndentUnit = "    ";

    public void Write(string text) => _sb.Append(text);

    public void WriteLine(string text = "")
    {
        if (text.Length == 0)
            _sb.AppendLine();
        else
            _sb.Append(CurrentIndent).AppendLine(text);
    }

    public void OpenBrace()
    {
        WriteLine("{");
        _indent++;
    }

    public void CloseBrace(string suffix = "")
    {
        _indent--;
        WriteLine("}" + suffix);
    }

    private string CurrentIndent => _indent == 0 ? "" : string.Concat(
        System.Linq.Enumerable.Repeat(IndentUnit, _indent));

    public override string ToString() => _sb.ToString();
}
