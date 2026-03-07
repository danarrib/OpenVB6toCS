namespace VB6toCS.Core.Analysis;

public enum DiagnosticSeverity { Warning, Error }

public sealed record AnalysisDiagnostic(
    DiagnosticSeverity Severity,
    string Message,
    int Line,
    int Column)
{
    public override string ToString() =>
        $"{Severity} at {Line}:{Column}: {Message}";
}
