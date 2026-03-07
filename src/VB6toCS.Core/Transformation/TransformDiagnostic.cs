using VB6toCS.Core.Analysis;

namespace VB6toCS.Core.Transformation;

public sealed record TransformDiagnostic(
    DiagnosticSeverity Severity,
    string Message,
    int Line,
    int Column)
{
    public override string ToString() =>
        $"{Severity} at {Line}:{Column}: {Message}";
}
