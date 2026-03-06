using VB6toCS.Core.Parsing;

namespace VB6toCS.Core.Projects;

public enum VbSourceKind { Class, StaticModule }

public sealed record VbSourceFile(
    string Name,
    string FullPath,
    VbSourceKind Kind);

public sealed record VbComReference(
    string Description,
    Guid Guid,
    int VersionMajor,
    int VersionMinor,
    int Lcid);

public sealed record VbSkippedFile(
    string Name,
    string FullPath,
    string Reason);

public sealed record VbProject(
    string Name,
    string VbpPath,
    IReadOnlyList<VbSourceFile> SourceFiles,
    IReadOnlyList<VbComReference> ComReferences,
    IReadOnlyList<VbSkippedFile> SkippedFiles)
{
    public string VbpDirectory => Path.GetDirectoryName(VbpPath)!;
}
