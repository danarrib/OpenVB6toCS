namespace VB6toCS.Core.Analysis;

/// <summary>
/// Classifies how a VB6 <c>Collection</c> field is used, so the code generator
/// can emit the most appropriate C# collection type.
/// </summary>
public enum CollectionKind
{
    /// <summary>
    /// Every observed <c>.Add item, key</c> call site provides a string key.
    /// Translated to <c>Dictionary&lt;string, T&gt;</c>.
    /// </summary>
    Dictionary,

    /// <summary>
    /// No observed <c>.Add</c> call site provides a key argument.
    /// Translated to <c>List&lt;T&gt;</c>.
    /// </summary>
    List,

    /// <summary>
    /// Mixed usage (some Add calls with key, some without) or no Add calls observed.
    /// Translated to <c>Collection&lt;T&gt;</c> with a REVIEW comment.
    /// </summary>
    Collection,
}
