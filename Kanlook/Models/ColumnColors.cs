namespace Kanlook.Models;

/// <summary>
/// The colours a column can be given. One list, so the picker offers exactly what a fresh board is
/// built from - and so the set stays legible against both themes, which is why none of them are
/// near-white or near-black.
/// </summary>
public static class ColumnColors
{
    public const string Blue = "#5B8DEF";
    public const string Amber = "#F2A93B";
    public const string Purple = "#B87CE0";
    public const string Green = "#4CB782";
    public const string Red = "#E5566B";
    public const string Teal = "#2FB6C4";
    public const string Slate = "#8C93A6";
    public const string Pink = "#D6558C";

    public static IReadOnlyList<string> All { get; } =
        [Blue, Amber, Purple, Green, Red, Teal, Slate, Pink];

    /// <summary>The columns a board starts life with.</summary>
    public static IReadOnlyList<(string Name, string ColorHex)> DefaultColumns { get; } =
    [
        ("New", Blue),
        ("In Progress", Amber),
        ("Waiting", Purple),
        ("Done", Green),
    ];
}
