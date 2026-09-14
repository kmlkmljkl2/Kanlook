namespace Kanlook.Services;

/// <summary>
/// Outlook's category colours, looked up by name. Loaded once after connecting; kept here rather
/// than threaded through every card so mail cards can colour their chips without extra plumbing.
/// </summary>
public static class CategoryPalette
{
    private const string FallbackHex = "#8C93A6";

    private static IReadOnlyDictionary<string, string> _colors = new Dictionary<string, string>();

    public static void Load(IReadOnlyDictionary<string, string> colors) => _colors = colors;

    public static string ColorHexFor(string categoryName) =>
        _colors.TryGetValue(categoryName, out var hex) ? hex : FallbackHex;
}
