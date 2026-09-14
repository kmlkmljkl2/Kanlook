using Kanlook.Services;

namespace Kanlook.ViewModels;

/// <summary>An Outlook category shown on a mail card.</summary>
public sealed record CategoryChip(string Name, string ColorHex)
{
    public static CategoryChip For(string categoryName) =>
        new(categoryName, CategoryPalette.ColorHexFor(categoryName));
}
