using System.Collections;
using System.Windows;
using Kanlook.Models;
using Microsoft.Win32;

namespace Kanlook.Services;

/// <summary>
/// Swaps the theme dictionary in the app's resources. Everything outside Themes/ reaches the brushes
/// with <c>DynamicResource</c>, so replacing the dictionary repaints views that are already on
/// screen - brushes loaded from XAML are frozen, which rules out recolouring them in place.
/// </summary>
public static class ThemeManager
{
    private const string ThemeFolderMarker = "/Themes/Theme.";

    private static AppTheme _preference = AppTheme.System;
    private static bool _listeningToWindows;

    /// <summary>The theme actually on screen, with <see cref="AppTheme.System"/> already resolved.</summary>
    public static AppTheme Current { get; private set; } = AppTheme.Light;

    /// <summary>
    /// Paints the app in <paramref name="theme"/>. With <see cref="AppTheme.System"/> the choice
    /// comes from Windows and keeps following it, so changing the mode in Windows settings arrives
    /// here without a restart.
    /// </summary>
    public static void Apply(AppTheme theme)
    {
        _preference = theme;

        if (theme == AppTheme.System)
            StartListeningToWindows();

        Load(theme == AppTheme.System ? ReadWindowsTheme() : theme);
    }

    private static void Load(AppTheme resolved)
    {
        if (Application.Current is not { } app)
            return;

        var loaded = new ResourceDictionary { Source = UriFor(resolved) };
        var merged = app.Resources.MergedDictionaries;

        var index = IndexOfTheme(merged);
        if (index < 0)
            merged.Insert(0, loaded);
        else
            merged[index] = loaded;

        Current = resolved;
        Verify(loaded, resolved);
    }

    private static int IndexOfTheme(IList<ResourceDictionary> merged)
    {
        for (var i = 0; i < merged.Count; i++)
        {
            if (merged[i].Source is { } source &&
                source.OriginalString.Contains(ThemeFolderMarker, StringComparison.OrdinalIgnoreCase))
                return i;
        }

        return -1;
    }

    /// <summary>
    /// Fails loudly when a theme is missing a key the light one has. Without this a dropped brush
    /// would only show up as an invisible control somewhere, and only in one theme.
    /// </summary>
    private static void Verify(ResourceDictionary loaded, AppTheme resolved)
    {
        if (resolved == AppTheme.Light)
            return;

        var expected = new ResourceDictionary { Source = UriFor(AppTheme.Light) };
        var missing = expected.Keys
            .Cast<object>()
            .Where(key => !loaded.Contains(key))
            .Select(key => key.ToString())
            .ToList();

        if (missing.Count > 0)
            throw new InvalidOperationException(
                $"Theme.{resolved}.xaml is missing: {string.Join(", ", missing)}");
    }

    private static Uri UriFor(AppTheme theme) =>
        new($"pack://application:,,,/Kanlook;component/Themes/Theme.{theme}.xaml");

    private static AppTheme ReadWindowsTheme()
    {
        try
        {
            // 0 means "use the dark app mode". Absent on editions that never had the setting.
            var value = Registry.GetValue(
                @"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
                "AppsUseLightTheme",
                1);

            return value is int and 0 ? AppTheme.Dark : AppTheme.Light;
        }
        catch (Exception)
        {
            return AppTheme.Light; // no registry access - light is the safer guess
        }
    }

    private static void StartListeningToWindows()
    {
        if (_listeningToWindows)
            return;

        SystemEvents.UserPreferenceChanged += (_, e) =>
        {
            if (e.Category == UserPreferenceCategory.General && _preference == AppTheme.System)
                Application.Current?.Dispatcher.Invoke(() => Load(ReadWindowsTheme()));
        };

        _listeningToWindows = true;
    }
}
