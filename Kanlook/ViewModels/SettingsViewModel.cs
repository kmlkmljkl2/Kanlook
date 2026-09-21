using System.Runtime.CompilerServices;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Kanlook.Models;
using Kanlook.Services;

namespace Kanlook.ViewModels;

/// <summary>Backs the settings dialog. Every change is persisted and applied immediately.</summary>
public sealed partial class SettingsViewModel : ObservableObject
{
    private readonly BoardStateStore _store;

    /// <summary>Re-lays out the open folder - the card settings change its order.</summary>
    private readonly Action _onLayoutChanged;

    public SettingsViewModel(BoardStateStore store, Action onLayoutChanged)
    {
        _store = store;
        _onLayoutChanged = onLayoutChanged;
    }

    public bool GroupByConversation
    {
        get => _store.Settings.GroupByConversation;
        set
        {
            if (_store.Settings.GroupByConversation == value)
                return;

            _store.Settings.GroupByConversation = value;
            Persist();
            _onLayoutChanged();
        }
    }

    public bool PinHighPriority
    {
        get => _store.Settings.PinHighPriority;
        set
        {
            if (_store.Settings.PinHighPriority == value)
                return;

            _store.Settings.PinHighPriority = value;
            Persist();
            _onLayoutChanged();
        }
    }

    /// <summary>
    /// How many of a folder's newest mails to load. Typed into a text box, so anything that isn't a
    /// number leaves the setting alone rather than resetting it mid-keystroke.
    /// </summary>
    public string MailsPerFolderText
    {
        get => _store.Settings.MailsPerFolder.ToString();
        set
        {
            if (!int.TryParse(value, out var parsed) || parsed == _store.Settings.MailsPerFolder)
                return;

            _store.Settings.MailsPerFolder = parsed;
            _store.Save();
        }
    }

    /// <summary>Spelled out, because the box silently clamps what's typed into it.</summary>
    public string MailsPerFolderHint =>
        $"Between {AppSettings.MinMailsPerFolder} and {AppSettings.MaxMailsPerFolder}. " +
        "Takes effect the next time a folder is opened.";

    /// <summary>
    /// The three segments of the theme picker. Booleans rather than the enum itself, so the picker
    /// can highlight the chosen one without a value converter.
    /// </summary>
    public bool IsThemeSystem => _store.Settings.Theme == AppTheme.System;

    public bool IsThemeLight => _store.Settings.Theme == AppTheme.Light;

    public bool IsThemeDark => _store.Settings.Theme == AppTheme.Dark;

    [RelayCommand]
    private void SetTheme(AppTheme theme)
    {
        if (_store.Settings.Theme == theme)
            return;

        _store.Settings.Theme = theme;
        ThemeManager.Apply(theme);
        _store.Save();

        OnPropertyChanged(nameof(IsThemeSystem));
        OnPropertyChanged(nameof(IsThemeLight));
        OnPropertyChanged(nameof(IsThemeDark));
    }

    private void Persist([CallerMemberName] string? propertyName = null)
    {
        _store.Save();
        OnPropertyChanged(propertyName);
    }
}
