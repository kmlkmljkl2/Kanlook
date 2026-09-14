using CommunityToolkit.Mvvm.ComponentModel;
using Kanlook.Services;

namespace Kanlook.ViewModels;

/// <summary>Backs the settings dialog. Every change is persisted and applied immediately.</summary>
public sealed class SettingsViewModel : ObservableObject
{
    private readonly BoardStateStore _store;

    /// <summary>Re-lays out the open folder - each of these settings changes its card order.</summary>
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
            Applied();
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
            Applied();
        }
    }

    private void Applied([System.Runtime.CompilerServices.CallerMemberName] string? propertyName = null)
    {
        _store.Save();
        OnPropertyChanged(propertyName);
        _onLayoutChanged();
    }
}
