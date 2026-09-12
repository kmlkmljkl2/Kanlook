using CommunityToolkit.Mvvm.ComponentModel;
using Kanlook.Services;

namespace Kanlook.ViewModels;

/// <summary>Backs the settings dialog. Every change is persisted and applied immediately.</summary>
public sealed class SettingsViewModel : ObservableObject
{
    private readonly BoardStateStore _store;
    private readonly Action _onGroupingChanged;

    public SettingsViewModel(BoardStateStore store, Action onGroupingChanged)
    {
        _store = store;
        _onGroupingChanged = onGroupingChanged;
    }

    public bool GroupByConversation
    {
        get => _store.Settings.GroupByConversation;
        set
        {
            if (_store.Settings.GroupByConversation == value)
                return;

            _store.Settings.GroupByConversation = value;
            _store.Save();
            OnPropertyChanged();
            _onGroupingChanged();
        }
    }
}
