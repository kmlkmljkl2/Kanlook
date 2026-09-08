using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Kanlook.Models;

namespace Kanlook.ViewModels;

public sealed partial class KanbanColumnViewModel : ObservableObject
{
    public string Id { get; }
    public int Order { get; set; }
    public KanbanBoardViewModel Board { get; }

    [ObservableProperty]
    private string _name;

    [ObservableProperty]
    private string _colorHex;

    [ObservableProperty]
    private bool _isEditingName;

    public ObservableCollection<MailCardViewModel> Cards { get; } = [];

    private readonly Action<KanbanColumnViewModel> _onRemove;
    private readonly Action<KanbanColumnViewModel, int> _onMove;
    private readonly Action _onChanged;

    public KanbanColumnViewModel(
        KanbanColumnDefinition definition,
        Action<KanbanColumnViewModel> onRemove,
        Action<KanbanColumnViewModel, int> onMove,
        Action onChanged,
        KanbanBoardViewModel board)
    {
        Id = definition.Id;
        Order = definition.Order;
        _name = definition.Name;
        _colorHex = definition.ColorHex;
        _onRemove = onRemove;
        _onMove = onMove;
        _onChanged = onChanged;
        Board = board;
    }

    [RelayCommand]
    private void BeginRename() => IsEditingName = true;

    [RelayCommand]
    private void CommitRename()
    {
        if (string.IsNullOrWhiteSpace(Name))
            Name = "Untitled";
        IsEditingName = false;
        _onChanged();
    }

    [RelayCommand]
    private void Remove() => _onRemove(this);

    [RelayCommand]
    private void MoveLeft() => _onMove(this, -1);

    [RelayCommand]
    private void MoveRight() => _onMove(this, 1);
}
