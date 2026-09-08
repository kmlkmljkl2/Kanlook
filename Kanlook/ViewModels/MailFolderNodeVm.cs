using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Kanlook.Models;

namespace Kanlook.ViewModels;

public sealed partial class MailFolderNodeVm : ObservableObject
{
    public MailFolderNode Node { get; }
    public MailFolderNodeVm? Parent { get; }

    public string Name => Node.Name;
    public string EntryId => Node.EntryId;
    public string StoreId => Node.StoreId;
    public bool IsSharedMailbox => Node.IsSharedMailbox;

    public ObservableCollection<MailFolderNodeVm> Children { get; }

    [ObservableProperty]
    private bool _isExpanded;

    private readonly Action<MailFolderNodeVm, int> _onMove;

    public MailFolderNodeVm(MailFolderNode node, Action<MailFolderNodeVm, int> onMove, MailFolderNodeVm? parent = null)
    {
        Node = node;
        Parent = parent;
        _onMove = onMove;
        Children = new ObservableCollection<MailFolderNodeVm>(
            node.Children.Select(c => new MailFolderNodeVm(c, onMove, this)));
    }

    [RelayCommand]
    private void MoveUp() => _onMove(this, -1);

    [RelayCommand]
    private void MoveDown() => _onMove(this, 1);
}
