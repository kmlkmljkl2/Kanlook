using CommunityToolkit.Mvvm.ComponentModel;
using Kanlook.Models;

namespace Kanlook.ViewModels;

public sealed partial class MailFolderNodeVm : ObservableObject
{
    public MailFolderNode Node { get; }

    public string Name => Node.Name;
    public string EntryId => Node.EntryId;
    public string StoreId => Node.StoreId;
    public bool IsSharedMailbox => Node.IsSharedMailbox;

    public List<MailFolderNodeVm> Children { get; }

    [ObservableProperty]
    private bool _isExpanded;

    public MailFolderNodeVm(MailFolderNode node)
    {
        Node = node;
        Children = node.Children.Select(c => new MailFolderNodeVm(c)).ToList();
    }
}
