using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Kanlook.Models;
using Kanlook.Services;

namespace Kanlook.ViewModels;

/// <summary>What a tree node needs from the window around it, so a node can load its own children.</summary>
public sealed class FolderTreeContext
{
    /// <summary>Reads one folder's immediate subfolders from Outlook.</summary>
    public required Func<string, string, Task<List<MailFolderNode>>> LoadChildren { get; init; }

    /// <summary>Puts a freshly loaded level into the order the user dragged it into.</summary>
    public required Action<ObservableCollection<MailFolderNodeVm>, string> ApplyOrder { get; init; }

    /// <summary>Moves a node one place up or down among its siblings, and remembers it.</summary>
    public required Action<MailFolderNodeVm, int> Move { get; init; }
}

public sealed partial class MailFolderNodeVm : ObservableObject
{
    public MailFolderNode Node { get; }
    public MailFolderNodeVm? Parent { get; }

    public string Name => Node.Name;
    public string EntryId => Node.EntryId;
    public string StoreId => Node.StoreId;
    public bool IsSharedMailbox => Node.IsSharedMailbox;

    /// <summary>
    /// The "Loading…" stand-in a collapsed folder holds so the tree draws an expander for it. It is
    /// not a folder and must not be treated as one - opening it would ask Outlook for nothing.
    /// </summary>
    public bool IsPlaceholder { get; private init; }

    public ObservableCollection<MailFolderNodeVm> Children { get; } = [];

    [ObservableProperty]
    private bool _isExpanded;

    private readonly FolderTreeContext _context;

    /// <summary>
    /// The one read of this folder's children, kept so a second caller joins it rather than either
    /// starting another or - worse - returning as if the children were already there. Expanding the
    /// node and asking for them outright both happen, and can happen at once.
    /// </summary>
    private Task? _childrenLoad;

    public MailFolderNodeVm(MailFolderNode node, FolderTreeContext context, MailFolderNodeVm? parent = null)
    {
        Node = node;
        Parent = parent;
        _context = context;

        if (node.HasChildren)
            Children.Add(CreatePlaceholder(context));
    }

    private static MailFolderNodeVm CreatePlaceholder(FolderTreeContext context) =>
        new(new MailFolderNode { EntryId = "", StoreId = "", Name = "Loading…" }, context)
        {
            IsPlaceholder = true,
        };

    /// <summary>Subfolders are read the first time the node is opened, never before.</summary>
    partial void OnIsExpandedChanged(bool value)
    {
        if (value)
            _ = LoadChildrenAsync();
    }

    public Task LoadChildrenAsync() =>
        IsPlaceholder ? Task.CompletedTask : _childrenLoad ??= LoadChildrenCoreAsync();

    private async Task LoadChildrenCoreAsync()
    {
        List<MailFolderNode> children;
        try
        {
            children = await _context.LoadChildren(StoreId, EntryId);
        }
        catch (Exception)
        {
            // A folder we can't reach simply shows nothing under it. Forget the attempt, so
            // collapsing and reopening tries again in case Outlook was only momentarily away.
            _childrenLoad = null;
            Children.Clear();
            return;
        }

        Children.Clear();
        foreach (var child in children)
            Children.Add(new MailFolderNodeVm(child, _context, this));

        _context.ApplyOrder(Children, FolderKeyHelper.BuildKey(StoreId, EntryId));
    }

    [RelayCommand]
    private void MoveUp() => _context.Move(this, -1);

    [RelayCommand]
    private void MoveDown() => _context.Move(this, 1);
}
