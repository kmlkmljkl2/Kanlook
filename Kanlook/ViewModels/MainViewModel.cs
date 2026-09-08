using System.Collections.ObjectModel;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Kanlook.Services;

namespace Kanlook.ViewModels;

public sealed partial class MainViewModel : ObservableObject
{
    private const string RootOrderKey = "__root__";

    private readonly IOutlookService _outlook;
    private readonly BoardStateStore _boardStore = new();

    public ObservableCollection<MailFolderNodeVm> RootFolders { get; } = [];

    [ObservableProperty]
    private MailFolderNodeVm? _selectedFolder;

    [ObservableProperty]
    private object? _currentContent;

    [ObservableProperty]
    private PreviewPaneViewModel? _preview;

    [ObservableProperty]
    private GridLength _previewColumnWidth = new(0);

    [ObservableProperty]
    private string? _connectionErrorMessage;

    [ObservableProperty]
    private bool _isLoadingFolder;

    public MainViewModel(IOutlookService outlook)
    {
        _outlook = outlook;

        try
        {
            _outlook.Connect();
            foreach (var root in _outlook.BuildFolderTree())
                RootFolders.Add(new MailFolderNodeVm(root, MoveFolder));
            ApplyFolderOrder(RootFolders, RootOrderKey);
        }
        catch (Exception ex)
        {
            ConnectionErrorMessage =
                $"Couldn't connect to Outlook. Make sure the classic desktop Outlook app is installed and try again. ({ex.Message})";
        }
    }

    partial void OnSelectedFolderChanged(MailFolderNodeVm? value)
    {
        if (value is null)
            return;

        IsLoadingFolder = true;
        try
        {
            var mails = _outlook.GetMailSummaries(value.StoreId, value.EntryId);

            (CurrentContent as IDisposable)?.Dispose();
            CurrentContent = value.IsSharedMailbox
                ? new SharedFolderViewModel(value.Name, mails, ShowPreview)
                : new KanbanBoardViewModel(
                    value.Name,
                    FolderKeyHelper.BuildKey(value.StoreId, value.EntryId),
                    value.StoreId,
                    value.EntryId,
                    mails,
                    _boardStore,
                    _outlook,
                    ShowPreview);
        }
        catch (Exception ex)
        {
            ConnectionErrorMessage = $"Couldn't load folder '{value.Name}': {ex.Message}";
        }
        finally
        {
            IsLoadingFolder = false;
        }
    }

    private void ApplyFolderOrder(ObservableCollection<MailFolderNodeVm> nodes, string parentKey)
    {
        var order = _boardStore.GetFolderOrder(parentKey);
        if (order is { Count: > 0 })
        {
            var sorted = nodes
                .OrderBy(n =>
                {
                    var i = order.IndexOf(n.EntryId);
                    return i < 0 ? int.MaxValue : i;
                })
                .ToList();

            nodes.Clear();
            foreach (var n in sorted)
                nodes.Add(n);
        }

        foreach (var n in nodes)
            ApplyFolderOrder(n.Children, FolderKeyHelper.BuildKey(n.StoreId, n.EntryId));
    }

    private void MoveFolder(MailFolderNodeVm node, int direction)
    {
        var siblings = node.Parent?.Children ?? RootFolders;
        var index = siblings.IndexOf(node);
        var newIndex = index + direction;
        if (newIndex < 0 || newIndex >= siblings.Count)
            return;

        siblings.Move(index, newIndex);

        var parentKey = node.Parent is null
            ? RootOrderKey
            : FolderKeyHelper.BuildKey(node.Parent.StoreId, node.Parent.EntryId);
        _boardStore.SetFolderOrder(parentKey, siblings.Select(s => s.EntryId).ToList());
    }

    private void ShowPreview(MailCardViewModel card)
    {
        Preview = new PreviewPaneViewModel(card.Summary, _outlook, ClosePreview);
        if (PreviewColumnWidth.Value <= 0)
            PreviewColumnWidth = new GridLength(420);
    }

    [RelayCommand]
    private void ClosePreview()
    {
        Preview = null;
        PreviewColumnWidth = new GridLength(0);
    }

    public void Shutdown()
    {
        (CurrentContent as IDisposable)?.Dispose();
        _outlook.Dispose();
    }
}
