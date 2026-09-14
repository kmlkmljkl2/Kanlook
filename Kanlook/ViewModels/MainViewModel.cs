using System.Collections.ObjectModel;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Kanlook.Models;
using Kanlook.Services;

namespace Kanlook.ViewModels;

public sealed partial class MainViewModel : ObservableObject
{
    private const string RootOrderKey = "__root__";

    private readonly IOutlookService _outlook;
    private readonly BoardStateStore _boardStore = new();
    private readonly AttachmentIndex _attachmentIndex = new();
    private readonly AttachmentIndexer _indexer;

    /// <summary>
    /// Shared by every board: the sent history of a store is the same wherever it's shown, and
    /// reading Sent Items is expensive enough to be worth keeping across folder switches.
    /// </summary>
    private readonly SentMailIndex _sentMail;

    public ObservableCollection<MailFolderNodeVm> RootFolders { get; } = [];

    public SettingsViewModel Settings { get; }

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
        _indexer = new AttachmentIndexer(outlook, _attachmentIndex);
        _sentMail = new SentMailIndex(outlook);
        Settings = new SettingsViewModel(_boardStore, RelayoutCurrentFolder);

        try
        {
            _outlook.Connect();
            CategoryPalette.Load(_outlook.GetCategoryColors());

            foreach (var root in _outlook.BuildFolderTree())
            {
                // Open each mailbox straight away so its folders are there without a click.
                RootFolders.Add(new MailFolderNodeVm(root, MoveFolder) { IsExpanded = true });
            }

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
                ? new SharedFolderViewModel(value.Name, mails, _boardStore, ShowPreview, DeleteCard, SetRead)
                : new KanbanBoardViewModel(
                    value.Name,
                    FolderKeyHelper.BuildKey(value.StoreId, value.EntryId),
                    value.StoreId,
                    value.EntryId,
                    mails,
                    _boardStore,
                    _outlook,
                    _attachmentIndex,
                    _indexer,
                    _sentMail,
                    ShowPreview,
                    DeleteCard,
                    SetRead);
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

    private void ShowPreview(MailSummary summary)
    {
        Preview = new PreviewPaneViewModel(summary, _outlook, ClosePreview, DeleteMails, SetRead);
        if (PreviewColumnWidth.Value <= 0)
            PreviewColumnWidth = new GridLength(420);
    }

    private void DeleteCard(MailCardViewModel card) => DeleteMails(card.Messages);

    /// <summary>
    /// Moves mails to Outlook's Deleted Items folder and takes them off the board right away, rather
    /// than waiting for the next sync to notice.
    /// </summary>
    private void DeleteMails(IReadOnlyList<MailSummary> mails)
    {
        var deleted = new List<string>();
        foreach (var mail in mails)
        {
            try
            {
                _outlook.DeleteMail(mail.StoreId, mail.EntryId);
                deleted.Add(mail.EntryId);
            }
            catch (Exception ex)
            {
                ConnectionErrorMessage = $"Couldn't delete '{mail.Subject}': {ex.Message}";
            }
        }

        if (deleted.Count == 0)
            return;

        switch (CurrentContent)
        {
            case KanbanBoardViewModel board:
                board.DropMails(deleted);
                break;
            case SharedFolderViewModel shared:
                shared.DropMails(deleted);
                break;
        }

        if (Preview is not null && deleted.Contains(Preview.EntryId, StringComparer.OrdinalIgnoreCase))
            ClosePreview();
    }

    /// <summary>
    /// Marks mails read or unread in Outlook and mirrors it onto the board right away. Takes the
    /// whole list, so a conversation tile carries its entire thread.
    /// </summary>
    private void SetRead(IReadOnlyList<MailSummary> mails, bool isRead)
    {
        var changed = new List<string>();
        foreach (var mail in mails.Where(m => m.IsRead != isRead))
        {
            try
            {
                _outlook.SetRead(mail.StoreId, mail.EntryId, isRead);
                mail.IsRead = isRead;
                changed.Add(mail.EntryId);
            }
            catch (Exception ex)
            {
                ConnectionErrorMessage =
                    $"Couldn't mark '{mail.Subject}' as {(isRead ? "read" : "unread")}: {ex.Message}";
            }
        }

        if (changed.Count == 0)
            return;

        switch (CurrentContent)
        {
            case KanbanBoardViewModel board:
                board.RefreshMailState(changed);
                break;
            case SharedFolderViewModel shared:
                shared.RefreshMailState(changed);
                break;
        }

        if (Preview is not null && changed.Contains(Preview.EntryId, StringComparer.OrdinalIgnoreCase))
            Preview.RefreshReadState();
    }

    /// <summary>Re-lays out the open folder after a setting changed how its cards group or sort.</summary>
    private void RelayoutCurrentFolder()
    {
        switch (CurrentContent)
        {
            case KanbanBoardViewModel board:
                board.RebuildCards();
                break;
            case SharedFolderViewModel shared:
                shared.RebuildCards();
                break;
        }
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
        _attachmentIndex.Save();
        _outlook.Dispose();
    }
}
