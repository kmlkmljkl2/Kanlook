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
    private readonly BoardStateStore _boardStore;
    private readonly AttachmentIndex _attachmentIndex = new();
    private readonly AttachmentIndexer _indexer;
    private readonly FolderTreeContext _folderTree;

    /// <summary>
    /// Shared by every board: a mail's body is the same wherever it's shown, and reading one costs a
    /// round-trip, so it's worth keeping across folder switches.
    /// </summary>
    private readonly MailBodyLoader _bodyLoader;

    /// <summary>
    /// Shared by every board: the sent history of a store is the same wherever it's shown, and
    /// reading Sent Items is expensive enough to be worth keeping across folder switches.
    /// </summary>
    private readonly SentMailIndex _sentMail;

    /// <summary>
    /// Which folder request is the current one. Opening a folder is asynchronous, so a click while
    /// another folder is still loading has to be able to tell the stale result to stay quiet.
    /// </summary>
    private int _openGeneration;

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

    /// <summary>The sidebar's width, as the user drags the splitter. Remembered on exit.</summary>
    [ObservableProperty]
    private GridLength _sidebarWidth;

    [ObservableProperty]
    private string? _connectionErrorMessage;

    /// <summary>Whether the folder tree is still being read - the sidebar says so while it is.</summary>
    [ObservableProperty]
    private bool _isLoadingFolders = true;

    public MainViewModel(IOutlookService outlook, BoardStateStore boardStore)
    {
        _outlook = outlook;
        _boardStore = boardStore;
        _indexer = new AttachmentIndexer(outlook, _attachmentIndex);
        _bodyLoader = new MailBodyLoader(outlook);
        _sentMail = new SentMailIndex(outlook);
        Settings = new SettingsViewModel(_boardStore, RelayoutCurrentFolder);

        _folderTree = new FolderTreeContext
        {
            LoadChildren = _outlook.GetChildFoldersAsync,
            ApplyOrder = ApplyFolderOrder,
            Move = MoveFolder,
        };

        _sidebarWidth = new GridLength(_boardStore.Settings.SidebarWidth);

        // Before anything is shown, so the first frame is already in the right colours.
        ThemeManager.Apply(_boardStore.Settings.Theme);
    }

    /// <summary>
    /// Connects to Outlook and puts the mailboxes in the sidebar. Awaited by the window rather than
    /// done in the constructor: on a large mailbox this is seconds of Outlook round-trips, and the
    /// window should already be up and painted while they happen.
    /// </summary>
    public async Task InitializeAsync()
    {
        try
        {
            await _outlook.ConnectAsync();
            CategoryPalette.Load(await _outlook.GetCategoryColorsAsync());

            foreach (var root in await _outlook.GetStoreRootsAsync())
                RootFolders.Add(new MailFolderNodeVm(root, _folderTree));

            ApplyFolderOrder(RootFolders, RootOrderKey);
        }
        catch (Exception ex)
        {
            ConnectionErrorMessage =
                $"Couldn't connect to Outlook. Make sure the classic desktop Outlook app is installed and try again. ({ex.Message})";
            return;
        }
        finally
        {
            IsLoadingFolders = false;
        }

        // Each mailbox opens straight away so its folders are there without a click. Expanding
        // queues the read; the awaits below just keep the mailboxes filling in one after another
        // rather than all at once.
        foreach (var root in RootFolders)
        {
            root.IsExpanded = true;
            await root.LoadChildrenAsync();
        }
    }

    partial void OnSelectedFolderChanged(MailFolderNodeVm? value)
    {
        if (value is null || value.IsPlaceholder)
            return;

        _ = OpenFolderAsync(value, ++_openGeneration);
    }

    /// <summary>
    /// Puts a folder's board on screen. The board itself appears at once, with its columns; the mail
    /// follows as soon as Outlook hands it over. Shared mailboxes get the same board as your own
    /// folders - the state is keyed by store and folder, so each keeps its own columns.
    /// </summary>
    private async Task OpenFolderAsync(MailFolderNodeVm node, int generation)
    {
        var board = new KanbanBoardViewModel(
            node.Name,
            FolderKeyHelper.BuildKey(node.StoreId, node.EntryId),
            node.StoreId,
            node.EntryId,
            _boardStore,
            _outlook,
            _attachmentIndex,
            _indexer,
            _sentMail,
            _bodyLoader,
            new MailCardActions(ShowPreview, DeleteCard, SetRead, OnCardAnnotationsChanged));

        (CurrentContent as IDisposable)?.Dispose();
        CurrentContent = board;

        try
        {
            await board.LoadAsync();
        }
        catch (Exception ex) when (generation == _openGeneration)
        {
            ConnectionErrorMessage = $"Couldn't load folder '{node.Name}': {ex.Message}";
        }
        catch (Exception)
        {
            // The user moved on to another folder while this one was loading - its failure is no
            // longer anything they can act on.
        }
    }

    private void ApplyFolderOrder(ObservableCollection<MailFolderNodeVm> nodes, string parentKey)
    {
        var order = _boardStore.GetFolderOrder(parentKey);
        if (order is not { Count: > 0 })
            return;

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
        Preview = new PreviewPaneViewModel(
            summary,
            _outlook,
            _boardStore.Annotations,
            ClosePreview,
            DeleteMails,
            SetRead,
            OnPreviewAnnotationsChanged);

        if (PreviewColumnWidth.Value <= 0)
            PreviewColumnWidth = new GridLength(420);
    }

    /// <summary>
    /// A note or flag set in the reading pane has to reach the card behind it - including its place
    /// in the column, which priority pinning may have just changed.
    /// </summary>
    private void OnPreviewAnnotationsChanged(string entryId)
    {
        if (CurrentContent is KanbanBoardViewModel board)
            board.RefreshAnnotations(entryId);
    }

    /// <summary>And the other way round: a card's note reaches the reading pane showing that mail.</summary>
    private void OnCardAnnotationsChanged(MailCardViewModel card)
    {
        if (Preview is { } preview && card.Messages.Any(m => m.EntryId == preview.EntryId))
            preview.RefreshAnnotations();
    }

    private void DeleteCard(MailCardViewModel card) => DeleteMails(card.Messages);

    private void DeleteMails(IReadOnlyList<MailSummary> mails) => _ = DeleteMailsAsync(mails);

    /// <summary>
    /// Moves mails to Outlook's Deleted Items folder and takes them off the board right away, rather
    /// than waiting for the next sync to notice.
    /// </summary>
    private async Task DeleteMailsAsync(IReadOnlyList<MailSummary> mails)
    {
        var deleted = new List<string>();
        foreach (var mail in mails)
        {
            try
            {
                await _outlook.DeleteMailAsync(mail.StoreId, mail.EntryId);
                deleted.Add(mail.EntryId);
            }
            catch (Exception ex)
            {
                ConnectionErrorMessage = $"Couldn't delete '{mail.Subject}': {ex.Message}";
            }
        }

        if (deleted.Count == 0)
            return;

        if (CurrentContent is KanbanBoardViewModel board)
            board.DropMails(deleted);

        if (Preview is not null && deleted.Contains(Preview.EntryId, StringComparer.OrdinalIgnoreCase))
            ClosePreview();
    }

    private void SetRead(IReadOnlyList<MailSummary> mails, bool isRead) => _ = SetReadAsync(mails, isRead);

    /// <summary>
    /// Marks mails read or unread in Outlook and mirrors it onto the board right away. Takes the
    /// whole list, so a conversation tile carries its entire thread.
    /// </summary>
    private async Task SetReadAsync(IReadOnlyList<MailSummary> mails, bool isRead)
    {
        var changed = new List<string>();
        foreach (var mail in mails.Where(m => m.IsRead != isRead))
        {
            try
            {
                await _outlook.SetReadAsync(mail.StoreId, mail.EntryId, isRead);
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

        if (CurrentContent is KanbanBoardViewModel board)
            board.RefreshMailState(changed);

        if (Preview is not null && changed.Contains(Preview.EntryId, StringComparer.OrdinalIgnoreCase))
            Preview.RefreshReadState();
    }

    /// <summary>Re-lays out the open folder after a setting changed how its cards group or sort.</summary>
    private void RelayoutCurrentFolder()
    {
        if (CurrentContent is KanbanBoardViewModel board)
            board.RebuildCards();
    }

    [RelayCommand]
    private void ClosePreview()
    {
        Preview = null;
        PreviewColumnWidth = new GridLength(0);
    }

    /// <summary>Puts the error banner away. The next failure brings it back with its own message.</summary>
    [RelayCommand]
    private void DismissError() => ConnectionErrorMessage = null;

    public void Shutdown()
    {
        (CurrentContent as IDisposable)?.Dispose();

        // Written on the way out rather than on every pixel of a splitter drag.
        _boardStore.Settings.SidebarWidth = SidebarWidth.Value;
        _boardStore.Save();

        _attachmentIndex.Save();
        _outlook.Dispose();
    }
}
