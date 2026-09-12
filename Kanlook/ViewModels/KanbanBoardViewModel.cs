using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using GongSolutions.Wpf.DragDrop;
using Kanlook.Models;
using Kanlook.Services;

namespace Kanlook.ViewModels;

public sealed partial class KanbanBoardViewModel : ObservableObject, IDropTarget, IDragSource, IDisposable
{
    private static readonly TimeSpan SyncInterval = TimeSpan.FromSeconds(10);

    /// <summary>
    /// How far back the sync looks. Must match the initial load's window, otherwise mails the board
    /// shows but the sync can't see would look like they'd been removed in Outlook.
    /// </summary>
    private const int SyncMaxCount = 300;

    private readonly BoardStateStore _boardStore;
    private readonly BoardState _board;
    private readonly Action<MailSummary> _onMailSelected;
    private readonly Action<MailCardViewModel> _onCardDelete;
    private readonly Action<IReadOnlyList<MailSummary>, bool> _onSetRead;
    private readonly MailSearch _search;
    private readonly AttachmentIndexer _indexer;
    private readonly IOutlookService _outlook;
    private readonly string _storeId;
    private readonly string _folderEntryId;
    private readonly HashSet<string> _knownEntryIds = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<MailSummary> _mails = [];
    private readonly DispatcherTimer _syncTimer;

    public string FolderName { get; }

    public ObservableCollection<KanbanColumnViewModel> Columns { get; } = [];

    /// <summary>Filters the board as you type. Empty shows everything.</summary>
    [ObservableProperty]
    private string _searchText = "";

    [ObservableProperty]
    private string? _indexingStatus;

    private bool GroupByConversation => _boardStore.Settings.GroupByConversation;

    public KanbanBoardViewModel(
        string folderName,
        string folderKey,
        string storeId,
        string folderEntryId,
        List<MailSummary> mails,
        BoardStateStore boardStore,
        IOutlookService outlook,
        AttachmentIndex attachmentIndex,
        AttachmentIndexer indexer,
        Action<MailSummary> onMailSelected,
        Action<MailCardViewModel> onCardDelete,
        Action<IReadOnlyList<MailSummary>, bool> onSetRead)
    {
        FolderName = folderName;
        _boardStore = boardStore;
        _board = boardStore.GetOrCreateBoard(folderKey);
        _outlook = outlook;
        _storeId = storeId;
        _folderEntryId = folderEntryId;
        _onMailSelected = onMailSelected;
        _onCardDelete = onCardDelete;
        _onSetRead = onSetRead;
        _search = new MailSearch(attachmentIndex);
        _indexer = indexer;

        foreach (var def in _board.Columns.OrderBy(c => c.Order))
            Columns.Add(CreateColumnVm(def));

        RefreshDefaultTargetFlags();

        var dirty = false;
        foreach (var mail in mails)
        {
            EnsureAssigned(mail, ref dirty);
            _mails.Add(mail);
            _knownEntryIds.Add(mail.EntryId);
        }

        if (dirty)
            Save();

        RebuildCards();

        _indexer.Progressed += OnIndexerProgressed;
        _indexer.Enqueue(_mails);
        RefreshIndexingStatus();

        _syncTimer = new DispatcherTimer { Interval = SyncInterval };
        _syncTimer.Tick += (_, _) => SyncWithOutlook();
        _syncTimer.Start();
    }

    partial void OnSearchTextChanged(string value) => RebuildCards();

    [RelayCommand]
    private void ClearSearch() => SearchText = "";

    private void OnIndexerProgressed()
    {
        RefreshIndexingStatus();

        // Newly indexed attachment text can change what an active search matches.
        if (SearchText.Length > 0)
            RebuildCards();
    }

    private void RefreshIndexingStatus() =>
        IndexingStatus = _indexer.PendingCount > 0
            ? $"Indexing attachments… {_indexer.PendingCount} left"
            : null;

    private void EnsureAssigned(MailSummary mail, ref bool dirty)
    {
        if (_board.CardAssignments.TryGetValue(mail.EntryId, out var columnId) &&
            Columns.Any(c => c.Id == columnId))
            return;

        _board.CardAssignments[mail.EntryId] = DefaultColumnId();
        dirty = true;
    }

    private string ColumnIdOf(MailSummary mail) =>
        _board.CardAssignments.TryGetValue(mail.EntryId, out var id) ? id : DefaultColumnId();

    /// <summary>
    /// Rebuilds every column's tiles from the mails we know about, newest creation time first, and
    /// filtered by the search box. Grouping happens inside a column, so mails of one conversation
    /// that the user pulled apart by hand stay apart.
    /// </summary>
    public void RebuildCards()
    {
        var terms = MailSearch.ParseTerms(SearchText);

        foreach (var column in Columns)
        {
            column.Cards.Clear();

            var mails = _mails.Where(m => ColumnIdOf(m) == column.Id);
            if (GroupByConversation)
            {
                // A conversation is kept whole when any of its messages matches.
                var conversations = mails
                    .GroupBy(m => m.ConversationKey)
                    .Where(g => g.Any(m => _search.Matches(m, terms)))
                    .OrderByDescending(g => g.Max(m => m.CreationTime));

                foreach (var conversation in conversations)
                    column.Cards.Add(CreateCard(conversation));
            }
            else
            {
                var matching = mails
                    .Where(m => _search.Matches(m, terms))
                    .OrderByDescending(m => m.CreationTime);

                foreach (var mail in matching)
                    column.Cards.Add(CreateCard([mail]));
            }
        }
    }

    private MailCardViewModel CreateCard(IEnumerable<MailSummary> messages) =>
        new(messages, _onMailSelected, _onCardDelete, _onSetRead);

    private string DefaultColumnId() =>
        _board.DefaultColumnId is { } id && Columns.Any(c => c.Id == id) ? id : Columns[0].Id;

    public void SetDefaultColumn(KanbanColumnViewModel column)
    {
        _board.DefaultColumnId = column.Id;
        RefreshDefaultTargetFlags();
        Save();
    }

    private void RefreshDefaultTargetFlags()
    {
        var defaultId = DefaultColumnId();
        foreach (var c in Columns)
            c.IsDefaultTarget = c.Id == defaultId;
    }

    /// <summary>
    /// Reconciles the board against the folder as Outlook currently has it: mails deleted or moved
    /// away disappear, category and read-state edits come across, and new arrivals land on top.
    /// Outlook's own ItemRemove event doesn't say which item went, so comparing snapshots is both
    /// simpler and more reliable than event plumbing - and it can't miss a change.
    /// </summary>
    private void SyncWithOutlook()
    {
        List<MailItemState> state;
        try
        {
            state = _outlook.GetFolderState(_storeId, _folderEntryId, SyncMaxCount);
        }
        catch (Exception)
        {
            return; // best-effort - a transient Outlook hiccup shouldn't tear down the sync loop
        }

        ApplyRemovals(state); // persists itself
        ApplyStateChanges(state);

        if (ApplyArrivals(state))
            Save();
    }

    /// <summary>Drops mails that are no longer in the folder - deleted, or moved somewhere else.</summary>
    private void ApplyRemovals(List<MailItemState> state)
    {
        var present = state.Select(s => s.EntryId).ToHashSet(StringComparer.OrdinalIgnoreCase);

        // The snapshot only covers the folder's newest SyncMaxCount items. If it came back full,
        // anything older than its oldest row is simply out of view and must not be read as removed.
        // Non-mail rows count towards the window, which is why they're in the snapshot at all.
        var oldestObserved = state.Count >= SyncMaxCount
            ? state.Min(s => s.ReceivedTime)
            : DateTime.MinValue;

        var gone = _mails
            .Where(m => !present.Contains(m.EntryId) && m.ReceivedTime >= oldestObserved)
            .Select(m => m.EntryId)
            .ToList();

        DropMails(gone);
    }

    /// <summary>
    /// Forgets mails entirely - used both when Outlook no longer has them in this folder and when
    /// the user deletes them from the board.
    /// </summary>
    public void DropMails(IEnumerable<string> entryIds)
    {
        var ids = entryIds.ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (ids.Count == 0)
            return;

        foreach (var column in Columns)
        {
            foreach (var card in column.Cards.ToList())
            {
                if (card.RemoveMessages(ids))
                    column.Cards.Remove(card);
            }
        }

        _mails.RemoveAll(m => ids.Contains(m.EntryId));
        foreach (var entryId in ids)
        {
            _knownEntryIds.Remove(entryId);
            _board.CardAssignments.Remove(entryId);
        }

        Save();
    }

    /// <summary>Mirrors category and read-state edits made in Outlook onto the cards.</summary>
    private void ApplyStateChanges(List<MailItemState> state)
    {
        var byId = _mails.ToDictionary(m => m.EntryId, StringComparer.OrdinalIgnoreCase);
        var touched = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var item in state)
        {
            if (!byId.TryGetValue(item.EntryId, out var mail) ||
                (mail.IsRead == item.IsRead && mail.Categories == item.Categories))
                continue;

            mail.IsRead = item.IsRead;
            mail.Categories = item.Categories;
            touched.Add(mail.EntryId);
        }

        RefreshMailState(touched);
    }

    /// <summary>
    /// Restates the tiles holding these mails. Used both by the sync and after the board itself
    /// changed read state, so a card repaints without waiting for the next poll.
    /// </summary>
    public void RefreshMailState(IEnumerable<string> entryIds)
    {
        var ids = entryIds.ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (ids.Count == 0)
            return;

        foreach (var card in Columns.SelectMany(c => c.Cards))
            card.RefreshState(ids);
    }

    private bool ApplyArrivals(List<MailItemState> state)
    {
        var added = false;

        var arrivals = state
            .Where(s => s.IsMail && !_knownEntryIds.Contains(s.EntryId))
            .OrderBy(s => s.ReceivedTime);

        foreach (var arrival in arrivals)
        {
            MailSummary? mail;
            try
            {
                mail = _outlook.GetMailSummary(_storeId, arrival.EntryId);
            }
            catch (Exception)
            {
                continue; // moved or deleted again between the snapshot and now
            }

            if (mail is null)
                continue;

            var columnId = ColumnIdForArrival(mail);
            _board.CardAssignments[mail.EntryId] = columnId;
            _mails.Add(mail);
            _knownEntryIds.Add(mail.EntryId);
            Insert(Columns.First(c => c.Id == columnId), mail);
            added = true;
        }

        if (added)
        {
            _indexer.Enqueue(_mails);
            RefreshIndexingStatus();
        }

        return added;
    }

    /// <summary>
    /// Which column a new arrival belongs in. With grouping on it follows its conversation to
    /// wherever the user filed it - even another column - the way Outlook keeps a thread together
    /// instead of starting a second tile back in the default column.
    /// </summary>
    private string ColumnIdForArrival(MailSummary mail)
    {
        if (!GroupByConversation)
            return DefaultColumnId();

        // Resolved from the mails rather than the tiles, because an active search may be hiding the
        // conversation's tile while its messages are still on the board.
        var sibling = _mails
            .Where(m => m.ConversationKey == mail.ConversationKey)
            .MaxBy(m => m.CreationTime);

        if (sibling is null)
            return DefaultColumnId();

        // A hand-split conversation can straddle columns; follow its newest message.
        var columnId = ColumnIdOf(sibling);
        return Columns.Any(c => c.Id == columnId) ? columnId : DefaultColumnId();
    }

    /// <summary>
    /// Files a newly arrived mail into its column by creation time, folding it into its conversation
    /// tile if one is already there. Being the newest mail, it lifts that tile to the top.
    /// </summary>
    private void Insert(KanbanColumnViewModel column, MailSummary mail)
    {
        if (GroupByConversation &&
            column.Cards.FirstOrDefault(c => c.ConversationKey == mail.ConversationKey) is { } existing)
        {
            // A shown conversation is kept whole - the same rule RebuildCards applies - so the new
            // message joins it even when the search text only matches its siblings.
            existing.AddMessages([mail]);
            MoveToSortedPosition(column, existing);
            return;
        }

        if (!_search.Matches(mail, MailSearch.ParseTerms(SearchText)))
            return; // filtered out by the active search - RebuildCards will pick it up when cleared

        var card = CreateCard([mail]);
        column.Cards.Insert(SortedIndexFor(column, card.CreationTime), card);
    }

    private static void MoveToSortedPosition(KanbanColumnViewModel column, MailCardViewModel card)
    {
        var from = column.Cards.IndexOf(card);
        var to = SortedIndexFor(column, card.CreationTime, ignoring: card);
        if (from >= 0 && from != to)
            column.Cards.Move(from, Math.Clamp(to, 0, column.Cards.Count - 1));
    }

    /// <summary>Where a tile of this age belongs in a column ordered newest creation time first.</summary>
    private static int SortedIndexFor(KanbanColumnViewModel column, DateTime creationTime, MailCardViewModel? ignoring = null)
    {
        var index = 0;
        foreach (var card in column.Cards)
        {
            if (ReferenceEquals(card, ignoring))
                continue;
            if (card.CreationTime <= creationTime)
                break;
            index++;
        }

        return index;
    }

    public void Dispose()
    {
        _syncTimer.Stop();
        _indexer.Progressed -= OnIndexerProgressed;
    }

    private KanbanColumnViewModel CreateColumnVm(KanbanColumnDefinition def) =>
        new(def, RemoveColumn, MoveColumn, Save, this);

    [RelayCommand]
    private void AddColumn()
    {
        var def = new KanbanColumnDefinition
        {
            Order = Columns.Count == 0 ? 0 : Columns.Max(c => c.Order) + 1,
        };
        _board.Columns.Add(def);
        var vm = CreateColumnVm(def);
        Columns.Add(vm);
        vm.IsEditingName = true;
        Save();
    }

    private void RemoveColumn(KanbanColumnViewModel column)
    {
        if (Columns.Count <= 1)
            return;

        var target = Columns.Where(c => c != column).OrderBy(c => Math.Abs(c.Order - column.Order)).First();
        foreach (var card in column.Cards.ToList())
            MoveCard(card, column, target, target.Cards.Count);

        Columns.Remove(column);
        _board.Columns.RemoveAll(c => c.Id == column.Id);

        var order = 0;
        foreach (var c in Columns.OrderBy(c => c.Order))
            c.Order = order++;

        RefreshDefaultTargetFlags();
        Save();
    }

    private void MoveColumn(KanbanColumnViewModel column, int direction)
    {
        var index = Columns.IndexOf(column);
        var newIndex = index + direction;
        if (newIndex < 0 || newIndex >= Columns.Count)
            return;

        (Columns[index].Order, Columns[newIndex].Order) = (Columns[newIndex].Order, Columns[index].Order);
        Columns.Move(index, newIndex);
        Save();
    }

    private void Save()
    {
        _board.Columns.Clear();
        _board.Columns.AddRange(Columns.Select(c => new KanbanColumnDefinition
        {
            Id = c.Id,
            Name = c.Name,
            ColorHex = c.ColorHex,
            Order = c.Order,
        }));
        _boardStore.Save();
    }

    public bool CanStartDrag(IDragInfo dragInfo) => dragInfo.SourceItem is MailCardViewModel;

    public void StartDrag(IDragInfo dragInfo)
    {
        dragInfo.Data = dragInfo.SourceItem;
        dragInfo.Effects = dragInfo.SourceItem is MailCardViewModel ? DragDropEffects.Move : DragDropEffects.None;
    }

    public void Dropped(IDropInfo dropInfo) { }

    public void DragDropOperationFinished(DragDropEffects operationResult, IDragInfo dragInfo) { }

    public void DragCancelled() { }

    public bool TryCatchOccurredException(Exception exception) => false;

    public void DragOver(IDropInfo dropInfo)
    {
        if (dropInfo.Data is MailCardViewModel)
        {
            dropInfo.Effects = DragDropEffects.Move;
            dropInfo.DropTargetAdorner = dropInfo.TargetItem is null
                ? DropTargetAdorners.Highlight
                : DropTargetAdorners.Insert;
        }
        else
        {
            dropInfo.Effects = DragDropEffects.None;
        }
    }

    public void Drop(IDropInfo dropInfo)
    {
        if (dropInfo.Data is not MailCardViewModel card)
            return;

        if (dropInfo.TargetCollection is not ObservableCollection<MailCardViewModel> targetCards)
            return;

        var targetColumn = Columns.FirstOrDefault(c => c.Cards == targetCards);
        var sourceColumn = Columns.FirstOrDefault(c => c.Cards.Contains(card));
        if (targetColumn is null || sourceColumn is null)
            return;

        var insertIndex = Math.Clamp(dropInfo.InsertIndex, 0, targetCards.Count);

        if (sourceColumn == targetColumn)
        {
            var oldIndex = sourceColumn.Cards.IndexOf(card);
            if (oldIndex < insertIndex)
                insertIndex--;
            if (oldIndex == insertIndex)
                return;
            sourceColumn.Cards.Move(oldIndex, Math.Clamp(insertIndex, 0, sourceColumn.Cards.Count - 1));
            _boardStore.Save();
        }
        else
        {
            MoveCard(card, sourceColumn, targetColumn, insertIndex);
            _boardStore.Save();
        }
    }

    /// <summary>
    /// Moves a tile between columns, reassigning every mail on it. With grouping on, dropping onto a
    /// column that already shows the same conversation merges the two tiles instead of duplicating it.
    /// </summary>
    private void MoveCard(
        MailCardViewModel card,
        KanbanColumnViewModel source,
        KanbanColumnViewModel target,
        int insertIndex)
    {
        source.Cards.Remove(card);

        if (GroupByConversation &&
            target.Cards.FirstOrDefault(c => c.ConversationKey == card.ConversationKey) is { } existing)
            existing.AddMessages(card.Messages);
        else
            target.Cards.Insert(Math.Clamp(insertIndex, 0, target.Cards.Count), card);

        foreach (var message in card.Messages)
            _board.CardAssignments[message.EntryId] = target.Id;
    }
}
