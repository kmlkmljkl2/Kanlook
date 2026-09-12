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
    private readonly IOutlookService _outlook;
    private readonly string _storeId;
    private readonly string _folderEntryId;
    private readonly HashSet<string> _knownEntryIds = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<MailSummary> _mails = [];
    private readonly DispatcherTimer _syncTimer;

    public string FolderName { get; }

    public ObservableCollection<KanbanColumnViewModel> Columns { get; } = [];

    private bool GroupByConversation => _boardStore.Settings.GroupByConversation;

    public KanbanBoardViewModel(
        string folderName,
        string folderKey,
        string storeId,
        string folderEntryId,
        List<MailSummary> mails,
        BoardStateStore boardStore,
        IOutlookService outlook,
        Action<MailSummary> onMailSelected)
    {
        FolderName = folderName;
        _boardStore = boardStore;
        _board = boardStore.GetOrCreateBoard(folderKey);
        _outlook = outlook;
        _storeId = storeId;
        _folderEntryId = folderEntryId;
        _onMailSelected = onMailSelected;

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

        _syncTimer = new DispatcherTimer { Interval = SyncInterval };
        _syncTimer.Tick += (_, _) => SyncWithOutlook();
        _syncTimer.Start();
    }

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
    /// Rebuilds every column's tiles from the mails we know about, newest first. Grouping happens
    /// inside a column, so mails of one conversation that the user pulled apart by hand stay apart.
    /// </summary>
    public void RebuildCards()
    {
        foreach (var column in Columns)
        {
            column.Cards.Clear();

            var mails = _mails.Where(m => ColumnIdOf(m) == column.Id);
            if (GroupByConversation)
            {
                var conversations = mails
                    .GroupBy(m => m.ConversationKey)
                    .OrderByDescending(g => g.Max(m => m.ReceivedTime));

                foreach (var conversation in conversations)
                    column.Cards.Add(new MailCardViewModel(conversation, _onMailSelected));
            }
            else
            {
                foreach (var mail in mails.OrderByDescending(m => m.ReceivedTime))
                    column.Cards.Add(new MailCardViewModel(mail, _onMailSelected));
            }
        }
    }

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

        var dirty = ApplyRemovals(state);
        ApplyStateChanges(state);
        dirty |= ApplyArrivals(state);

        if (dirty)
            Save();
    }

    /// <summary>Drops mails that are no longer in the folder - deleted, or moved somewhere else.</summary>
    private bool ApplyRemovals(List<MailItemState> state)
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
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        if (gone.Count == 0)
            return false;

        foreach (var column in Columns)
        {
            foreach (var card in column.Cards.ToList())
            {
                if (card.RemoveMessages(gone))
                    column.Cards.Remove(card);
            }
        }

        _mails.RemoveAll(m => gone.Contains(m.EntryId));
        foreach (var entryId in gone)
        {
            _knownEntryIds.Remove(entryId);
            _board.CardAssignments.Remove(entryId);
        }

        return true;
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

        if (touched.Count == 0)
            return;

        foreach (var card in Columns.SelectMany(c => c.Cards))
            card.RefreshState(touched);
    }

    private bool ApplyArrivals(List<MailItemState> state)
    {
        var added = false;

        // Oldest first, so inserting each at the top leaves the newest on top.
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

            var columnId = DefaultColumnId();
            _board.CardAssignments[mail.EntryId] = columnId;
            _mails.Add(mail);
            _knownEntryIds.Add(mail.EntryId);
            AddToTop(Columns.First(c => c.Id == columnId), mail);
            added = true;
        }

        return added;
    }

    /// <summary>Places a newly arrived mail at the top of its column, folding it into its conversation tile if one is already there.</summary>
    private void AddToTop(KanbanColumnViewModel column, MailSummary mail)
    {
        if (GroupByConversation &&
            column.Cards.FirstOrDefault(c => c.ConversationKey == mail.ConversationKey) is { } existing)
        {
            existing.AddMessages([mail]);
            var index = column.Cards.IndexOf(existing);
            if (index > 0)
                column.Cards.Move(index, 0);
            return;
        }

        column.Cards.Insert(0, new MailCardViewModel(mail, _onMailSelected));
    }

    public void Dispose() => _syncTimer.Stop();

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
