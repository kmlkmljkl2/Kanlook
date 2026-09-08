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
    private static readonly TimeSpan PollInterval = TimeSpan.FromSeconds(25);
    private const int PollMaxCount = 30;

    private readonly BoardStateStore _boardStore;
    private readonly BoardState _board;
    private readonly Action<MailCardViewModel> _onCardSelected;
    private readonly IOutlookService _outlook;
    private readonly string _storeId;
    private readonly string _folderEntryId;
    private readonly HashSet<string> _knownEntryIds = [];
    private readonly DispatcherTimer _pollTimer;

    public string FolderName { get; }

    public ObservableCollection<KanbanColumnViewModel> Columns { get; } = [];

    public KanbanBoardViewModel(
        string folderName,
        string folderKey,
        string storeId,
        string folderEntryId,
        List<MailSummary> mails,
        BoardStateStore boardStore,
        IOutlookService outlook,
        Action<MailCardViewModel> onCardSelected)
    {
        FolderName = folderName;
        _boardStore = boardStore;
        _board = boardStore.GetOrCreateBoard(folderKey);
        _outlook = outlook;
        _storeId = storeId;
        _folderEntryId = folderEntryId;
        _onCardSelected = onCardSelected;

        foreach (var def in _board.Columns.OrderBy(c => c.Order))
            Columns.Add(CreateColumnVm(def));

        RefreshDefaultTargetFlags();

        var dirty = false;
        foreach (var mail in mails)
        {
            AssignNewMail(mail, ref dirty);
            _knownEntryIds.Add(mail.EntryId);
        }

        if (dirty)
            Save();

        _pollTimer = new DispatcherTimer { Interval = PollInterval };
        _pollTimer.Tick += (_, _) => PollForNewMail();
        _pollTimer.Start();
    }

    private void AssignNewMail(MailSummary mail, ref bool dirty)
    {
        if (!_board.CardAssignments.TryGetValue(mail.EntryId, out var columnId) ||
            Columns.All(c => c.Id != columnId))
        {
            columnId = DefaultColumnId();
            _board.CardAssignments[mail.EntryId] = columnId;
            dirty = true;
        }

        var column = Columns.First(c => c.Id == columnId);
        column.Cards.Add(new MailCardViewModel(mail, card => _onCardSelected(card)));
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

    private void PollForNewMail()
    {
        List<MailSummary> latest;
        try
        {
            latest = _outlook.GetMailSummaries(_storeId, _folderEntryId, PollMaxCount);
        }
        catch (Exception)
        {
            return; // best-effort - a transient Outlook hiccup shouldn't tear down the poll loop
        }

        var dirty = false;
        foreach (var mail in latest.Where(m => !_knownEntryIds.Contains(m.EntryId)))
        {
            var columnId = DefaultColumnId();
            _board.CardAssignments[mail.EntryId] = columnId;
            var column = Columns.First(c => c.Id == columnId);
            column.Cards.Add(new MailCardViewModel(mail, card => _onCardSelected(card)));
            _knownEntryIds.Add(mail.EntryId);
            dirty = true;
        }

        if (dirty)
            Save();
    }

    public void Dispose() => _pollTimer.Stop();

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
        {
            column.Cards.Remove(card);
            target.Cards.Add(card);
            _board.CardAssignments[card.Summary.EntryId] = target.Id;
        }

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
            sourceColumn.Cards.Remove(card);
            targetColumn.Cards.Insert(Math.Clamp(insertIndex, 0, targetColumn.Cards.Count), card);
            _board.CardAssignments[card.Summary.EntryId] = targetColumn.Id;
            _boardStore.Save();
        }
    }
}
