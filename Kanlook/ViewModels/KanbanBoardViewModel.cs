using System.Collections.ObjectModel;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using GongSolutions.Wpf.DragDrop;
using Kanlook.Models;
using Kanlook.Services;

namespace Kanlook.ViewModels;

public sealed partial class KanbanBoardViewModel : ObservableObject, IDropTarget, IDragSource
{
    private readonly BoardStateStore _boardStore;
    private readonly BoardState _board;
    private readonly Action<MailCardViewModel> _onCardSelected;

    public string FolderName { get; }

    public ObservableCollection<KanbanColumnViewModel> Columns { get; } = [];

    public KanbanBoardViewModel(
        string folderName,
        string folderKey,
        List<MailSummary> mails,
        BoardStateStore boardStore,
        Action<MailCardViewModel> onCardSelected)
    {
        FolderName = folderName;
        _boardStore = boardStore;
        _board = boardStore.GetOrCreateBoard(folderKey);
        _onCardSelected = onCardSelected;

        foreach (var def in _board.Columns.OrderBy(c => c.Order))
            Columns.Add(CreateColumnVm(def));

        var dirty = false;
        foreach (var mail in mails)
        {
            if (!_board.CardAssignments.TryGetValue(mail.EntryId, out var columnId) ||
                Columns.All(c => c.Id != columnId))
            {
                columnId = Columns[0].Id;
                _board.CardAssignments[mail.EntryId] = columnId;
                dirty = true;
            }

            var column = Columns.First(c => c.Id == columnId);
            column.Cards.Add(new MailCardViewModel(mail, card => _onCardSelected(card)));
        }

        if (dirty)
            Save();
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
