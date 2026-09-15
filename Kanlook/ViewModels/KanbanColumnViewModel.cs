using System.Collections.ObjectModel;
using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Kanlook.Models;

namespace Kanlook.ViewModels;

public sealed partial class KanbanColumnViewModel : ObservableObject
{
    public string Id { get; }
    public int Order { get; set; }
    public KanbanBoardViewModel Board { get; }

    [ObservableProperty]
    private string _name;

    [ObservableProperty]
    private string _colorHex;

    [ObservableProperty]
    private bool _isEditingName;

    [ObservableProperty]
    private bool _isColorPickerOpen;

    [ObservableProperty]
    private bool _isDefaultTarget;

    [ObservableProperty]
    private bool _isWaitSettingsOpen;

    /// <summary>The header's overflow menu - rename, colour, move, remove and the rest.</summary>
    [ObservableProperty]
    private bool _isMenuOpen;

    /// <summary>
    /// Mail parked here is waiting on somebody else. A reply to one of its conversations sends the
    /// whole conversation back to the default column - it needs attention again.
    /// </summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasReturnTimer))]
    [NotifyPropertyChangedFor(nameof(WaitingToolTip))]
    private bool _waitsForReply;

    /// <summary>Days until a waiting conversation returns even without a reply. Null turns it off.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ReturnAfterDaysLabel))]
    [NotifyPropertyChangedFor(nameof(HasReturnTimer))]
    [NotifyPropertyChangedFor(nameof(WaitingToolTip))]
    private int? _returnAfterDays;

    /// <summary>Time of day the return falls due, "HH:mm". Null keeps the time the mail was parked.</summary>
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(WaitingToolTip))]
    private string? _returnAtTime;

    /// <summary>What the settings popup edits. Only applied - and corrected - on commit.</summary>
    [ObservableProperty]
    private string _returnAfterDaysText = "";

    [ObservableProperty]
    private string _returnAtTimeText = "";

    /// <summary>The colours the picker offers.</summary>
    public IReadOnlyList<string> Palette => ColumnColors.All;

    /// <summary>The return deadline, short enough to sit in the header next to the clock.</summary>
    public string ReturnAfterDaysLabel => ReturnAfterDays is { } days ? $"{days}d" : "";

    public bool HasReturnTimer => WaitsForReply && ReturnAfterDays is not null;

    public string WaitingToolTip => !WaitsForReply
        ? "Waiting for reply"
        : ReturnAfterDays is { } days
            ? $"Waiting for reply - returns to the default column after {days} day{(days == 1 ? "" : "s")}" +
              (ReturnAtTime is null ? "" : $" at {ReturnAtTime}")
            : "Waiting for reply - returns to the default column when an answer arrives";

    public ObservableCollection<MailCardViewModel> Cards { get; } = [];

    private readonly Action<KanbanColumnViewModel> _onRemove;
    private readonly Action<KanbanColumnViewModel, int> _onMove;
    private readonly Action _onChanged;

    public KanbanColumnViewModel(
        KanbanColumnDefinition definition,
        Action<KanbanColumnViewModel> onRemove,
        Action<KanbanColumnViewModel, int> onMove,
        Action onChanged,
        KanbanBoardViewModel board)
    {
        Id = definition.Id;
        Order = definition.Order;
        _name = definition.Name;
        _colorHex = definition.ColorHex;

        // Straight onto the backing fields: going through the properties would fire the change
        // hooks below and have the board re-save the state it is still loading.
        _waitsForReply = definition.WaitsForReply;
        _returnAfterDays = definition.ReturnAfterDays;
        _returnAtTime = definition.ReturnAtTime;
        _returnAfterDaysText = definition.ReturnAfterDays?.ToString(CultureInfo.InvariantCulture) ?? "";
        _returnAtTimeText = definition.ReturnAtTime ?? "";

        _onRemove = onRemove;
        _onMove = onMove;
        _onChanged = onChanged;
        Board = board;
    }

    [RelayCommand]
    private void ToggleMenu() => IsMenuOpen = !IsMenuOpen;

    [RelayCommand]
    private void BeginRename()
    {
        IsMenuOpen = false;
        IsEditingName = true;
    }

    [RelayCommand]
    private void CommitRename()
    {
        if (string.IsNullOrWhiteSpace(Name))
            Name = "Untitled";
        IsEditingName = false;
        _onChanged();
    }

    /// <summary>
    /// Opens the colour picker. Reached from the overflow menu, so the menu steps aside first -
    /// two popups stacked on one button would fight over the click that dismisses them.
    /// </summary>
    [RelayCommand]
    private void OpenColorPicker()
    {
        IsMenuOpen = false;
        IsColorPickerOpen = true;
    }

    [RelayCommand]
    private void SetColor(string hex)
    {
        ColorHex = hex;
        IsColorPickerOpen = false;
        _onChanged();
    }

    [RelayCommand]
    private void SetAsDefault()
    {
        IsMenuOpen = false;
        Board.SetDefaultColumn(this);
    }

    [RelayCommand]
    private void OpenWaitSettings()
    {
        IsMenuOpen = false;
        IsWaitSettingsOpen = true;
    }

    /// <summary>
    /// Turning the flag on has to reach the board: the mail already sitting here needs a parked-at
    /// timestamp before any return timer can mean anything.
    /// </summary>
    partial void OnWaitsForReplyChanged(bool value)
    {
        _onChanged();
        Board.OnWaitingSettingsChanged(this);
    }

    /// <summary>Reads the popup's text back into the real settings, correcting what it can't parse.</summary>
    [RelayCommand]
    private void CommitWaitSettings()
    {
        ReturnAfterDays = int.TryParse(ReturnAfterDaysText, NumberStyles.Integer, CultureInfo.InvariantCulture, out var days) && days >= 0
            ? Math.Min(days, 365)
            : null;

        ReturnAtTime = ParseTimeOfDay(ReturnAtTimeText) is { } time ? time.ToString(@"hh\:mm") : null;

        // Echo the understood values back, so a typo doesn't sit there looking accepted.
        ReturnAfterDaysText = ReturnAfterDays?.ToString(CultureInfo.InvariantCulture) ?? "";
        ReturnAtTimeText = ReturnAtTime ?? "";

        IsWaitSettingsOpen = false;
        _onChanged();
        Board.OnWaitingSettingsChanged(this);
    }

    /// <summary>
    /// When a mail parked at <paramref name="since"/> is due back in the default column, or null when
    /// this column has no return timer. With no time of day set, "2 days" means exactly 48 hours.
    /// </summary>
    public DateTime? ReturnDueAt(DateTime since)
    {
        if (!WaitsForReply || ReturnAfterDays is not { } days)
            return null;

        var due = since.AddDays(days);
        return ParseTimeOfDay(ReturnAtTime) is { } time ? due.Date + time : due;
    }

    private static TimeSpan? ParseTimeOfDay(string? value) =>
        TimeSpan.TryParseExact(value, [@"hh\:mm", @"h\:mm"], CultureInfo.InvariantCulture, out var parsed)
            ? parsed
            : null;

    [RelayCommand]
    private void Remove()
    {
        IsMenuOpen = false;
        _onRemove(this);
    }

    [RelayCommand]
    private void MoveLeft()
    {
        IsMenuOpen = false;
        _onMove(this, -1);
    }

    [RelayCommand]
    private void MoveRight()
    {
        IsMenuOpen = false;
        _onMove(this, 1);
    }
}
