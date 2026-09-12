using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Kanlook.Models;

namespace Kanlook.ViewModels;

/// <summary>
/// One board tile. Represents a single mail, or - when conversation grouping is on - a whole
/// conversation whose newest message is shown on the tile and whose older messages are revealed
/// by the expander.
/// </summary>
public sealed partial class MailCardViewModel : ObservableObject
{
    private readonly Action<MailSummary> _onSelect;
    private readonly Action<MailCardViewModel> _onDelete;
    private readonly Action<IReadOnlyList<MailSummary>, bool> _onSetRead;
    private readonly List<MailSummary> _messages;

    /// <summary>Every message on this tile, newest first.</summary>
    public IReadOnlyList<MailSummary> Messages => _messages;

    /// <summary>Everything but <see cref="Summary"/>, newest first - what the expander reveals.</summary>
    public ObservableCollection<MailThreadItemViewModel> OlderMessages { get; } = [];

    /// <summary>Every Outlook category on the tile, deduplicated across the conversation.</summary>
    public ObservableCollection<CategoryChip> Categories { get; } = [];

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ExpandLabel))]
    private bool _isExpanded;

    public MailCardViewModel(
        MailSummary summary,
        Action<MailSummary> onSelect,
        Action<MailCardViewModel> onDelete,
        Action<IReadOnlyList<MailSummary>, bool> onSetRead)
        : this([summary], onSelect, onDelete, onSetRead)
    {
    }

    public MailCardViewModel(
        IEnumerable<MailSummary> messages,
        Action<MailSummary> onSelect,
        Action<MailCardViewModel> onDelete,
        Action<IReadOnlyList<MailSummary>, bool> onSetRead)
    {
        _messages = [.. messages.OrderByDescending(m => m.CreationTime)];
        if (_messages.Count == 0)
            throw new ArgumentException("A card needs at least one message.", nameof(messages));

        _onSelect = onSelect;
        _onDelete = onDelete;
        _onSetRead = onSetRead;
        RebuildProjections();
    }

    /// <summary>The newest message - the one the tile itself shows.</summary>
    public MailSummary Summary => _messages[0];

    public string ConversationKey => Summary.ConversationKey;

    public string Subject => Summary.Subject;
    public string SenderName => Summary.SenderName;
    public string Snippet => Summary.Snippet;
    public DateTime ReceivedTime => Summary.ReceivedTime;

    /// <summary>What columns sort on - the newest creation time on the tile.</summary>
    public DateTime CreationTime => Summary.CreationTime;

    public bool IsRead => _messages.All(m => m.IsRead);
    public bool HasAttachments => _messages.Any(m => m.HasAttachments);
    public MailImportance Importance => _messages.Max(m => m.Importance);

    public string DeleteToolTip => IsConversation
        ? $"Move all {_messages.Count} messages to Deleted Items"
        : "Move to Deleted Items";

    /// <summary>Shows the action the button performs, not the current state.</summary>
    public string ReadToggleGlyph => IsRead ? "✉" : "✔";

    public string ReadToggleToolTip => (IsRead, IsConversation) switch
    {
        (true, true) => $"Mark all {_messages.Count} messages unread",
        (true, false) => "Mark unread",
        (false, true) => $"Mark all {_messages.Count} messages read",
        (false, false) => "Mark read",
    };

    public bool HasCategories => Categories.Count > 0;

    public int MessageCount => _messages.Count;
    public bool IsConversation => _messages.Count > 1;
    public string ExpandLabel => IsExpanded
        ? "Hide older messages"
        : $"Show {_messages.Count - 1} older message{(_messages.Count == 2 ? "" : "s")}";

    public string Initial => string.IsNullOrWhiteSpace(SenderName) ? "?" : SenderName.Trim()[..1].ToUpperInvariant();

    /// <summary>Folds another message of the same conversation into this tile.</summary>
    public void AddMessages(IEnumerable<MailSummary> messages)
    {
        var added = false;
        foreach (var message in messages.Where(m => _messages.All(existing => existing.EntryId != m.EntryId)))
        {
            _messages.Add(message);
            added = true;
        }

        if (!added)
            return;

        _messages.Sort((a, b) => b.CreationTime.CompareTo(a.CreationTime));
        RebuildProjections();
        NotifyProjectionChanged();
    }

    /// <summary>
    /// Drops messages that vanished from the folder in Outlook (deleted, or moved elsewhere - a move
    /// gives the mail a new entry id, so it reads as gone). Returns true when nothing would be left
    /// and the caller should drop the whole tile; in that case the tile is left untouched.
    /// </summary>
    public bool RemoveMessages(ICollection<string> entryIds)
    {
        if (_messages.All(m => entryIds.Contains(m.EntryId)))
            return true;

        if (_messages.RemoveAll(m => entryIds.Contains(m.EntryId)) == 0)
            return false;

        RebuildProjections();
        NotifyProjectionChanged();
        return false;
    }

    /// <summary>Re-reads the tile's messages after Outlook changed their categories or read state.</summary>
    public void RefreshState(ICollection<string> entryIds)
    {
        if (_messages.All(m => !entryIds.Contains(m.EntryId)))
            return;

        RebuildProjections();
        NotifyProjectionChanged();
    }

    private void RebuildProjections()
    {
        OlderMessages.Clear();
        foreach (var message in _messages.Skip(1))
            OlderMessages.Add(new MailThreadItemViewModel(message, _onSelect));

        Categories.Clear();
        foreach (var name in _messages.SelectMany(m => m.CategoryNames).Distinct(StringComparer.OrdinalIgnoreCase))
            Categories.Add(CategoryChip.For(name));
    }

    /// <summary>The tile projects the newest message, so any change to the set restates everything.</summary>
    private void NotifyProjectionChanged()
    {
        OnPropertyChanged(nameof(Summary));
        OnPropertyChanged(nameof(Subject));
        OnPropertyChanged(nameof(SenderName));
        OnPropertyChanged(nameof(Snippet));
        OnPropertyChanged(nameof(ReceivedTime));
        OnPropertyChanged(nameof(CreationTime));
        OnPropertyChanged(nameof(DeleteToolTip));
        OnPropertyChanged(nameof(IsRead));
        OnPropertyChanged(nameof(ReadToggleGlyph));
        OnPropertyChanged(nameof(ReadToggleToolTip));
        OnPropertyChanged(nameof(HasAttachments));
        OnPropertyChanged(nameof(Importance));
        OnPropertyChanged(nameof(Initial));
        OnPropertyChanged(nameof(HasCategories));
        OnPropertyChanged(nameof(MessageCount));
        OnPropertyChanged(nameof(IsConversation));
        OnPropertyChanged(nameof(ExpandLabel));
    }

    [RelayCommand]
    private void Select() => _onSelect(Summary);

    [RelayCommand]
    private void Delete() => _onDelete(this);

    /// <summary>
    /// Flips the tile's read state. A conversation tile is read only once every message on it is, so
    /// the toggle carries the whole thread with it rather than just the newest message.
    /// </summary>
    [RelayCommand]
    private void ToggleRead() => _onSetRead(_messages, !IsRead);

    [RelayCommand]
    private void ToggleExpanded() => IsExpanded = !IsExpanded;
}
