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

    public MailCardViewModel(MailSummary summary, Action<MailSummary> onSelect)
        : this([summary], onSelect)
    {
    }

    public MailCardViewModel(IEnumerable<MailSummary> messages, Action<MailSummary> onSelect)
    {
        _messages = messages.OrderByDescending(m => m.ReceivedTime).ToList();
        if (_messages.Count == 0)
            throw new ArgumentException("A card needs at least one message.", nameof(messages));

        _onSelect = onSelect;
        RebuildProjections();
    }

    /// <summary>The newest message - the one the tile itself shows.</summary>
    public MailSummary Summary => _messages[0];

    public string ConversationKey => Summary.ConversationKey;

    public string Subject => Summary.Subject;
    public string SenderName => Summary.SenderName;
    public string Snippet => Summary.Snippet;
    public DateTime ReceivedTime => Summary.ReceivedTime;
    public bool IsRead => _messages.All(m => m.IsRead);
    public bool HasAttachments => _messages.Any(m => m.HasAttachments);
    public MailImportance Importance => _messages.Max(m => m.Importance);

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

        _messages.Sort((a, b) => b.ReceivedTime.CompareTo(a.ReceivedTime));
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
        OnPropertyChanged(nameof(IsRead));
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
    private void ToggleExpanded() => IsExpanded = !IsExpanded;
}
