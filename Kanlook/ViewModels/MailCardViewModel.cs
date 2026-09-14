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
    private readonly MailCardContext _context;
    private readonly List<MailSummary> _messages;

    /// <summary>
    /// Replies the user sent in this conversation, newest first. Deliberately kept apart from
    /// <see cref="_messages"/>: they show up in the history but are not on the board, so they can't
    /// become <see cref="Summary"/> and never affect the column, sort order, read state or delete.
    /// </summary>
    private readonly List<MailSummary> _sent = [];

    /// <summary>Every message on this tile, newest first. Sent replies are not part of it.</summary>
    public IReadOnlyList<MailSummary> Messages => _messages;

    /// <summary>
    /// The tile's history, newest first - everything but <see cref="Summary"/>, with the user's own
    /// replies interleaved by time, which is what the expander reveals.
    /// </summary>
    public ObservableCollection<MailThreadItemViewModel> OlderMessages { get; } = [];

    /// <summary>Every Outlook category on the tile, deduplicated across the conversation.</summary>
    public ObservableCollection<CategoryChip> Categories { get; } = [];

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(ExpandLabel))]
    private bool _isExpanded;

    /// <summary>Swaps the note for an editable text box. The note itself only changes on commit.</summary>
    [ObservableProperty]
    private bool _isEditingComment;

    /// <summary>What the note editor holds. Written back to the store by <see cref="CommitComment"/>.</summary>
    [ObservableProperty]
    private string _commentDraft = "";

    public MailCardViewModel(MailSummary summary, MailCardContext context)
        : this([summary], context)
    {
    }

    public MailCardViewModel(IEnumerable<MailSummary> messages, MailCardContext context)
    {
        _messages = [.. messages.OrderByDescending(m => m.CreationTime)];
        if (_messages.Count == 0)
            throw new ArgumentException("A card needs at least one message.", nameof(messages));

        _context = context;
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

    // Delete and the read toggle only ever act on the board's own messages, so their labels count
    // those - a tile with one received mail and two replies must not offer to delete three.
    public string DeleteToolTip => _messages.Count > 1
        ? $"Move all {_messages.Count} messages to Deleted Items"
        : "Move to Deleted Items";

    /// <summary>Shows the action the button performs, not the current state.</summary>
    public string ReadToggleGlyph => IsRead ? "✉" : "✔";

    public string ReadToggleToolTip => (IsRead, Threaded: _messages.Count > 1) switch
    {
        (true, true) => $"Mark all {_messages.Count} messages unread",
        (true, false) => "Mark unread",
        (false, true) => $"Mark all {_messages.Count} messages read",
        (false, false) => "Mark read",
    };

    public bool HasCategories => Categories.Count > 0;

    /// <summary>
    /// The user's own priority, not the sender's. A conversation counts as high priority as soon as
    /// one of its messages is flagged - the same rule the attachment and importance badges follow.
    /// </summary>
    public bool IsHighPriority => _messages.Any(m => _context.Annotations.IsHighPriority(m.EntryId));

    public string PriorityToolTip => IsHighPriority
        ? "High priority - click to clear"
        : "Mark as high priority";

    /// <summary>
    /// Which message the note is filed under. A note belongs to a mail rather than to the tile, so a
    /// conversation keeps the note it was given even once a newer reply takes over the tile's face.
    /// </summary>
    private MailSummary CommentTarget =>
        _messages.FirstOrDefault(m => _context.Annotations.CommentFor(m.EntryId).Length > 0) ?? Summary;

    public string Comment => _context.Annotations.CommentFor(CommentTarget.EntryId);

    public bool HasComment => Comment.Length > 0;

    public string CommentToolTip => HasComment ? "Edit note" : "Add a note";

    /// <summary>The badge count, which - like Outlook's - includes the user's own replies.</summary>
    public int MessageCount => _messages.Count + _sent.Count;

    /// <summary>Messages the expander reveals.</summary>
    private int HistoryCount => _messages.Count - 1 + _sent.Count;

    public bool IsConversation => HistoryCount > 0;
    public string ExpandLabel => IsExpanded
        ? "Hide older messages"
        : $"Show {HistoryCount} older message{(HistoryCount == 1 ? "" : "s")}";

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

    /// <summary>
    /// Replaces the replies shown in the tile's history. Anything already on the board is skipped,
    /// so viewing the Sent Items folder itself doesn't list a mail twice.
    /// </summary>
    public void SetSentMessages(IEnumerable<MailSummary> sent)
    {
        var next = sent
            .Where(s => _messages.All(m => m.EntryId != s.EntryId))
            .OrderByDescending(s => s.CreationTime)
            .ToList();

        if (next.Count == _sent.Count && next.Zip(_sent).All(pair => pair.First.EntryId == pair.Second.EntryId))
            return;

        _sent.Clear();
        _sent.AddRange(next);
        RebuildProjections();
        NotifyProjectionChanged();
    }

    private void RebuildProjections()
    {
        // Received history and sent replies read as one thread, ordered by time like Outlook's.
        OlderMessages.Clear();
        var history = _messages.Skip(1).Select(m => (Message: m, IsSent: false))
            .Concat(_sent.Select(m => (Message: m, IsSent: true)))
            .OrderByDescending(entry => entry.Message.CreationTime);

        foreach (var entry in history)
            OlderMessages.Add(new MailThreadItemViewModel(entry.Message, entry.IsSent, _context.OnSelect));

        // Only the board's own mail contributes categories - a reply's are the user's own doing.
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
        NotifyAnnotationChanged();
    }

    /// <summary>Both annotations are read across the tile's messages, so a changed set restates them.</summary>
    private void NotifyAnnotationChanged()
    {
        OnPropertyChanged(nameof(IsHighPriority));
        OnPropertyChanged(nameof(PriorityToolTip));
        OnPropertyChanged(nameof(Comment));
        OnPropertyChanged(nameof(HasComment));
        OnPropertyChanged(nameof(CommentToolTip));
    }

    [RelayCommand]
    private void Select() => _context.OnSelect(Summary);

    [RelayCommand]
    private void Delete() => _context.OnDelete(this);

    /// <summary>
    /// Flips the user's priority flag. Like the read toggle it carries the whole thread, so a
    /// conversation is either flagged or not rather than partly both.
    /// </summary>
    [RelayCommand]
    private void TogglePriority()
    {
        if (!_context.Annotations.SetHighPriority(_messages.Select(m => m.EntryId), !IsHighPriority))
            return;

        NotifyAnnotationChanged();
        _context.OnPriorityChanged?.Invoke(this);
    }

    [RelayCommand]
    private void BeginEditComment()
    {
        CommentDraft = Comment;
        IsEditingComment = true;
    }

    [RelayCommand]
    private void CommitComment()
    {
        if (_context.Annotations.SetComment(CommentTarget.EntryId, CommentDraft))
            NotifyAnnotationChanged();

        IsEditingComment = false;
    }

    [RelayCommand]
    private void CancelEditComment() => IsEditingComment = false;

    /// <summary>Clears the note outright, so it doesn't take an edit-select-all-delete to get rid of.</summary>
    [RelayCommand]
    private void RemoveComment()
    {
        IsEditingComment = false;
        if (_context.Annotations.SetComment(CommentTarget.EntryId, ""))
            NotifyAnnotationChanged();
    }

    /// <summary>
    /// Flips the tile's read state. A conversation tile is read only once every message on it is, so
    /// the toggle carries the whole thread with it rather than just the newest message.
    /// </summary>
    [RelayCommand]
    private void ToggleRead() => _context.OnSetRead(_messages, !IsRead);

    [RelayCommand]
    private void ToggleExpanded() => IsExpanded = !IsExpanded;
}
