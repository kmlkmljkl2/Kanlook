using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Kanlook.Models;
using Kanlook.Services;

namespace Kanlook.ViewModels;

public sealed partial class PreviewPaneViewModel : ObservableObject
{
    public string Subject { get; }
    public string SenderDisplay { get; }
    public string ToNames { get; }
    public DateTime ReceivedTime { get; }
    public string EntryId => _summary.EntryId;

    public ObservableCollection<AttachmentViewModel> Attachments { get; } = [];

    [ObservableProperty]
    private string? _htmlBody;

    [ObservableProperty]
    private bool _isLoading = true;

    /// <summary>Set when opening an attachment failed, so the pane can say why.</summary>
    [ObservableProperty]
    private string? _attachmentError;

    /// <summary>Set when Outlook refused to open a reply or forward window.</summary>
    [ObservableProperty]
    private string? _respondError;

    /// <summary>Swaps the note for an editable text box. The note itself only changes on commit.</summary>
    [ObservableProperty]
    private bool _isEditingComment;

    [ObservableProperty]
    private string _commentDraft = "";

    public bool IsRead => _summary.IsRead;

    /// <summary>Names the action, not the state - the button flips whatever the mail currently is.</summary>
    public string ReadToggleLabel => _summary.IsRead ? "Mark unread" : "Mark read";

    /// <summary>The user's own priority for this mail, independent of the sender's importance.</summary>
    public bool IsHighPriority => _annotations.IsHighPriority(EntryId);

    public string PriorityLabel => IsHighPriority ? "High priority" : "Set high priority";

    /// <summary>
    /// This mail's note. Unlike a card's, it is never another message's - the pane shows one mail, so
    /// the note it edits is that mail's own.
    /// </summary>
    public string Comment => _annotations.CommentFor(EntryId);

    public bool HasComment => Comment.Length > 0;

    public bool ShowsImportance => _summary.Importance != MailImportance.Normal;

    public MailImportance Importance => _summary.Importance;

    private readonly Action _onClose;
    private readonly Action<IReadOnlyList<MailSummary>> _onDelete;
    private readonly Action<IReadOnlyList<MailSummary>, bool> _onSetRead;

    /// <summary>Tells the board that this mail's note or flag changed, so its card catches up.</summary>
    private readonly Action<string> _onAnnotationsChanged;

    private readonly MailAnnotationStore _annotations;
    private readonly IOutlookService _outlook;
    private readonly MailSummary _summary;

    public PreviewPaneViewModel(
        MailSummary summary,
        IOutlookService outlook,
        MailAnnotationStore annotations,
        Action onClose,
        Action<IReadOnlyList<MailSummary>> onDelete,
        Action<IReadOnlyList<MailSummary>, bool> onSetRead,
        Action<string> onAnnotationsChanged)
    {
        _outlook = outlook;
        _summary = summary;
        _annotations = annotations;

        Subject = summary.Subject;
        SenderDisplay = string.IsNullOrEmpty(summary.SenderEmail)
            ? summary.SenderName
            : $"{summary.SenderName} <{summary.SenderEmail}>";
        ToNames = summary.ToNames;
        ReceivedTime = summary.ReceivedTime;
        _onClose = onClose;
        _onDelete = onDelete;
        _onSetRead = onSetRead;
        _onAnnotationsChanged = onAnnotationsChanged;

        // The pane's frame - subject, sender, the buttons - is on screen straight away; the body
        // and attachment list arrive when Outlook has them.
        _ = LoadAsync();
    }

    private async Task LoadAsync()
    {
        try
        {
            _summary.HtmlBody ??= await _outlook.GetHtmlBodyAsync(_summary.StoreId, _summary.EntryId);
            HtmlBody = _summary.HtmlBody ?? "<i>(no content)</i>";
        }
        catch (Exception ex)
        {
            HtmlBody = $"<i>Couldn't load message body: {ex.Message}</i>";
        }
        finally
        {
            IsLoading = false;
        }

        if (!_summary.HasAttachments)
            return;

        try
        {
            foreach (var attachment in await _outlook.GetAttachmentsAsync(_summary.StoreId, _summary.EntryId))
                Attachments.Add(new AttachmentViewModel(attachment, OpenAttachment));
        }
        catch (Exception)
        {
            // Best-effort - a missing attachment list isn't worth failing the whole preview over.
        }
    }

    private async void OpenAttachment(AttachmentInfo attachment)
    {
        try
        {
            AttachmentError = null;
            await _outlook.OpenAttachmentAsync(_summary.StoreId, _summary.EntryId, attachment.Index);
        }
        catch (Exception ex)
        {
            AttachmentError = $"Couldn't open '{attachment.FileName}': {ex.Message}";
        }
    }

    [RelayCommand]
    private void Close() => _onClose();

    [RelayCommand]
    private void Delete() => _onDelete([_summary]);

    [RelayCommand]
    private void ToggleRead() => _onSetRead([_summary], !_summary.IsRead);

    /// <summary>Restates the toggle after the mail's read state changed, here or on the board.</summary>
    public void RefreshReadState()
    {
        OnPropertyChanged(nameof(IsRead));
        OnPropertyChanged(nameof(ReadToggleLabel));
    }

    /// <summary>Re-reads the note and flag from the store - after an edit here, or on the card.</summary>
    public void RefreshAnnotations()
    {
        OnPropertyChanged(nameof(IsHighPriority));
        OnPropertyChanged(nameof(PriorityLabel));
        OnPropertyChanged(nameof(Comment));
        OnPropertyChanged(nameof(HasComment));
    }

    [RelayCommand]
    private void TogglePriority()
    {
        if (_annotations.SetHighPriority([EntryId], !IsHighPriority))
            Applied();
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
        if (_annotations.SetComment(EntryId, CommentDraft))
            Applied();

        IsEditingComment = false;
    }

    [RelayCommand]
    private void CancelEditComment() => IsEditingComment = false;

    [RelayCommand]
    private void RemoveComment()
    {
        IsEditingComment = false;
        if (_annotations.SetComment(EntryId, ""))
            Applied();
    }

    private void Applied()
    {
        RefreshAnnotations();
        _onAnnotationsChanged(EntryId);
    }

    [RelayCommand]
    private Task Reply() => TryRespondAsync(() => _outlook.ReplyAsync(_summary.StoreId, _summary.EntryId));

    [RelayCommand]
    private Task ReplyAll() => TryRespondAsync(() => _outlook.ReplyAllAsync(_summary.StoreId, _summary.EntryId));

    [RelayCommand]
    private Task Forward() => TryRespondAsync(() => _outlook.ForwardAsync(_summary.StoreId, _summary.EntryId));

    private async Task TryRespondAsync(Func<Task> action)
    {
        try
        {
            RespondError = null;
            await action();
        }
        catch (Exception ex)
        {
            // Usually the item was moved or deleted in Outlook since the pane was opened.
            RespondError = $"Couldn't open the message window: {ex.Message}";
        }
    }
}
