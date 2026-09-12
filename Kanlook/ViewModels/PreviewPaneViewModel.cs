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

    public ObservableCollection<AttachmentViewModel> Attachments { get; } = [];

    [ObservableProperty]
    private string? _htmlBody;

    [ObservableProperty]
    private bool _isLoading = true;

    /// <summary>Set when opening an attachment failed, so the pane can say why.</summary>
    [ObservableProperty]
    private string? _attachmentError;

    private readonly Action _onClose;
    private readonly IOutlookService _outlook;
    private readonly MailSummary _summary;

    public PreviewPaneViewModel(MailSummary summary, IOutlookService outlook, Action onClose)
    {
        _outlook = outlook;
        _summary = summary;

        Subject = summary.Subject;
        SenderDisplay = string.IsNullOrEmpty(summary.SenderEmail)
            ? summary.SenderName
            : $"{summary.SenderName} <{summary.SenderEmail}>";
        ToNames = summary.ToNames;
        ReceivedTime = summary.ReceivedTime;
        _onClose = onClose;

        try
        {
            summary.HtmlBody ??= outlook.GetHtmlBody(summary.StoreId, summary.EntryId);
            HtmlBody = summary.HtmlBody ?? "<i>(no content)</i>";
        }
        catch (Exception ex)
        {
            HtmlBody = $"<i>Couldn't load message body: {ex.Message}</i>";
        }
        finally
        {
            IsLoading = false;
        }

        if (summary.HasAttachments)
        {
            try
            {
                foreach (var attachment in outlook.GetAttachments(summary.StoreId, summary.EntryId))
                    Attachments.Add(new AttachmentViewModel(attachment, OpenAttachment));
            }
            catch (Exception)
            {
                // Best-effort - a missing attachment list isn't worth failing the whole preview over.
            }
        }
    }

    private void OpenAttachment(AttachmentInfo attachment)
    {
        try
        {
            AttachmentError = null;
            _outlook.OpenAttachment(_summary.StoreId, _summary.EntryId, attachment.Index);
        }
        catch (Exception ex)
        {
            AttachmentError = $"Couldn't open '{attachment.FileName}': {ex.Message}";
        }
    }

    [RelayCommand]
    private void Close() => _onClose();

    [RelayCommand]
    private void Reply() => TryRespond(() => _outlook.Reply(_summary.StoreId, _summary.EntryId));

    [RelayCommand]
    private void ReplyAll() => TryRespond(() => _outlook.ReplyAll(_summary.StoreId, _summary.EntryId));

    [RelayCommand]
    private void Forward() => TryRespond(() => _outlook.Forward(_summary.StoreId, _summary.EntryId));

    private static void TryRespond(Action action)
    {
        try
        {
            action();
        }
        catch (Exception)
        {
            // Best-effort for the demo - e.g. the item was moved/deleted since it was loaded.
        }
    }
}
