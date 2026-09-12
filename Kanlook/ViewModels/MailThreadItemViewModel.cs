using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Kanlook.Models;

namespace Kanlook.ViewModels;

/// <summary>An older message inside a conversation tile - either received, or a reply the user sent.</summary>
public sealed partial class MailThreadItemViewModel : ObservableObject
{
    public MailSummary Summary { get; }

    /// <summary>True for a message out of Sent Items, shown as history without being on the board.</summary>
    public bool IsSent { get; }

    /// <summary>For a reply, who it went to - the interesting half, since the sender is always the user.</summary>
    public string SenderName => IsSent
        ? string.IsNullOrWhiteSpace(Summary.ToNames) ? "You" : "To: " + Summary.ToNames
        : Summary.SenderName;

    public string Snippet => Summary.Snippet;

    /// <summary>Sent mail carries no meaningful received time, so fall back to when it was composed.</summary>
    public DateTime ReceivedTime => IsSent ? Summary.CreationTime : Summary.ReceivedTime;

    /// <summary>A reply the user wrote is never unread.</summary>
    public bool IsRead => IsSent || Summary.IsRead;

    public bool HasAttachments => Summary.HasAttachments;

    private readonly Action<MailSummary> _onSelect;

    public MailThreadItemViewModel(MailSummary summary, bool isSent, Action<MailSummary> onSelect)
    {
        Summary = summary;
        IsSent = isSent;
        _onSelect = onSelect;
    }

    [RelayCommand]
    private void Select() => _onSelect(Summary);
}
