using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Kanlook.Models;

namespace Kanlook.ViewModels;

public sealed partial class MailCardViewModel : ObservableObject
{
    public MailSummary Summary { get; }

    public string Subject => Summary.Subject;
    public string SenderName => Summary.SenderName;
    public string Snippet => Summary.Snippet;
    public DateTime ReceivedTime => Summary.ReceivedTime;
    public bool IsRead => Summary.IsRead;
    public bool HasAttachments => Summary.HasAttachments;
    public MailImportance Importance => Summary.Importance;

    public string Initial => string.IsNullOrWhiteSpace(SenderName) ? "?" : SenderName.Trim()[..1].ToUpperInvariant();

    private readonly Action<MailCardViewModel> _onSelect;

    public MailCardViewModel(MailSummary summary, Action<MailCardViewModel> onSelect)
    {
        Summary = summary;
        _onSelect = onSelect;
    }

    [RelayCommand]
    private void Select() => _onSelect(this);
}
