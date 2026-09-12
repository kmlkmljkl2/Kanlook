using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Kanlook.Models;

namespace Kanlook.ViewModels;

/// <summary>An older message inside a conversation tile.</summary>
public sealed partial class MailThreadItemViewModel : ObservableObject
{
    public MailSummary Summary { get; }

    public string SenderName => Summary.SenderName;
    public string Snippet => Summary.Snippet;
    public DateTime ReceivedTime => Summary.ReceivedTime;
    public bool IsRead => Summary.IsRead;
    public bool HasAttachments => Summary.HasAttachments;

    private readonly Action<MailSummary> _onSelect;

    public MailThreadItemViewModel(MailSummary summary, Action<MailSummary> onSelect)
    {
        Summary = summary;
        _onSelect = onSelect;
    }

    [RelayCommand]
    private void Select() => _onSelect(Summary);
}
