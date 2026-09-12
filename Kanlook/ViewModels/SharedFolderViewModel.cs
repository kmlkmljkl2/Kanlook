using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using Kanlook.Models;

namespace Kanlook.ViewModels;

public sealed partial class SharedFolderViewModel : ObservableObject
{
    public string FolderName { get; }
    public ObservableCollection<MailCardViewModel> Cards { get; } = [];

    private readonly List<MailSummary> _mails;
    private readonly AppSettings _settings;
    private readonly Action<MailSummary> _onMailSelected;

    public SharedFolderViewModel(
        string folderName,
        List<MailSummary> mails,
        AppSettings settings,
        Action<MailSummary> onMailSelected)
    {
        FolderName = folderName;
        _mails = mails;
        _settings = settings;
        _onMailSelected = onMailSelected;
        RebuildCards();
    }

    public void RebuildCards()
    {
        Cards.Clear();

        if (_settings.GroupByConversation)
        {
            var conversations = _mails
                .GroupBy(m => m.ConversationKey)
                .OrderByDescending(g => g.Max(m => m.ReceivedTime));

            foreach (var conversation in conversations)
                Cards.Add(new MailCardViewModel(conversation, _onMailSelected));
        }
        else
        {
            foreach (var mail in _mails.OrderByDescending(m => m.ReceivedTime))
                Cards.Add(new MailCardViewModel(mail, _onMailSelected));
        }
    }
}
