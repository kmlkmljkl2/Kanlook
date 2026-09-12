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
    private readonly Action<MailCardViewModel> _onCardDelete;
    private readonly Action<IReadOnlyList<MailSummary>, bool> _onSetRead;

    public SharedFolderViewModel(
        string folderName,
        List<MailSummary> mails,
        AppSettings settings,
        Action<MailSummary> onMailSelected,
        Action<MailCardViewModel> onCardDelete,
        Action<IReadOnlyList<MailSummary>, bool> onSetRead)
    {
        FolderName = folderName;
        _mails = mails;
        _settings = settings;
        _onMailSelected = onMailSelected;
        _onCardDelete = onCardDelete;
        _onSetRead = onSetRead;
        RebuildCards();
    }

    public void RebuildCards()
    {
        Cards.Clear();

        if (_settings.GroupByConversation)
        {
            var conversations = _mails
                .GroupBy(m => m.ConversationKey)
                .OrderByDescending(g => g.Max(m => m.CreationTime));

            foreach (var conversation in conversations)
                Cards.Add(new MailCardViewModel(conversation, _onMailSelected, _onCardDelete, _onSetRead));
        }
        else
        {
            foreach (var mail in _mails.OrderByDescending(m => m.CreationTime))
                Cards.Add(new MailCardViewModel(mail, _onMailSelected, _onCardDelete, _onSetRead));
        }
    }

    /// <summary>Restates the cards holding these mails after their read state changed.</summary>
    public void RefreshMailState(IEnumerable<string> entryIds)
    {
        var ids = entryIds.ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (ids.Count == 0)
            return;

        foreach (var card in Cards)
            card.RefreshState(ids);
    }

    public void DropMails(IEnumerable<string> entryIds)
    {
        var ids = entryIds.ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (ids.Count == 0)
            return;

        _mails.RemoveAll(m => ids.Contains(m.EntryId));

        foreach (var card in Cards.ToList())
        {
            if (card.RemoveMessages(ids))
                Cards.Remove(card);
        }
    }
}
