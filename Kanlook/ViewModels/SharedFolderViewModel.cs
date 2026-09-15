using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using Kanlook.Models;
using Kanlook.Services;

namespace Kanlook.ViewModels;

public sealed partial class SharedFolderViewModel : ObservableObject
{
    public string FolderName { get; }
    public ObservableCollection<MailCardViewModel> Cards { get; } = [];

    private readonly List<MailSummary> _mails;
    private readonly AppSettings _settings;
    private readonly MailCardContext _cardContext;

    public SharedFolderViewModel(
        string folderName,
        List<MailSummary> mails,
        BoardStateStore boardStore,
        MailCardActions actions)
    {
        FolderName = folderName;
        _mails = mails;
        _settings = boardStore.Settings;
        _cardContext = new MailCardContext
        {
            Actions = actions,
            Annotations = boardStore.Annotations,
            OnPriorityChanged = card => MailCardOrder.Reposition(Cards, card, _settings.PinHighPriority),
        };
        RebuildCards();
    }

    public void RebuildCards()
    {
        var cards = _settings.GroupByConversation
            ? _mails.GroupBy(m => m.ConversationKey).Select(g => new MailCardViewModel(g, _cardContext))
            : _mails.Select(m => new MailCardViewModel(m, _cardContext));

        Cards.Clear();
        foreach (var card in MailCardOrder.Sort(cards, _settings.PinHighPriority))
            Cards.Add(card);
    }

    /// <summary>Picks up a note or flag set in the reading pane, and re-sorts if the flag moved it.</summary>
    public void RefreshAnnotations(string entryId)
    {
        if (Cards.FirstOrDefault(c => c.Messages.Any(m => m.EntryId == entryId)) is not { } card)
            return;

        card.RefreshAnnotations();
        MailCardOrder.Reposition(Cards, card, _settings.PinHighPriority);
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
        _cardContext.Annotations.Forget(ids);

        foreach (var card in Cards.ToList())
        {
            if (card.RemoveMessages(ids))
                Cards.Remove(card);
        }
    }
}
