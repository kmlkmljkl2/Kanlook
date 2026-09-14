using System.Collections.ObjectModel;

namespace Kanlook.ViewModels;

/// <summary>
/// The order mail cards read in: newest creation time first, with the ones the user flagged high
/// priority held above the rest when the setting asks for it. Shared by the board's columns and the
/// shared-mailbox list, so a flagged card lands in the same place in both.
/// </summary>
public static class MailCardOrder
{
    public static IEnumerable<MailCardViewModel> Sort(IEnumerable<MailCardViewModel> cards, bool pinHighPriority) =>
        pinHighPriority
            ? cards.OrderByDescending(c => c.IsHighPriority).ThenByDescending(c => c.CreationTime)
            : cards.OrderByDescending(c => c.CreationTime);

    /// <summary>
    /// Where a card belongs in a list already in this order. <paramref name="ignoring"/> skips an
    /// entry, which is how a card already in the list is placed without measuring against itself.
    /// </summary>
    public static int IndexFor(
        IList<MailCardViewModel> cards,
        MailCardViewModel card,
        bool pinHighPriority,
        MailCardViewModel? ignoring = null)
    {
        var index = 0;
        foreach (var existing in cards)
        {
            if (ReferenceEquals(existing, ignoring))
                continue;
            if (!RanksAbove(existing, card, pinHighPriority))
                break;
            index++;
        }

        return index;
    }

    /// <summary>
    /// Moves a card to where it now belongs - after its priority flag was toggled, or after a merge
    /// gave it a newer message and with it a new creation time.
    /// </summary>
    public static void Reposition(
        ObservableCollection<MailCardViewModel> cards,
        MailCardViewModel card,
        bool pinHighPriority)
    {
        var from = cards.IndexOf(card);
        if (from < 0)
            return;

        var to = IndexFor(cards, card, pinHighPriority, ignoring: card);
        if (from != to)
            cards.Move(from, Math.Clamp(to, 0, cards.Count - 1));
    }

    private static bool RanksAbove(MailCardViewModel card, MailCardViewModel other, bool pinHighPriority) =>
        pinHighPriority && card.IsHighPriority != other.IsHighPriority
            ? card.IsHighPriority
            : card.CreationTime > other.CreationTime;
}
