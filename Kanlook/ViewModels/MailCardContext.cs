using Kanlook.Models;
using Kanlook.Services;

namespace Kanlook.ViewModels;

/// <summary>
/// What a card needs from the app as a whole, rather than from the board or list showing it. Passed
/// down as one value so adding an action doesn't mean threading another argument through every
/// folder view's constructor.
/// </summary>
/// <param name="Select">Open a message in the reading pane.</param>
/// <param name="Delete">Move a card's messages to Deleted Items.</param>
/// <param name="SetRead">Mark messages read or unread in Outlook.</param>
/// <param name="AnnotationsChanged">
/// A card's note or priority flag changed, so anything else showing that mail can catch up.
/// </param>
public sealed record MailCardActions(
    Action<MailSummary> Select,
    Action<MailCardViewModel> Delete,
    Action<IReadOnlyList<MailSummary>, bool> SetRead,
    Action<MailCardViewModel> AnnotationsChanged);

/// <summary>
/// Everything a <see cref="MailCardViewModel"/> needs, with the host's own hooks filled in. Bundled
/// rather than passed as a row of loose callbacks, because both the board and the shared-mailbox
/// list build cards and would otherwise have to keep the same argument order in their heads.
/// </summary>
public sealed class MailCardContext
{
    public required MailCardActions Actions { get; init; }

    /// <summary>Where the card reads and writes the user's note and priority flag.</summary>
    public required MailAnnotationStore Annotations { get; init; }

    /// <summary>
    /// Called after a card's priority flag flipped, so a host that pins high priority to the top can
    /// move the card there.
    /// </summary>
    public required Action<MailCardViewModel> OnPriorityChanged { get; init; }
}
