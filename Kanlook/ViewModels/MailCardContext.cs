using Kanlook.Models;
using Kanlook.Services;

namespace Kanlook.ViewModels;

/// <summary>
/// Everything a <see cref="MailCardViewModel"/> needs from whoever is showing it. Bundled rather
/// than passed as a row of loose callbacks, because both the board and the shared-mailbox list build
/// cards and would otherwise have to keep the same argument order in their heads.
/// </summary>
public sealed class MailCardContext
{
    public required Action<MailSummary> OnSelect { get; init; }
    public required Action<MailCardViewModel> OnDelete { get; init; }
    public required Action<IReadOnlyList<MailSummary>, bool> OnSetRead { get; init; }

    /// <summary>Where the card reads and writes the user's note and priority flag.</summary>
    public required MailAnnotationStore Annotations { get; init; }

    /// <summary>
    /// Called after a card's priority flag flipped, so a host that pins high priority to the top can
    /// move the card there. Null when nothing re-sorts.
    /// </summary>
    public Action<MailCardViewModel>? OnPriorityChanged { get; init; }
}
