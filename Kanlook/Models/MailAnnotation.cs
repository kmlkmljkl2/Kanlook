namespace Kanlook.Models;

/// <summary>
/// What the user added to a mail themselves - a note and a priority flag. Deliberately kept out of
/// Outlook: neither is written back to the item, so a note can say something the sender should never
/// see, and the flag stays independent of the importance the sender chose.
/// </summary>
public sealed class MailAnnotation
{
    public string Comment { get; set; } = "";

    /// <summary>The user's own "deal with this first", unrelated to <see cref="MailSummary.Importance"/>.</summary>
    public bool IsHighPriority { get; set; }

    /// <summary>Nothing left worth storing - the entry can go rather than sit in the state file empty.</summary>
    public bool IsEmpty => !IsHighPriority && Comment.Length == 0;
}
