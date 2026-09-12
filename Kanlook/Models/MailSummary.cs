using System.Text.RegularExpressions;

namespace Kanlook.Models;

public sealed partial class MailSummary
{
    public required string EntryId { get; init; }
    public required string StoreId { get; init; }
    public required string Subject { get; init; }
    public required string SenderName { get; init; }
    public string SenderEmail { get; init; } = "";
    public string ToNames { get; init; } = "";
    public required DateTime ReceivedTime { get; init; }

    /// <summary>Outlook's CreationTime - what the board sorts on.</summary>
    public required DateTime CreationTime { get; init; }

    public string Snippet { get; init; } = "";

    /// <summary>
    /// Body text kept for searching (capped). Free to collect: the body is already read to build
    /// <see cref="Snippet"/>.
    /// </summary>
    public string SearchBody { get; init; } = "";
    public bool HasAttachments { get; init; }
    public MailImportance Importance { get; init; } = MailImportance.Normal;

    /// <summary>Kept in sync with Outlook, so marking a mail read there clears it here too.</summary>
    public bool IsRead { get; set; }

    /// <summary>Outlook's comma-separated category list. Kept in sync with Outlook.</summary>
    public string Categories { get; set; } = "";

    public IReadOnlyList<string> CategoryNames => string.IsNullOrWhiteSpace(Categories)
        ? []
        : Categories.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);

    /// <summary>Outlook's own conversation id. Empty when the store/item doesn't expose one.</summary>
    public string ConversationId { get; init; } = "";

    /// <summary>Outlook's conversation topic (the subject with Re:/Fw: prefixes already stripped).</summary>
    public string ConversationTopic { get; init; } = "";

    /// <summary>Lazily populated the first time the mail is previewed.</summary>
    public string? HtmlBody { get; set; }

    /// <summary>
    /// Key used to group mails into one conversation tile. Prefers Outlook's conversation id and
    /// falls back to the reply-prefix-stripped topic so mails from stores without conversation
    /// support (or older items) still thread together.
    /// </summary>
    public string ConversationKey => string.IsNullOrEmpty(ConversationId)
        ? "topic:" + NormalizeTopic(string.IsNullOrEmpty(ConversationTopic) ? Subject : ConversationTopic)
        : ConversationId;

    private static string NormalizeTopic(string topic) =>
        ReplyPrefixRegex().Replace(topic, "").Trim().ToLowerInvariant();

    [GeneratedRegex(@"^\s*((re|aw|fw|fwd|wg|antw)\s*(\[\d+\])?\s*:\s*)+", RegexOptions.IgnoreCase)]
    private static partial Regex ReplyPrefixRegex();
}

public enum MailImportance
{
    Low,
    Normal,
    High
}
