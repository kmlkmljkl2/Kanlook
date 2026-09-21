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

    /// <summary>
    /// The preview line on the card. Empty until the body has been read - a folder's mail arrives
    /// from a MAPI table, which can carry everything about a mail except its body.
    /// </summary>
    public string Snippet { get; private set; } = "";

    /// <summary>Body text kept for searching (capped). Arrives with <see cref="Snippet"/>.</summary>
    public string SearchBody { get; private set; } = "";

    /// <summary>
    /// Whether the body has been read yet. False on a freshly listed folder, and set once
    /// <see cref="SetBodyText"/> has run - which is how <c>MailBodyLoader</c> knows what's left.
    /// </summary>
    public bool BodyLoaded { get; private set; }

    private const int SnippetMaxChars = 160;
    private const int SearchBodyMaxChars = 8_000;

    /// <summary>
    /// Records the mail's body: a one-line snippet for the card and a capped copy for searching.
    /// Both are flattened, because a card shows one line and a search doesn't care about layout.
    /// </summary>
    public void SetBodyText(string? body)
    {
        var flattened = (body ?? "").Replace('\r', ' ').Replace('\n', ' ').Trim();

        Snippet = flattened.Length > SnippetMaxChars ? flattened[..SnippetMaxChars] + "…" : flattened;
        SearchBody = flattened.Length > SearchBodyMaxChars ? flattened[..SearchBodyMaxChars] : flattened;
        BodyLoaded = true;
    }

    public bool HasAttachments { get; init; }
    public MailImportance Importance { get; init; } = MailImportance.Normal;

    /// <summary>Kept in sync with Outlook, so marking a mail read there clears it here too.</summary>
    public bool IsRead { get; set; }

    /// <summary>Outlook's comma-separated category list. Kept in sync with Outlook.</summary>
    public string Categories { get; set; } = "";

    public IReadOnlyList<string> CategoryNames => string.IsNullOrWhiteSpace(Categories)
        ? []
        : Categories.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);

    /// <summary>
    /// Outlook's own conversation id. Empty when the store/item doesn't expose one. Upper-cased on
    /// the way in: it reaches us as hex either from the object model or from a table's raw bytes,
    /// and two spellings of one id would split a conversation across two cards.
    /// </summary>
    public string ConversationId
    {
        get => _conversationId;
        init => _conversationId = value.ToUpperInvariant();
    }

    private readonly string _conversationId = "";

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
