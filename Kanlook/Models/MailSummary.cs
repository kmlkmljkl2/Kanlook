namespace Kanlook.Models;

public sealed class MailSummary
{
    public required string EntryId { get; init; }
    public required string StoreId { get; init; }
    public required string Subject { get; init; }
    public required string SenderName { get; init; }
    public string SenderEmail { get; init; } = "";
    public string ToNames { get; init; } = "";
    public required DateTime ReceivedTime { get; init; }
    public string Snippet { get; init; } = "";
    public bool IsRead { get; init; }
    public bool HasAttachments { get; init; }
    public MailImportance Importance { get; init; } = MailImportance.Normal;

    /// <summary>Lazily populated the first time the mail is previewed.</summary>
    public string? HtmlBody { get; set; }
}

public enum MailImportance
{
    Low,
    Normal,
    High
}
