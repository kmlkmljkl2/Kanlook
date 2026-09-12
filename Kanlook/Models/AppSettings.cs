namespace Kanlook.Models;

/// <summary>User-facing options from the settings dialog. Persisted alongside the board state.</summary>
public sealed class AppSettings
{
    /// <summary>When true, mails of the same conversation share a single card instead of one card each.</summary>
    public bool GroupByConversation { get; set; }
}
