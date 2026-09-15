namespace Kanlook.Models;

/// <summary>User-facing options from the settings dialog. Persisted alongside the board state.</summary>
public sealed class AppSettings
{
    /// <summary>Which colour scheme to paint the app in. Follows Windows unless told otherwise.</summary>
    public AppTheme Theme { get; set; } = AppTheme.System;

    /// <summary>When true, mails of the same conversation share a single card instead of one card each.</summary>
    public bool GroupByConversation { get; set; }

    /// <summary>
    /// When true, cards the user marked high priority sort above the rest of their column instead of
    /// taking their place by date. On by default - flagging a mail and not seeing it move would be
    /// the surprising half.
    /// </summary>
    public bool PinHighPriority { get; set; } = true;
}
