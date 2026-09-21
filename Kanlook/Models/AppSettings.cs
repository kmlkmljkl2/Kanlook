namespace Kanlook.Models;

/// <summary>User-facing options from the settings dialog. Persisted alongside the board state.</summary>
public sealed class AppSettings
{
    /// <summary>Smallest and largest number of mails a board will load, whatever the file says.</summary>
    public const int MinMailsPerFolder = 50;
    public const int MaxMailsPerFolder = 5_000;

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

    /// <summary>
    /// How many of a folder's newest mails a board loads, and how far back its sync looks. The
    /// listing itself is one table read whatever this is; what grows with it is the background pass
    /// that reads bodies for the card previews and search.
    /// </summary>
    public int MailsPerFolder
    {
        get => _mailsPerFolder;
        set => _mailsPerFolder = Math.Clamp(value, MinMailsPerFolder, MaxMailsPerFolder);
    }

    private int _mailsPerFolder = 300;

    /// <summary>Width of the folder sidebar, as the user last dragged it.</summary>
    public double SidebarWidth
    {
        get => _sidebarWidth;
        set => _sidebarWidth = Math.Clamp(value, MinSidebarWidth, MaxSidebarWidth);
    }

    private double _sidebarWidth = 250;

    public const double MinSidebarWidth = 170;
    public const double MaxSidebarWidth = 560;
}
