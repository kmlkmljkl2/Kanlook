namespace Kanlook.Models;

/// <summary>
/// Per-item state read straight from Outlook's folder table - enough to notice that an item was
/// deleted or moved away, or had its categories / read state changed in Outlook, without opening
/// the item itself.
/// </summary>
/// <param name="IsMail">
/// False for items the board never shows (meeting requests, reports, ...). They're still reported so
/// the caller can tell how far back the snapshot reaches.
/// </param>
public sealed record MailItemState(
    string EntryId,
    DateTime ReceivedTime,
    string Categories,
    bool IsRead,
    bool IsMail);
