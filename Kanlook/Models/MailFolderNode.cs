namespace Kanlook.Models;

public sealed class MailFolderNode
{
    public required string EntryId { get; init; }
    public required string StoreId { get; init; }
    public required string Name { get; init; }
    public bool IsSharedMailbox { get; init; }

    /// <summary>
    /// Whether the folder has subfolders. They aren't read until the node is expanded - enumerating
    /// every mailbox's whole hierarchy up front is what used to keep the app from opening.
    /// </summary>
    public bool HasChildren { get; init; }
}
