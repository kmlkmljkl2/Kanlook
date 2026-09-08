namespace Kanlook.Models;

public sealed class MailFolderNode
{
    public required string EntryId { get; init; }
    public required string StoreId { get; init; }
    public required string Name { get; init; }
    public bool IsSharedMailbox { get; init; }
    public List<MailFolderNode> Children { get; init; } = [];
}
