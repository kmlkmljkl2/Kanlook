using Kanlook.Models;

namespace Kanlook.Services;

/// <summary>
/// Every call runs on the service's own STA thread (see <see cref="ComWorker"/>) and returns a task
/// the UI thread can await, so no Outlook round-trip ever blocks the window.
/// </summary>
public interface IOutlookService : IDisposable
{
    /// <summary>Connects to the locally running/installed Outlook. Throws if unavailable.</summary>
    Task ConnectAsync();

    /// <summary>
    /// One node per Outlook store (personal mailbox + any shared mailboxes), without its subtree.
    /// Cheap: walking a whole mailbox's folder hierarchy up front is what made opening the app slow,
    /// so children are fetched a level at a time with <see cref="GetChildFoldersAsync"/>.
    /// </summary>
    Task<List<MailFolderNode>> GetStoreRootsAsync();

    /// <summary>The immediate subfolders of one folder. Empty when it has none or can't be opened.</summary>
    Task<List<MailFolderNode>> GetChildFoldersAsync(string storeId, string folderEntryId);

    /// <summary>
    /// Latest mail items in a folder, newest first, capped at <paramref name="maxCount"/>. Read from
    /// the folder's MAPI table in a single round-trip, which is why it returns in about a second even
    /// on a large online mailbox. The body is the one thing a table can't carry, so
    /// <see cref="MailSummary.Snippet"/> and <see cref="MailSummary.SearchBody"/> may arrive empty
    /// and get filled in afterwards by <see cref="MailBodyLoader"/>.
    /// </summary>
    Task<List<MailSummary>> GetMailSummariesAsync(string storeId, string folderEntryId, int maxCount);

    /// <summary>
    /// Cheap batched read of a folder's newest <paramref name="maxCount"/> mails - their ids plus the
    /// properties that change under us in Outlook. Used to spot items deleted or moved away, and
    /// category / read-state edits, without opening every item.
    /// </summary>
    Task<List<MailItemState>> GetFolderStateAsync(string storeId, string folderEntryId, int maxCount);

    /// <summary>Full summary for one mail. Null when the id no longer refers to a mail item.</summary>
    Task<MailSummary?> GetMailSummaryAsync(string storeId, string entryId);

    /// <summary>
    /// Plain-text body of one mail, for the card's preview line and body search. Opening the item is
    /// a round-trip each, so this is background work - see <see cref="MailBodyLoader"/>.
    /// </summary>
    Task<string?> GetBodyTextAsync(string storeId, string entryId);

    /// <summary>
    /// Entry id of a store's Sent Items folder, so replies can be read back with
    /// <see cref="GetMailSummariesAsync"/>. Null when the store has no such folder (some shared stores).
    /// </summary>
    Task<string?> GetSentItemsFolderIdAsync(string storeId);

    /// <summary>Outlook's master category list, as category name -&gt; display hex.</summary>
    Task<IReadOnlyDictionary<string, string>> GetCategoryColorsAsync();

    /// <summary>
    /// Lazily fetches what the reading pane shows: the body with its inline pictures made to
    /// display, and the attachments left over once those are accounted for.
    /// </summary>
    Task<MailContent> GetMailContentAsync(string storeId, string entryId);

    /// <summary>
    /// Every attachment of a mail, inline pictures included - for indexing, where a picture the
    /// body happens to display is still a file with a searchable name.
    /// </summary>
    Task<List<AttachmentInfo>> GetAttachmentsAsync(string storeId, string entryId);

    /// <summary>
    /// Extracts the attachment to a temp file and opens it with its default application.
    /// Returns the attachment's file name.
    /// </summary>
    Task<string> OpenAttachmentAsync(string storeId, string entryId, int attachmentIndex);

    /// <summary>Extracts the attachment to a temp file and returns its path, without opening it.</summary>
    Task<string> SaveAttachmentAsync(string storeId, string entryId, int attachmentIndex);

    /// <summary>Moves the mail to its store's Deleted Items folder, like Outlook's own Delete.</summary>
    Task DeleteMailAsync(string storeId, string entryId);

    /// <summary>Marks the mail read or unread in Outlook, so the change shows up there too.</summary>
    Task SetReadAsync(string storeId, string entryId, bool isRead);

    /// <summary>Opens Outlook's own Reply compose window for the given mail.</summary>
    Task ReplyAsync(string storeId, string entryId);

    /// <summary>Opens Outlook's own Reply All compose window for the given mail.</summary>
    Task ReplyAllAsync(string storeId, string entryId);

    /// <summary>Opens Outlook's own Forward compose window for the given mail.</summary>
    Task ForwardAsync(string storeId, string entryId);
}
