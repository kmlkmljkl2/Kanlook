using Kanlook.Models;

namespace Kanlook.Services;

public interface IOutlookService : IDisposable
{
    /// <summary>Connects to the locally running/installed Outlook. Throws if unavailable.</summary>
    void Connect();

    /// <summary>One root node per Outlook store (personal mailbox + any shared mailboxes), each with its full folder subtree.</summary>
    List<MailFolderNode> BuildFolderTree();

    /// <summary>Latest mail items in a folder, newest first, capped at <paramref name="maxCount"/>.</summary>
    List<MailSummary> GetMailSummaries(string storeId, string folderEntryId, int maxCount = 300);

    /// <summary>
    /// Cheap batched read of a folder's newest <paramref name="maxCount"/> mails - their ids plus the
    /// properties that change under us in Outlook. Used to spot items deleted or moved away, and
    /// category / read-state edits, without opening every item.
    /// </summary>
    List<MailItemState> GetFolderState(string storeId, string folderEntryId, int maxCount = 300);

    /// <summary>Full summary for one mail. Null when the id no longer refers to a mail item.</summary>
    MailSummary? GetMailSummary(string storeId, string entryId);

    /// <summary>Outlook's master category list, as category name -&gt; display hex.</summary>
    IReadOnlyDictionary<string, string> GetCategoryColors();

    /// <summary>Lazily fetches the HTML body of a single mail for preview.</summary>
    string? GetHtmlBody(string storeId, string entryId);

    /// <summary>Lazily fetches attachment names/sizes for a single mail for preview.</summary>
    List<AttachmentInfo> GetAttachments(string storeId, string entryId);

    /// <summary>
    /// Extracts the attachment to a temp file and opens it with its default application.
    /// Returns the attachment's file name.
    /// </summary>
    string OpenAttachment(string storeId, string entryId, int attachmentIndex);

    /// <summary>Opens Outlook's own Reply compose window for the given mail.</summary>
    void Reply(string storeId, string entryId);

    /// <summary>Opens Outlook's own Reply All compose window for the given mail.</summary>
    void ReplyAll(string storeId, string entryId);

    /// <summary>Opens Outlook's own Forward compose window for the given mail.</summary>
    void Forward(string storeId, string entryId);
}
