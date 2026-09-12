using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using Kanlook.Models;

namespace Kanlook.Services;

/// <summary>
/// Talks to the locally installed/running Outlook via late-bound COM automation (no PIA/type-library
/// reference needed - just the "Outlook.Application" ProgID). All calls must happen on the WPF UI
/// thread (STA), which is the apartment COM requires.
/// </summary>
public sealed class OutlookService : IOutlookService
{
    /// <summary>
    /// MAPI's long-term entry id for contents tables. A folder table's own "EntryID" column holds a
    /// short-term id that does NOT match <c>MailItem.EntryID</c>; this property does.
    /// </summary>
    private const string PrLongTermEntryIdFromTable = "http://schemas.microsoft.com/mapi/proptag/0x66700102";

    private const int OlDescending = 2;

    private dynamic? _app;
    private dynamic? _ns;
    private string? _defaultStoreId;
    private readonly Dictionary<string, string> _categoryColors = new(StringComparer.OrdinalIgnoreCase);

    public void Connect()
    {
        try
        {
            _app = Marshal.BindToMoniker("Outlook.Application");
        }
        catch (COMException)
        {
            var progType = Type.GetTypeFromProgID("Outlook.Application")
                ?? throw new InvalidOperationException("Outlook does not appear to be installed on this machine.");
            _app = Activator.CreateInstance(progType)
                ?? throw new InvalidOperationException("Could not start Outlook.");
        }

        _ns = _app!.GetNamespace("MAPI");
        _defaultStoreId = (string)_ns!.DefaultStore.StoreID;
        LoadCategoryColors();
    }

    /// <summary>Outlook's master category list, as category name -&gt; display hex.</summary>
    public IReadOnlyDictionary<string, string> GetCategoryColors() => _categoryColors;

    private void LoadCategoryColors()
    {
        _categoryColors.Clear();
        try
        {
            dynamic categories = _ns!.Categories;
            foreach (dynamic category in categories)
            {
                try
                {
                    _categoryColors[(string)category.Name] = CategoryColorHex((int)category.Color);
                }
                catch (COMException)
                {
                    // A single unreadable category shouldn't cost us the rest of the list.
                }
                finally
                {
                    ReleaseCom(category);
                }
            }

            ReleaseCom(categories);
        }
        catch (COMException)
        {
            // Categories are cosmetic - fall back to the neutral chip colour.
        }
    }

    /// <summary>Approximates Outlook's own swatch for an OlCategoryColor value.</summary>
    private static string CategoryColorHex(int olCategoryColor) => olCategoryColor switch
    {
        1 => "#D93F3C",  // Red
        2 => "#E8761B",  // Orange
        3 => "#F4B183",  // Peach
        4 => "#EFC000",  // Yellow
        5 => "#4CB782",  // Green
        6 => "#2FB6C4",  // Teal
        7 => "#9BA829",  // Olive
        8 => "#5B8DEF",  // Blue
        9 => "#B87CE0",  // Purple
        10 => "#A1436B", // Maroon
        11 => "#A6B1C2", // Steel
        12 => "#6E7B8B", // Dark Steel
        13 => "#B4B9C6", // Gray
        14 => "#7C8199", // Dark Gray
        15 => "#3B3F4C", // Black
        16 => "#A32C2A", // Dark Red
        17 => "#B35714", // Dark Orange
        18 => "#C98A5B", // Dark Peach
        19 => "#B08E00", // Dark Yellow
        20 => "#2F8A5E", // Dark Green
        21 => "#1F8792", // Dark Teal
        22 => "#6E7A1D", // Dark Olive
        23 => "#3A6FF7", // Dark Blue
        24 => "#8B5BB0", // Dark Purple
        25 => "#75304C", // Dark Maroon
        _ => "#8C93A6",  // None
    };

    public List<MailFolderNode> BuildFolderTree() => InvokeWithRetry(BuildFolderTreeCore);

    private List<MailFolderNode> BuildFolderTreeCore()
    {
        EnsureConnected();
        var roots = new List<MailFolderNode>();

        foreach (dynamic store in _ns!.Stores)
        {
            try
            {
                string storeId = store.StoreID;
                var isShared = storeId != _defaultStoreId;
                dynamic root = store.GetRootFolder();
                roots.Add(BuildNode(root, isShared));
                ReleaseCom(root);
            }
            catch (COMException)
            {
                // Some stores (e.g. public folders without permissions) can't be opened - skip.
            }
            finally
            {
                ReleaseCom(store);
            }
        }

        return roots;
    }

    private static MailFolderNode BuildNode(dynamic folder, bool isShared)
    {
        var node = new MailFolderNode
        {
            EntryId = (string)folder.EntryID,
            StoreId = (string)folder.StoreID,
            Name = (string)folder.Name,
            IsSharedMailbox = isShared,
        };

        try
        {
            foreach (dynamic child in folder.Folders)
            {
                try
                {
                    node.Children.Add(BuildNode(child, isShared));
                }
                catch (COMException)
                {
                    // Skip folders we can't enumerate (e.g. permission-restricted).
                }
                finally
                {
                    ReleaseCom(child);
                }
            }
        }
        catch (COMException) { }

        return node;
    }

    public List<MailSummary> GetMailSummaries(string storeId, string folderEntryId, int maxCount = 300) =>
        InvokeWithRetry(() => GetMailSummariesCore(storeId, folderEntryId, maxCount));

    private List<MailSummary> GetMailSummariesCore(string storeId, string folderEntryId, int maxCount)
    {
        EnsureConnected();
        var result = new List<MailSummary>();

        dynamic folder = _ns!.GetFolderFromID(folderEntryId, storeId);
        try
        {
            dynamic items = folder.Items;
            items.Sort("[ReceivedTime]", true);

            var count = 0;
            foreach (dynamic raw in items)
            {
                if (count >= maxCount)
                {
                    ReleaseCom(raw);
                    break;
                }

                // MailItem.Class == 43 (olMail); other item types (meeting requests, etc.) are skipped for this demo.
                if (IsMailItem(raw))
                {
                    result.Add(ToSummary(raw, storeId));
                    count++;
                }

                ReleaseCom(raw);
            }
        }
        finally
        {
            ReleaseCom(folder);
        }

        return result;
    }

    /// <summary>
    /// Reads an optional COM string property. Not every store/item exposes conversation properties
    /// (very old items, some public folders), and a missing one throws rather than returning null.
    /// </summary>
    private static string TryGetString(Func<dynamic> get)
    {
        try
        {
            return (string?)get() ?? "";
        }
        catch (Exception)
        {
            return "";
        }
    }

    public List<MailItemState> GetFolderState(string storeId, string folderEntryId, int maxCount = 300) =>
        InvokeWithRetry(() => GetFolderStateCore(storeId, folderEntryId, maxCount));

    private List<MailItemState> GetFolderStateCore(string storeId, string folderEntryId, int maxCount)
    {
        EnsureConnected();
        var result = new List<MailItemState>();

        dynamic folder = _ns!.GetFolderFromID(folderEntryId, storeId);
        try
        {
            dynamic table = folder.GetTable();
            try
            {
                dynamic columns = table.Columns;
                columns.RemoveAll();
                columns.Add(PrLongTermEntryIdFromTable);
                columns.Add("MessageClass");
                columns.Add("ReceivedTime");
                columns.Add("Categories");
                columns.Add("UnRead");
                ReleaseCom(columns);

                table.Sort("[ReceivedTime]", OlDescending);

                // One round-trip for the whole window. Walking Folder.Items instead costs a COM
                // call per property per item, which is far too slow to run on a timer.
                if (table.GetArray(maxCount) is not object[,] rows)
                    return result;

                var firstColumn = rows.GetLowerBound(1);
                for (var row = rows.GetLowerBound(0); row <= rows.GetUpperBound(0); row++)
                {
                    var entryId = rows[row, firstColumn] switch
                    {
                        byte[] bytes => Convert.ToHexString(bytes),
                        string text => text,
                        _ => null,
                    };

                    if (entryId is null)
                        continue;

                    // Matches IsMailItem. Non-mail rows are reported too (flagged), so the caller can
                    // see how far back the snapshot reaches - it only ever holds maxCount rows.
                    var isMail = rows[row, firstColumn + 1] is string messageClass &&
                                 messageClass.StartsWith("IPM.Note", StringComparison.OrdinalIgnoreCase);

                    result.Add(new MailItemState(
                        entryId,
                        rows[row, firstColumn + 2] as DateTime? ?? DateTime.MinValue,
                        rows[row, firstColumn + 3] as string ?? "",
                        rows[row, firstColumn + 4] is not bool unread || !unread,
                        isMail));
                }
            }
            finally
            {
                ReleaseCom(table);
            }
        }
        finally
        {
            ReleaseCom(folder);
        }

        return result;
    }

    public MailSummary? GetMailSummary(string storeId, string entryId) =>
        InvokeWithRetry(() => GetMailSummaryCore(storeId, entryId));

    private MailSummary? GetMailSummaryCore(string storeId, string entryId)
    {
        EnsureConnected();
        dynamic item = _ns!.GetItemFromID(entryId, storeId);
        try
        {
            return IsMailItem(item) ? ToSummary(item, storeId) : null;
        }
        finally
        {
            ReleaseCom(item);
        }
    }

    /// <summary>OlDefaultFolders.olFolderSentMail.</summary>
    private const int OlFolderSentMail = 5;

    public string? GetSentItemsFolderId(string storeId) =>
        InvokeWithRetry(() => GetSentItemsFolderIdCore(storeId));

    private string? GetSentItemsFolderIdCore(string storeId)
    {
        EnsureConnected();

        // Per-store, so a shared mailbox reports its own Sent Items rather than the user's.
        foreach (dynamic store in _ns!.Stores)
        {
            try
            {
                if ((string)store.StoreID != storeId)
                    continue;

                dynamic folder = store.GetDefaultFolder(OlFolderSentMail);
                try
                {
                    return (string)folder.EntryID;
                }
                finally
                {
                    ReleaseCom(folder);
                }
            }
            catch (COMException)
            {
                // Stores without a Sent Items folder (public folders, some delegated mailboxes) throw.
            }
            finally
            {
                ReleaseCom(store);
            }
        }

        if (storeId != _defaultStoreId)
            return null;

        try
        {
            dynamic fallback = _ns!.GetDefaultFolder(OlFolderSentMail);
            try
            {
                return (string)fallback.EntryID;
            }
            finally
            {
                ReleaseCom(fallback);
            }
        }
        catch (COMException)
        {
            return null;
        }
    }

    private static bool IsMailItem(dynamic item)
    {
        try
        {
            return (int)item.Class == 43;
        }
        catch (Exception)
        {
            return false;
        }
    }

    private const int SearchBodyMaxChars = 8_000;

    private static MailSummary ToSummary(dynamic mail, string storeId)
    {
        string body = mail.Body ?? "";
        var flattened = body.Replace('\r', ' ').Replace('\n', ' ').Trim();
        var snippet = flattened.Length > 160 ? flattened[..160] + "…" : flattened;
        var searchBody = flattened.Length > SearchBodyMaxChars ? flattened[..SearchBodyMaxChars] : flattened;

        string subject = mail.Subject ?? "";
        string senderName = mail.SenderName ?? "";
        int importance = (int)mail.Importance;

        return new MailSummary
        {
            ConversationId = TryGetString(() => mail.ConversationID),
            ConversationTopic = TryGetString(() => mail.ConversationTopic),
            Categories = TryGetString(() => mail.Categories),
            EntryId = mail.EntryID,
            StoreId = storeId,
            Subject = string.IsNullOrEmpty(subject) ? "(no subject)" : subject,
            SenderName = string.IsNullOrEmpty(senderName) ? "(unknown sender)" : senderName,
            SenderEmail = mail.SenderEmailAddress ?? "",
            ToNames = mail.To ?? "",
            ReceivedTime = mail.ReceivedTime,
            CreationTime = mail.CreationTime,
            Snippet = snippet,
            SearchBody = searchBody,
            IsRead = !(bool)mail.UnRead,
            HasAttachments = mail.Attachments != null && (int)mail.Attachments.Count > 0,
            Importance = importance switch
            {
                2 => MailImportance.High,
                0 => MailImportance.Low,
                _ => MailImportance.Normal,
            },
        };
    }

    public string? GetHtmlBody(string storeId, string entryId) => InvokeWithRetry(() => GetHtmlBodyCore(storeId, entryId));

    private string? GetHtmlBodyCore(string storeId, string entryId)
    {
        EnsureConnected();
        dynamic item = _ns!.GetItemFromID(entryId, storeId);
        try
        {
            if (!IsMailItem(item))
                return null;

            string? html = item.HTMLBody;
            return string.IsNullOrEmpty(html) ? (string?)item.Body : html;
        }
        finally
        {
            ReleaseCom(item);
        }
    }

    public List<AttachmentInfo> GetAttachments(string storeId, string entryId) =>
        InvokeWithRetry(() => GetAttachmentsCore(storeId, entryId));

    private List<AttachmentInfo> GetAttachmentsCore(string storeId, string entryId)
    {
        EnsureConnected();
        var result = new List<AttachmentInfo>();

        dynamic item = _ns!.GetItemFromID(entryId, storeId);
        try
        {
            if (!IsMailItem(item))
                return result;

            dynamic attachments = item.Attachments;
            foreach (dynamic att in attachments)
            {
                try
                {
                    string fileName = att.FileName ?? att.DisplayName ?? "(unnamed attachment)";
                    long size = (long)att.Size;
                    result.Add(new AttachmentInfo
                    {
                        FileName = fileName,
                        SizeDisplay = AttachmentInfo.FormatBytes(size),
                        Index = (int)att.Index,
                    });
                }
                finally
                {
                    ReleaseCom(att);
                }
            }
            ReleaseCom(attachments);
        }
        finally
        {
            ReleaseCom(item);
        }

        return result;
    }

    public string OpenAttachment(string storeId, string entryId, int attachmentIndex)
    {
        var path = SaveAttachment(storeId, entryId, attachmentIndex);
        Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        return Path.GetFileName(path);
    }

    public string SaveAttachment(string storeId, string entryId, int attachmentIndex) =>
        InvokeWithRetry(() => SaveAttachmentCore(storeId, entryId, attachmentIndex));

    /// <summary>
    /// Outlook can only hand an attachment over as a file, so extract it to a per-mail cache folder.
    /// </summary>
    private string SaveAttachmentCore(string storeId, string entryId, int attachmentIndex)
    {
        EnsureConnected();

        dynamic item = _ns!.GetItemFromID(entryId, storeId);
        try
        {
            dynamic attachments = item.Attachments;
            try
            {
                dynamic att = attachments.Item(attachmentIndex);
                try
                {
                    string fileName = att.FileName ?? att.DisplayName ?? "attachment";
                    var path = BuildCachePath(entryId, fileName);
                    att.SaveAsFile(path);
                    return path;
                }
                finally
                {
                    ReleaseCom(att);
                }
            }
            finally
            {
                ReleaseCom(attachments);
            }
        }
        finally
        {
            ReleaseCom(item);
        }
    }

    public void DeleteMail(string storeId, string entryId) =>
        InvokeWithRetry(() => DeleteMailCore(storeId, entryId));

    private void DeleteMailCore(string storeId, string entryId)
    {
        EnsureConnected();

        dynamic item = _ns!.GetItemFromID(entryId, storeId);
        try
        {
            // MailItem.Delete moves the mail to the store's Deleted Items folder, exactly like
            // pressing Delete in Outlook - so it stays recoverable from there.
            if (IsMailItem(item))
                item.Delete();
        }
        finally
        {
            ReleaseCom(item);
        }
    }

    public void SetRead(string storeId, string entryId, bool isRead) =>
        InvokeWithRetry(() => SetReadCore(storeId, entryId, isRead));

    private void SetReadCore(string storeId, string entryId, bool isRead)
    {
        EnsureConnected();

        dynamic item = _ns!.GetItemFromID(entryId, storeId);
        try
        {
            if (!IsMailItem(item))
                return;

            // Outlook models this the other way round, as UnRead. Save() commits it, otherwise the
            // change would live only on our copy of the item and never reach Outlook's own views.
            item.UnRead = !isRead;
            item.Save();
        }
        finally
        {
            ReleaseCom(item);
        }
    }

    /// <summary>
    /// Stable temp path per (mail, file name), so re-opening the same attachment reuses one file
    /// instead of littering temp, while same-named attachments of different mails stay separate.
    /// </summary>
    private static string BuildCachePath(string entryId, string fileName)
    {
        var hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(entryId)))[..12];
        var safeName = string.Concat(fileName.Select(c => Path.GetInvalidFileNameChars().Contains(c) ? '_' : c));
        if (string.IsNullOrWhiteSpace(safeName))
            safeName = "attachment";

        var dir = Path.Combine(Path.GetTempPath(), "Kanlook", hash);
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, safeName);
    }

    public void Reply(string storeId, string entryId) =>
        InvokeWithRetry(() => RespondTo(storeId, entryId, mail => mail.Reply()));

    public void ReplyAll(string storeId, string entryId) =>
        InvokeWithRetry(() => RespondTo(storeId, entryId, mail => mail.ReplyAll()));

    public void Forward(string storeId, string entryId) =>
        InvokeWithRetry(() => RespondTo(storeId, entryId, mail => mail.Forward()));

    private void RespondTo(string storeId, string entryId, Func<dynamic, dynamic> respond)
    {
        EnsureConnected();
        dynamic item = _ns!.GetItemFromID(entryId, storeId);
        try
        {
            if (!IsMailItem(item))
                return;

            // Don't release the response - it's now a visible, user-owned Inspector window.
            // Releasing our RCW here forces its COM ref count to zero while Outlook's own UI still
            // expects to use it (e.g. when the user later closes/discards it), which can corrupt
            // Outlook's COM state badly enough to break the whole automation server (RPC unavailable).
            dynamic response = respond(item);
            response.Display(false);
        }
        finally
        {
            ReleaseCom(item);
        }
    }

    private void EnsureConnected()
    {
        if (_app is null || _ns is null)
            throw new InvalidOperationException("OutlookService.Connect() must be called first.");
    }

    private const uint RpcServerUnavailableHResult = 0x800706BA;

    private T InvokeWithRetry<T>(Func<T> action)
    {
        try
        {
            return action();
        }
        catch (COMException ex) when ((uint)ex.HResult == RpcServerUnavailableHResult)
        {
            Connect();
            return action();
        }
    }

    private void InvokeWithRetry(Action action)
    {
        try
        {
            action();
        }
        catch (COMException ex) when ((uint)ex.HResult == RpcServerUnavailableHResult)
        {
            Connect();
            action();
        }
    }

    private static void ReleaseCom(object? comObject)
    {
        if (comObject is not null && Marshal.IsComObject(comObject))
            Marshal.ReleaseComObject(comObject);
    }

    public void Dispose()
    {
        if (_ns is not null) ReleaseCom(_ns);
        if (_app is not null) ReleaseCom(_app);
        _ns = null;
        _app = null;
    }
}
