using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Threading;
using Kanlook.Models;

namespace Kanlook.Services;

/// <summary>
/// Talks to the locally installed/running Outlook via late-bound COM automation (no PIA/type-library
/// reference needed - just the "Outlook.Application" ProgID).
///
/// COM is apartment-bound, so every call is marshalled onto the service's own STA thread (see
/// <see cref="ComWorker"/>) and handed back as a task. Nothing here touches the UI thread, and the
/// <c>*Core</c> methods below only ever run on that one thread - which is also what keeps the COM
/// objects safe without locking.
/// </summary>
public sealed class OutlookService : IOutlookService
{
    /// <summary>
    /// MAPI's long-term entry id for contents tables. A folder table's own "EntryID" column holds a
    /// short-term id that does NOT match <c>MailItem.EntryID</c>; this property does.
    /// </summary>
    private const string PrLongTermEntryIdFromTable = "http://schemas.microsoft.com/mapi/proptag/0x66700102";

    private const string PrSenderSmtpAddress = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F";
    private const string PrSenderEmailAddress = "http://schemas.microsoft.com/mapi/proptag/0x0C1F001F";
    private const string PrDisplayTo = "http://schemas.microsoft.com/mapi/proptag/0x0E04001F";
    private const string PrConversationId = "http://schemas.microsoft.com/mapi/proptag/0x30130102";
    private const string PrHasAttachment = "http://schemas.microsoft.com/mapi/proptag/0x0E1B000B";

    private const int OlDescending = 2;

    /// <summary>OlDefaultFolders.olFolderSentMail.</summary>
    private const int OlFolderSentMail = 5;

    private readonly ComWorker _worker = new();
    private readonly Dictionary<string, string> _categoryColors = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Table columns a store refused, per store id. Stores differ in what they'll hand over and a
    /// refusal costs a COM exception, so we only ever pay for each one once.
    /// </summary>
    private readonly Dictionary<string, HashSet<string>> _unsupportedColumns = new(StringComparer.Ordinal);

    private dynamic? _app;
    private dynamic? _ns;
    private string? _defaultStoreId;

    private Task<T> Run<T>(Func<T> work, DispatcherPriority priority = DispatcherPriority.Normal) =>
        _worker.RunAsync(() => InvokeWithRetry(work), priority);

    private Task Run(Action work, DispatcherPriority priority = DispatcherPriority.Normal) =>
        _worker.RunAsync(() => InvokeWithRetry(work), priority);

    public Task ConnectAsync() => _worker.RunAsync(Connect);

    private void Connect()
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

    public Task<IReadOnlyDictionary<string, string>> GetCategoryColorsAsync() =>
        Run<IReadOnlyDictionary<string, string>>(() => new Dictionary<string, string>(_categoryColors, StringComparer.OrdinalIgnoreCase));

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

    public Task<List<MailFolderNode>> GetStoreRootsAsync() => Run(GetStoreRootsCore);

    private List<MailFolderNode> GetStoreRootsCore()
    {
        EnsureConnected();
        var roots = new List<MailFolderNode>();

        foreach (dynamic store in _ns!.Stores)
        {
            try
            {
                string storeId = store.StoreID;
                dynamic root = store.GetRootFolder();
                try
                {
                    roots.Add(ToNode(root, storeId != _defaultStoreId));
                }
                finally
                {
                    ReleaseCom(root);
                }
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

    public Task<List<MailFolderNode>> GetChildFoldersAsync(string storeId, string folderEntryId) =>
        Run(() => GetChildFoldersCore(storeId, folderEntryId));

    private List<MailFolderNode> GetChildFoldersCore(string storeId, string folderEntryId)
    {
        EnsureConnected();
        var children = new List<MailFolderNode>();
        var isShared = storeId != _defaultStoreId;

        dynamic folder = _ns!.GetFolderFromID(folderEntryId, storeId);
        try
        {
            dynamic folders = folder.Folders;
            try
            {
                foreach (dynamic child in folders)
                {
                    try
                    {
                        children.Add(ToNode(child, isShared));
                    }
                    catch (COMException)
                    {
                        // Skip folders we can't read (e.g. permission-restricted).
                    }
                    finally
                    {
                        ReleaseCom(child);
                    }
                }
            }
            finally
            {
                ReleaseCom(folders);
            }
        }
        finally
        {
            ReleaseCom(folder);
        }

        return children;
    }

    private static MailFolderNode ToNode(dynamic folder, bool isShared) => new()
    {
        EntryId = (string)folder.EntryID,
        StoreId = (string)folder.StoreID,
        Name = (string)folder.Name,
        IsSharedMailbox = isShared,
        HasChildren = HasChildFolders(folder),
    };

    /// <summary>
    /// Whether a folder is worth an expander. Counting is one round-trip; enumerating the children
    /// to find out would be the whole cost of the eager tree we're trying to avoid.
    /// </summary>
    private static bool HasChildFolders(dynamic folder)
    {
        try
        {
            dynamic folders = folder.Folders;
            try
            {
                return (int)folders.Count > 0;
            }
            finally
            {
                ReleaseCom(folders);
            }
        }
        catch (COMException)
        {
            return false;
        }
    }

    public Task<List<MailSummary>> GetMailSummariesAsync(string storeId, string folderEntryId, int maxCount) =>
        Run(() => GetMailSummariesCore(storeId, folderEntryId, maxCount));

    /// <summary>
    /// One table read for the whole window. Walking <c>Folder.Items</c> and opening each mail
    /// instead costs a cross-process call per property per item - on a large mailbox that is
    /// minutes, which is what this replaces. The body is the one field a table can't carry;
    /// <see cref="MailBodyLoader"/> fills it in afterwards.
    /// </summary>
    private List<MailSummary> GetMailSummariesCore(string storeId, string folderEntryId, int maxCount)
    {
        EnsureConnected();
        var result = new List<MailSummary>();

        dynamic folder = _ns!.GetFolderFromID(folderEntryId, storeId);
        try
        {
            dynamic table = folder.GetTable();
            try
            {
                dynamic columns = table.Columns;
                columns.RemoveAll();

                var schema = new TableSchema(columns, UnsupportedColumnsFor(storeId));
                var entryId = schema.Add(PrLongTermEntryIdFromTable);
                var messageClass = schema.Add("MessageClass");
                var subject = schema.Add("Subject");
                var senderName = schema.Add("SenderName");

                // The SMTP address where the store keeps one; Outlook's own SenderEmailAddress hands
                // back an X500 path for Exchange senders, which is no use on a card.
                var senderEmail = schema.AddAny(PrSenderSmtpAddress, "SenderEmailAddress", PrSenderEmailAddress);
                var toNames = schema.AddAny("To", PrDisplayTo);
                var receivedTime = schema.Add("ReceivedTime");
                var creationTime = schema.Add("CreationTime");
                var unread = schema.Add("UnRead");
                var categories = schema.Add("Categories");
                var conversationTopic = schema.Add("ConversationTopic");
                var conversationId = schema.Add(PrConversationId);
                var importance = schema.Add("Importance");
                var hasAttachment = schema.AddAny(PrHasAttachment, "HasAttachment");
                ReleaseCom(columns);

                table.Sort("[ReceivedTime]", OlDescending);

                if (table.GetArray(maxCount) is not object[,] rows)
                    return result;

                var firstColumn = rows.GetLowerBound(1);
                for (var index = rows.GetLowerBound(0); index <= rows.GetUpperBound(0); index++)
                {
                    var row = new TableRow(rows, index, firstColumn);

                    var id = row.Hex(entryId);
                    if (id.Length == 0 || !IsMailClass(row.Text(messageClass)))
                        continue;

                    var senderDisplay = row.Text(senderName);
                    var mailSubject = row.Text(subject);

                    result.Add(new MailSummary
                    {
                        EntryId = id,
                        StoreId = storeId,
                        Subject = mailSubject.Length == 0 ? "(no subject)" : mailSubject,
                        SenderName = senderDisplay.Length == 0 ? "(unknown sender)" : senderDisplay,
                        SenderEmail = row.Text(senderEmail),
                        ToNames = row.Text(toNames),
                        ReceivedTime = row.Time(receivedTime),
                        CreationTime = row.Time(creationTime),
                        IsRead = !row.Flag(unread),
                        Categories = row.Text(categories),
                        ConversationTopic = row.Text(conversationTopic),
                        ConversationId = row.Hex(conversationId),
                        HasAttachments = row.Flag(hasAttachment),
                        Importance = ToImportance(row.Number(importance, 1)),
                    });
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

    private HashSet<string> UnsupportedColumnsFor(string storeId)
    {
        if (!_unsupportedColumns.TryGetValue(storeId, out var refused))
            _unsupportedColumns[storeId] = refused = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        return refused;
    }

    public Task<List<MailItemState>> GetFolderStateAsync(string storeId, string folderEntryId, int maxCount) =>
        Run(() => GetFolderStateCore(storeId, folderEntryId, maxCount));

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
                for (var index = rows.GetLowerBound(0); index <= rows.GetUpperBound(0); index++)
                {
                    var row = new TableRow(rows, index, firstColumn);

                    var entryId = row.Hex(0);
                    if (entryId.Length == 0)
                        continue;

                    // Non-mail rows are reported too (flagged), so the caller can see how far back
                    // the snapshot reaches - it only ever holds maxCount rows.
                    result.Add(new MailItemState(
                        entryId,
                        row.Time(2),
                        row.Text(3),
                        !row.Flag(4),
                        IsMailClass(row.Text(1))));
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

    public Task<MailSummary?> GetMailSummaryAsync(string storeId, string entryId) =>
        Run(() => GetMailSummaryCore(storeId, entryId));

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

    public Task<string?> GetBodyTextAsync(string storeId, string entryId) =>
        Run(() => GetBodyTextCore(storeId, entryId), DispatcherPriority.Background);

    private string? GetBodyTextCore(string storeId, string entryId)
    {
        EnsureConnected();
        dynamic item = _ns!.GetItemFromID(entryId, storeId);
        try
        {
            return IsMailItem(item) ? (string?)item.Body : null;
        }
        finally
        {
            ReleaseCom(item);
        }
    }

    public Task<string?> GetSentItemsFolderIdAsync(string storeId) =>
        Run(() => GetSentItemsFolderIdCore(storeId), DispatcherPriority.Background);

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

    /// <summary>Matches <see cref="IsMailItem"/>, from a table's MessageClass column.</summary>
    private static bool IsMailClass(string messageClass) =>
        messageClass.StartsWith("IPM.Note", StringComparison.OrdinalIgnoreCase);

    private static bool IsMailItem(dynamic item)
    {
        try
        {
            // MailItem.Class == 43 (olMail). Other item types - meeting requests, reports - aren't board material.
            return (int)item.Class == 43;
        }
        catch (Exception)
        {
            return false;
        }
    }

    private static MailImportance ToImportance(int olImportance) => olImportance switch
    {
        2 => MailImportance.High,
        0 => MailImportance.Low,
        _ => MailImportance.Normal,
    };

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

    /// <summary>
    /// Builds a summary by opening the item. Only used for single mails - a new arrival, or a reply
    /// picked up in Sent Items - where the extra round-trips buy a body straight away.
    /// </summary>
    private static MailSummary ToSummary(dynamic mail, string storeId)
    {
        string subject = mail.Subject ?? "";
        string senderName = mail.SenderName ?? "";

        var summary = new MailSummary
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
            IsRead = !(bool)mail.UnRead,
            HasAttachments = mail.Attachments != null && (int)mail.Attachments.Count > 0,
            Importance = ToImportance((int)mail.Importance),
        };

        summary.SetBodyText(TryGetString(() => mail.Body));
        return summary;
    }

    public Task<string?> GetHtmlBodyAsync(string storeId, string entryId) =>
        Run(() => GetHtmlBodyCore(storeId, entryId));

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

    public Task<List<AttachmentInfo>> GetAttachmentsAsync(string storeId, string entryId) =>
        Run(() => GetAttachmentsCore(storeId, entryId));

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

    public async Task<string> OpenAttachmentAsync(string storeId, string entryId, int attachmentIndex)
    {
        var path = await SaveAttachmentAsync(storeId, entryId, attachmentIndex);
        Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        return Path.GetFileName(path);
    }

    public Task<string> SaveAttachmentAsync(string storeId, string entryId, int attachmentIndex) =>
        Run(() => SaveAttachmentCore(storeId, entryId, attachmentIndex));

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

    public Task DeleteMailAsync(string storeId, string entryId) =>
        Run(() => DeleteMailCore(storeId, entryId));

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

    public Task SetReadAsync(string storeId, string entryId, bool isRead) =>
        Run(() => SetReadCore(storeId, entryId, isRead));

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

    public Task ReplyAsync(string storeId, string entryId) =>
        Run(() => RespondTo(storeId, entryId, mail => mail.Reply()));

    public Task ReplyAllAsync(string storeId, string entryId) =>
        Run(() => RespondTo(storeId, entryId, mail => mail.ReplyAll()));

    public Task ForwardAsync(string storeId, string entryId) =>
        Run(() => RespondTo(storeId, entryId, mail => mail.Forward()));

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
            throw new InvalidOperationException("OutlookService.ConnectAsync() must be called first.");
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
        try
        {
            // The RCWs belong to the worker thread and have to be let go there.
            _worker.RunAsync(() =>
            {
                if (_ns is not null) ReleaseCom(_ns);
                if (_app is not null) ReleaseCom(_app);
                _ns = null;
                _app = null;
            }).Wait(TimeSpan.FromSeconds(5));
        }
        catch (Exception)
        {
            // Shutting down - a wedged Outlook must not stop the app from closing.
        }

        _worker.Dispose();
    }

    /// <summary>
    /// Builds a folder table's column set one property at a time and remembers where each landed.
    /// Outlook blocks some properties outright and stores differ in what they'll hand over, so a
    /// refused column costs that one field rather than the whole read.
    /// </summary>
    private sealed class TableSchema
    {
        /// <summary>Column index standing for "this store wouldn't give us the property".</summary>
        public const int Missing = -1;

        private readonly dynamic _columns;
        private readonly HashSet<string> _unsupported;
        private int _next;

        public TableSchema(dynamic columns, HashSet<string> unsupported)
        {
            _columns = columns;
            _unsupported = unsupported;
        }

        public int Add(string property)
        {
            if (_unsupported.Contains(property))
                return Missing;

            try
            {
                _columns.Add(property);
                return _next++;
            }
            catch (Exception)
            {
                _unsupported.Add(property);
                return Missing;
            }
        }

        /// <summary>Takes the first of several spellings of one field that the store accepts.</summary>
        public int AddAny(params string[] properties)
        {
            foreach (var property in properties)
            {
                var column = Add(property);
                if (column != Missing)
                    return column;
            }

            return Missing;
        }
    }

    /// <summary>One row of a <c>Table.GetArray</c> result, read by the column indices above.</summary>
    private readonly struct TableRow(object[,] rows, int row, int firstColumn)
    {
        private object? Value(int column) =>
            column < 0 ? null : rows[row, firstColumn + column];

        public string Text(int column) => Value(column) as string ?? "";

        public bool Flag(int column) => Value(column) is true;

        public int Number(int column, int fallback) => Value(column) is int value ? value : fallback;

        public DateTime Time(int column) => Value(column) as DateTime? ?? DateTime.MinValue;

        /// <summary>Binary ids come back as bytes, which is the form the object model spells in hex.</summary>
        public string Hex(int column) => Value(column) switch
        {
            byte[] bytes => Convert.ToHexString(bytes),
            string text => text,
            _ => "",
        };
    }
}
