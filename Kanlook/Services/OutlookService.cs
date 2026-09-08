using System.Runtime.InteropServices;
using Kanlook.Models;

namespace Kanlook.Services;

/// <summary>
/// Talks to the locally installed/running Outlook via late-bound COM automation (no PIA/type-library
/// reference needed - just the "Outlook.Application" ProgID). All calls must happen on the WPF UI
/// thread (STA), which is the apartment COM requires.
/// </summary>
public sealed class OutlookService : IOutlookService
{
    private dynamic? _app;
    private dynamic? _ns;
    private string? _defaultStoreId;

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
    }

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

    private static MailSummary ToSummary(dynamic mail, string storeId)
    {
        string body = mail.Body ?? "";
        var snippet = body.Replace('\r', ' ').Replace('\n', ' ').Trim();
        if (snippet.Length > 160)
            snippet = snippet[..160] + "…";

        string subject = mail.Subject ?? "";
        string senderName = mail.SenderName ?? "";
        int importance = (int)mail.Importance;

        return new MailSummary
        {
            EntryId = mail.EntryID,
            StoreId = storeId,
            Subject = string.IsNullOrEmpty(subject) ? "(no subject)" : subject,
            SenderName = string.IsNullOrEmpty(senderName) ? "(unknown sender)" : senderName,
            SenderEmail = mail.SenderEmailAddress ?? "",
            ToNames = mail.To ?? "",
            ReceivedTime = mail.ReceivedTime,
            Snippet = snippet,
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
