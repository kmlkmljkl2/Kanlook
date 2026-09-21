using System.IO;
using System.Text;
using Kanlook.Models;

namespace Kanlook.Services;

/// <summary>
/// Fills the <see cref="AttachmentIndex"/> in the background so attachment contents become
/// searchable. Both halves stay off the UI thread: pulling the attachment out runs on the Outlook
/// service's own thread at background priority, and the parsing goes to the thread pool.
/// </summary>
public sealed class AttachmentIndexer
{
    private readonly IOutlookService _outlook;
    private readonly AttachmentIndex _index;
    private readonly Queue<MailSummary> _pending = new();
    private readonly HashSet<string> _queuedEntryIds = new(StringComparer.OrdinalIgnoreCase);

    private bool _running;

    /// <summary>Raised when a mail has been indexed, and when the queue drains.</summary>
    public event Action? Progressed;

    public int PendingCount => _pending.Count;

    public AttachmentIndexer(IOutlookService outlook, AttachmentIndex index)
    {
        _outlook = outlook;
        _index = index;
    }

    /// <summary>Queues every mail whose attachments haven't been read yet. Safe to call repeatedly.</summary>
    public void Enqueue(IEnumerable<MailSummary> mails)
    {
        foreach (var mail in mails)
        {
            if (!mail.HasAttachments || _index.Contains(mail.EntryId) || !_queuedEntryIds.Add(mail.EntryId))
                continue;

            _pending.Enqueue(mail);
        }

        if (_pending.Count > 0)
            _ = RunAsync();
    }

    private async Task RunAsync()
    {
        if (_running)
            return;

        _running = true;
        try
        {
            while (_pending.Count > 0)
            {
                var mail = _pending.Dequeue();
                _queuedEntryIds.Remove(mail.EntryId);

                _index.Set(mail.EntryId, await IndexOneAsync(mail));
                Progressed?.Invoke();
            }

            _index.Save();
        }
        finally
        {
            _running = false;
        }
    }

    private async Task<string> IndexOneAsync(MailSummary mail)
    {
        List<AttachmentInfo> attachments;
        try
        {
            attachments = await _outlook.GetAttachmentsAsync(mail.StoreId, mail.EntryId);
        }
        catch (Exception)
        {
            return ""; // moved or deleted since we queued it
        }

        var text = new StringBuilder();
        foreach (var attachment in attachments)
        {
            // File names are searchable either way; only supported types are worth extracting.
            text.Append(attachment.FileName).Append(' ');
            if (!AttachmentTextExtractor.Supports(attachment.FileName))
                continue;

            string path;
            try
            {
                path = await _outlook.SaveAttachmentAsync(mail.StoreId, mail.EntryId, attachment.Index);
            }
            catch (Exception)
            {
                continue;
            }

            try
            {
                if (new FileInfo(path).Length <= AttachmentTextExtractor.MaxFileBytes)
                    text.Append(await Task.Run(() => AttachmentTextExtractor.Extract(path))).Append(' ');
            }
            catch (Exception)
            {
                // Best-effort: an unreadable attachment just isn't searchable by content.
            }
            finally
            {
                TryDelete(path);
            }
        }

        return text.ToString();
    }

    private static void TryDelete(string path)
    {
        try
        {
            File.Delete(path);
        }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }
}
