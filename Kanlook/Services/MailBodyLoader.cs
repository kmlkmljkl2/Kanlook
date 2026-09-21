using Kanlook.Models;

namespace Kanlook.Services;

/// <summary>
/// Fills in the one thing a folder's table read can't give us: the mail body, which the card shows
/// as its preview line and the search box matches on.
///
/// A folder therefore opens with its cards already there and their preview lines filling in behind,
/// newest first, one mail at a time at background priority - so a body fetch never delays opening
/// another folder or a mail. Bodies already read stay cached for the session, so going back to a
/// folder costs nothing.
/// </summary>
public sealed class MailBodyLoader
{
    /// <summary>
    /// How many bodies to keep. Each is capped at a few KB, so this is a handful of megabytes -
    /// enough to cover any realistic set of folders visited in one sitting.
    /// </summary>
    private const int MaxCachedBodies = 3_000;

    private readonly IOutlookService _outlook;
    private readonly Queue<MailSummary> _pending = new();
    private readonly HashSet<string> _queuedEntryIds = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>entry id -&gt; body text already read this session, oldest first for eviction.</summary>
    private readonly Dictionary<string, string> _cache = new(StringComparer.OrdinalIgnoreCase);
    private readonly Queue<string> _cacheOrder = new();

    private bool _running;

    public MailBodyLoader(IOutlookService outlook) => _outlook = outlook;

    /// <summary>Raised for each mail whose body has landed, so its card can restate its preview.</summary>
    public event Action<MailSummary>? Loaded;

    /// <summary>Raised whenever the backlog changes, so a board can show how far along it is.</summary>
    public event Action? Progressed;

    public int PendingCount => _pending.Count;

    /// <summary>
    /// Queues every mail whose body we haven't got, newest first. Safe to call repeatedly - the
    /// board calls it again whenever new mail arrives.
    /// </summary>
    public void Enqueue(IEnumerable<MailSummary> mails)
    {
        var queued = false;

        foreach (var mail in mails.Where(m => !m.BodyLoaded).OrderByDescending(m => m.CreationTime))
        {
            if (_cache.TryGetValue(mail.EntryId, out var cached))
            {
                mail.SetBodyText(cached);
                Loaded?.Invoke(mail);
                continue;
            }

            if (!_queuedEntryIds.Add(mail.EntryId))
                continue;

            _pending.Enqueue(mail);
            queued = true;
        }

        if (!queued)
            return;

        Progressed?.Invoke();
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

                string? body;
                try
                {
                    body = await _outlook.GetBodyTextAsync(mail.StoreId, mail.EntryId);
                }
                catch (Exception)
                {
                    body = null; // moved or deleted since we queued it - it just stays previewless
                }

                mail.SetBodyText(body);

                // The already-capped copy, not the raw body: a cache of whole mails would be the
                // one part of this that could grow without bound.
                Remember(mail.EntryId, mail.SearchBody);

                Loaded?.Invoke(mail);
                Progressed?.Invoke();
            }
        }
        finally
        {
            _running = false;
        }
    }

    private void Remember(string entryId, string body)
    {
        if (!_cache.TryAdd(entryId, body))
            return;

        _cacheOrder.Enqueue(entryId);
        while (_cacheOrder.Count > MaxCachedBodies)
            _cache.Remove(_cacheOrder.Dequeue());
    }
}
