using Kanlook.Models;

namespace Kanlook.Services;

/// <summary>
/// The replies the user sent, grouped by conversation, so a conversation tile can show what was
/// answered - the way Outlook's threaded view does. Sent mail is history only: it never becomes a
/// card of its own and never moves a conversation between columns.
///
/// Only the newest <see cref="MaxSentItems"/> sent mails are indexed, which is the same recency
/// window the board itself works in. Replies older than that simply don't show up in a tile.
/// </summary>
public sealed class SentMailIndex
{
    private const int MaxSentItems = 300;

    /// <summary>Full fetches per refresh. Spreads a burst of sent mail over several ticks.</summary>
    private const int MaxFetchesPerRefresh = 25;

    private static readonly TimeSpan MinRefreshInterval = TimeSpan.FromMinutes(1);

    private readonly IOutlookService _outlook;

    /// <summary>conversation key -&gt; sent mails, newest first.</summary>
    private readonly Dictionary<string, List<MailSummary>> _byConversation = new(StringComparer.Ordinal);
    private readonly Dictionary<string, MailSummary> _byEntryId = new(StringComparer.Ordinal);

    /// <summary>storeId -&gt; its Sent Items folder id. A null value means the store hasn't got one.</summary>
    private readonly Dictionary<string, string?> _folders = new(StringComparer.Ordinal);
    private readonly Dictionary<string, DateTime> _lastRefresh = new(StringComparer.Ordinal);

    public SentMailIndex(IOutlookService outlook) => _outlook = outlook;

    /// <summary>Raised after the index changed, so open boards can refresh their tiles.</summary>
    public event Action? Changed;

    public IReadOnlyList<MailSummary> ForConversation(string conversationKey) =>
        _byConversation.TryGetValue(conversationKey, out var sent) ? sent : [];

    /// <summary>
    /// First full read of a store's Sent Items. One table read, so it's cheap, but it still belongs
    /// behind the folder the user asked for. Does nothing once the store has been read.
    /// </summary>
    public async Task EnsureLoadedAsync(string storeId)
    {
        if (_folders.ContainsKey(storeId))
            return;

        var folderId = await TryAsync(() => _outlook.GetSentItemsFolderIdAsync(storeId));
        _folders[storeId] = folderId;
        if (folderId is null)
            return;

        _lastRefresh[storeId] = DateTime.Now;
        var summaries = await TryAsync(() => _outlook.GetMailSummariesAsync(storeId, folderId, MaxSentItems));
        if (summaries is not { Count: > 0 })
            return;

        foreach (var summary in summaries)
            Add(summary);

        Changed?.Invoke();
    }

    /// <summary>
    /// Picks up newly sent mail: one cheap table read to spot ids we haven't got, then a full fetch
    /// for just those. Light enough to call from the board's sync tick, and throttled on top of that.
    /// </summary>
    public async Task RefreshAsync(string storeId)
    {
        if (!_folders.TryGetValue(storeId, out var folderId))
        {
            await EnsureLoadedAsync(storeId);
            return;
        }

        if (folderId is null)
            return;

        if (_lastRefresh.TryGetValue(storeId, out var last) && DateTime.Now - last < MinRefreshInterval)
            return;

        _lastRefresh[storeId] = DateTime.Now;

        var state = await TryAsync(() => _outlook.GetFolderStateAsync(storeId, folderId, MaxSentItems));
        if (state is null)
            return;

        var changed = false;
        var fetches = 0;
        var live = new HashSet<string>(StringComparer.Ordinal);

        foreach (var item in state)
        {
            if (!item.IsMail)
                continue;

            live.Add(item.EntryId);

            if (_byEntryId.ContainsKey(item.EntryId) || fetches >= MaxFetchesPerRefresh)
                continue;

            fetches++;
            var summary = await TryAsync(() => _outlook.GetMailSummaryAsync(storeId, item.EntryId));
            if (summary is null)
                continue;

            Add(summary);
            changed = true;
        }

        // Only prune when the snapshot covered the whole folder. Past that the window no longer
        // reaches the oldest indexed mail, and "missing from the snapshot" stops meaning "deleted".
        if (state.Count < MaxSentItems)
        {
            foreach (var gone in _byEntryId.Values
                         .Where(s => s.StoreId == storeId && !live.Contains(s.EntryId))
                         .ToList())
            {
                Remove(gone);
                changed = true;
            }
        }

        if (changed)
            Changed?.Invoke();
    }

    private void Add(MailSummary summary)
    {
        if (!_byEntryId.TryAdd(summary.EntryId, summary))
            return;

        var key = summary.ConversationKey;
        if (!_byConversation.TryGetValue(key, out var list))
            _byConversation[key] = list = [];

        list.Add(summary);

        // Sent mail has no meaningful ReceivedTime, so order the history by when it was composed -
        // which is also what the board sorts cards on.
        list.Sort((a, b) => b.CreationTime.CompareTo(a.CreationTime));
    }

    private void Remove(MailSummary summary)
    {
        _byEntryId.Remove(summary.EntryId);

        if (!_byConversation.TryGetValue(summary.ConversationKey, out var list))
            return;

        list.RemoveAll(s => s.EntryId == summary.EntryId);
        if (list.Count == 0)
            _byConversation.Remove(summary.ConversationKey);
    }

    /// <summary>Outlook can refuse a folder or item at any moment; sent history isn't worth failing over.</summary>
    private static async Task<T?> TryAsync<T>(Func<Task<T>> action)
    {
        try
        {
            return await action();
        }
        catch (Exception)
        {
            return default;
        }
    }
}
