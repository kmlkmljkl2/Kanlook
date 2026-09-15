namespace Kanlook.Models;

public sealed class KanbanColumnDefinition
{
    public string Id { get; init; } = Guid.NewGuid().ToString("N");
    public string Name { get; set; } = "New column";
    public int Order { get; set; }
    public string ColorHex { get; set; } = ColumnColors.Blue;

    /// <summary>
    /// Mail parked here is waiting on somebody else. When a reply to one of its conversations comes
    /// in, the whole conversation goes back to the default column - it needs attention again.
    /// </summary>
    public bool WaitsForReply { get; set; }

    /// <summary>
    /// Days after which a waiting conversation returns to the default column even without a reply.
    /// Null turns the timer off.
    /// </summary>
    public int? ReturnAfterDays { get; set; }

    /// <summary>
    /// Time of day the return falls due, as "HH:mm". Null means the same time of day the mail was
    /// parked, so "2 days" is exactly 48 hours.
    /// </summary>
    public string? ReturnAtTime { get; set; }
}

public sealed class BoardState
{
    public List<KanbanColumnDefinition> Columns { get; init; } = [];

    /// <summary>entryId -&gt; column id</summary>
    public Dictionary<string, string> CardAssignments { get; init; } = [];

    /// <summary>
    /// entryId -&gt; when the mail was parked in a column that waits for a reply. The return timer
    /// counts from here, so it survives restarts.
    /// </summary>
    public Dictionary<string, DateTime> WaitingSince { get; init; } = [];

    /// <summary>Column that unassigned/newly-arrived mail lands in. Falls back to the first column when null.</summary>
    public string? DefaultColumnId { get; set; }
}

public sealed class AppState
{
    /// <summary>folder key ("storeId|folderEntryId") -&gt; board state</summary>
    public Dictionary<string, BoardState> Boards { get; init; } = [];

    /// <summary>parent folder key (or the root sentinel) -&gt; ordered child EntryIds</summary>
    public Dictionary<string, List<string>> FolderOrder { get; init; } = [];

    /// <summary>
    /// entryId -&gt; the user's note and priority flag. Global rather than per board: an annotation
    /// belongs to the mail, so it reads the same wherever the mail is shown.
    /// </summary>
    public Dictionary<string, MailAnnotation> MailAnnotations { get; init; } = [];

    public AppSettings Settings { get; init; } = new();
}
