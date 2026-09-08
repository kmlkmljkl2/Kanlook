namespace Kanlook.Models;

public sealed class KanbanColumnDefinition
{
    public string Id { get; init; } = Guid.NewGuid().ToString("N");
    public string Name { get; set; } = "New column";
    public int Order { get; set; }
    public string ColorHex { get; set; } = "#5B8DEF";
}

public sealed class BoardState
{
    public List<KanbanColumnDefinition> Columns { get; init; } = [];

    /// <summary>entryId -&gt; column id</summary>
    public Dictionary<string, string> CardAssignments { get; init; } = [];
}

public sealed class AppState
{
    /// <summary>folder key ("storeId|folderEntryId") -&gt; board state</summary>
    public Dictionary<string, BoardState> Boards { get; init; } = [];
}
