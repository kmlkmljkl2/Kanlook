using System.IO;
using System.Text.Json;
using Kanlook.Models;

namespace Kanlook.Services;

public sealed class BoardStateStore
{
    private static readonly string FilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Kanlook", "board-state.json");

    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    private readonly AppState _state;

    public BoardStateStore()
    {
        _state = Load();
    }

    private static AppState Load()
    {
        try
        {
            if (File.Exists(FilePath))
            {
                var json = File.ReadAllText(FilePath);
                var loaded = JsonSerializer.Deserialize<AppState>(json);
                if (loaded is not null)
                    return loaded;
            }
        }
        catch (IOException) { }
        catch (JsonException) { }

        return new AppState();
    }

    public AppSettings Settings => _state.Settings;

    public BoardState GetOrCreateBoard(string folderKey)
    {
        if (_state.Boards.TryGetValue(folderKey, out var existing))
            return existing;

        var fresh = new BoardState();
        var defaults = new[] { ("New", "#5B8DEF"), ("In Progress", "#F2A93B"), ("Waiting", "#B87CE0"), ("Done", "#4CB782") };
        for (var i = 0; i < defaults.Length; i++)
        {
            fresh.Columns.Add(new KanbanColumnDefinition
            {
                Name = defaults[i].Item1,
                ColorHex = defaults[i].Item2,
                Order = i
            });
        }

        _state.Boards[folderKey] = fresh;
        return fresh;
    }

    public List<string>? GetFolderOrder(string parentKey) =>
        _state.FolderOrder.TryGetValue(parentKey, out var order) ? order : null;

    public void SetFolderOrder(string parentKey, List<string> orderedEntryIds)
    {
        _state.FolderOrder[parentKey] = orderedEntryIds;
        Save();
    }

    public void Save()
    {
        try
        {
            var dir = Path.GetDirectoryName(FilePath)!;
            Directory.CreateDirectory(dir);
            var json = JsonSerializer.Serialize(_state, JsonOptions);
            File.WriteAllText(FilePath, json);
        }
        catch (IOException) { }
    }
}
