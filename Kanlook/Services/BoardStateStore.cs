using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using Kanlook.Models;

namespace Kanlook.Services;

public sealed class BoardStateStore
{
    /// <summary>Where the state lives unless a caller asks for somewhere else.</summary>
    public static string DefaultFilePath { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Kanlook", "board-state.json");

    // Enums as names, so the file stays readable if anybody ever opens it.
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() },
    };

    private readonly AppState _state;
    private readonly string _filePath;

    /// <param name="filePath">
    /// Where to read and write the state. Defaults to <see cref="DefaultFilePath"/>; passing
    /// somewhere else keeps a throwaway board - a preview harness, say - out of the real file.
    /// </param>
    public BoardStateStore(string? filePath = null)
    {
        _filePath = filePath ?? DefaultFilePath;
        _state = Load(_filePath);
        Annotations = new MailAnnotationStore(_state.MailAnnotations, Save);
    }

    /// <summary>The user's notes and priority flags. Shared by every board and mail list.</summary>
    public MailAnnotationStore Annotations { get; }

    private static AppState Load(string filePath)
    {
        try
        {
            if (File.Exists(filePath))
            {
                var json = File.ReadAllText(filePath);
                var loaded = JsonSerializer.Deserialize<AppState>(json, JsonOptions);
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
        for (var i = 0; i < ColumnColors.DefaultColumns.Count; i++)
        {
            var (name, colorHex) = ColumnColors.DefaultColumns[i];
            fresh.Columns.Add(new KanbanColumnDefinition
            {
                Name = name,
                ColorHex = colorHex,
                Order = i,
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
            Directory.CreateDirectory(Path.GetDirectoryName(_filePath)!);
            var json = JsonSerializer.Serialize(_state, JsonOptions);
            File.WriteAllText(_filePath, json);
        }
        catch (IOException) { }
    }
}
