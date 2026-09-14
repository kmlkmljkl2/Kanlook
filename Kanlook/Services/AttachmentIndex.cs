using System.IO;
using System.Text.Json;

namespace Kanlook.Services;

/// <summary>
/// Searchable text of every attachment we've already read, keyed by mail entry id. Persisted so a
/// mail's attachments are only ever extracted and parsed once.
/// </summary>
public sealed class AttachmentIndex
{
    private static readonly string FilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "Kanlook", "attachment-index.json");

    private readonly Dictionary<string, string> _textByEntryId;
    private bool _dirty;

    public AttachmentIndex()
    {
        _textByEntryId = Load();
    }

    private static Dictionary<string, string> Load()
    {
        try
        {
            if (File.Exists(FilePath))
            {
                var loaded = JsonSerializer.Deserialize<Dictionary<string, string>>(File.ReadAllText(FilePath));
                if (loaded is not null)
                    return new Dictionary<string, string>(loaded, StringComparer.OrdinalIgnoreCase);
            }
        }
        catch (IOException) { }
        catch (JsonException) { }

        return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    }

    public bool Contains(string entryId) => _textByEntryId.ContainsKey(entryId);

    public string TextFor(string entryId) => _textByEntryId.GetValueOrDefault(entryId, "");

    public void Set(string entryId, string text)
    {
        _textByEntryId[entryId] = text;
        _dirty = true;
    }

    public void Save()
    {
        if (!_dirty)
            return;

        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
            File.WriteAllText(FilePath, JsonSerializer.Serialize(_textByEntryId));
            _dirty = false;
        }
        catch (IOException) { }
    }
}
