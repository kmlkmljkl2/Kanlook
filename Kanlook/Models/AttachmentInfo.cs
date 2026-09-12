namespace Kanlook.Models;

public sealed class AttachmentInfo
{
    public required string FileName { get; init; }

    /// <summary>Precomputed friendly size (e.g. "24 KB") so views don't need a size converter.</summary>
    public required string SizeDisplay { get; init; }

    /// <summary>1-based index into the mail's Outlook Attachments collection - used to re-open the attachment.</summary>
    public required int Index { get; init; }

    public static string FormatBytes(long bytes)
    {
        string[] units = ["B", "KB", "MB", "GB"];
        double size = bytes;
        var unitIndex = 0;
        while (size >= 1024 && unitIndex < units.Length - 1)
        {
            size /= 1024;
            unitIndex++;
        }

        return unitIndex == 0 ? $"{size:0} {units[unitIndex]}" : $"{size:0.#} {units[unitIndex]}";
    }
}
