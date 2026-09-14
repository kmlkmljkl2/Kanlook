using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;

namespace Kanlook.Services;

/// <summary>
/// Pulls searchable text out of an attachment file. Everything here runs off the UI thread, so it
/// must not touch Outlook COM - it only ever sees a file that was already extracted to disk.
/// </summary>
public static partial class AttachmentTextExtractor
{
    /// <summary>
    /// Per-attachment cap (~25 pages of prose). Keeps the on-disk index and the per-search scan
    /// bounded no matter how large the attachment is.
    /// </summary>
    private const int MaxChars = 60_000;

    /// <summary>Files larger than this are skipped - extraction cost isn't worth it.</summary>
    public const long MaxFileBytes = 32L * 1024 * 1024;

    private static readonly HashSet<string> PlainTextExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".txt", ".csv", ".log", ".md", ".json", ".xml", ".htm", ".html", ".ini", ".yml", ".yaml", ".rtf",
    };

    /// <summary>Office Open XML containers - a zip of XML parts.</summary>
    private static readonly HashSet<string> OpenXmlExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".docx", ".docm", ".xlsx", ".xlsm", ".pptx", ".pptm",
    };

    public static bool Supports(string fileName)
    {
        var extension = Path.GetExtension(fileName);
        return PlainTextExtensions.Contains(extension) ||
               OpenXmlExtensions.Contains(extension) ||
               extension.Equals(".pdf", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Extracted text, or an empty string when the file can't be read or holds no text.</summary>
    public static string Extract(string filePath)
    {
        try
        {
            var extension = Path.GetExtension(filePath);

            if (PlainTextExtensions.Contains(extension))
                return Cap(StripMarkup(File.ReadAllText(filePath)));

            if (OpenXmlExtensions.Contains(extension))
                return Cap(ExtractOpenXml(filePath));

            if (extension.Equals(".pdf", StringComparison.OrdinalIgnoreCase))
                return Cap(ExtractPdf(filePath));
        }
        catch (Exception)
        {
            // A corrupt, encrypted or unexpected file shouldn't break indexing of the rest.
        }

        return "";
    }

    private static string ExtractPdf(string filePath)
    {
        using var document = PdfDocument.Open(filePath);
        var text = new StringBuilder();

        foreach (var page in document.GetPages())
        {
            text.Append(page.Text).Append(' ');
            if (text.Length >= MaxChars)
                break;
        }

        return text.ToString();
    }

    /// <summary>
    /// Reads a Word / Excel / PowerPoint package. Only the parts that actually hold words are
    /// opened - styles, themes and settings would otherwise fill the character budget with noise.
    /// </summary>
    private static string ExtractOpenXml(string filePath)
    {
        using var archive = ZipFile.OpenRead(filePath);
        var text = new StringBuilder();

        foreach (var part in archive.Entries.Where(e => HoldsText(e.FullName)).OrderBy(e => e.FullName))
        {
            using var reader = new StreamReader(part.Open());
            AppendTextNodes(reader.ReadToEnd(), text);
            if (text.Length >= MaxChars)
                break;
        }

        return text.ToString();
    }

    /// <summary>The parts of an Office package that carry authored text rather than formatting.</summary>
    private static bool HoldsText(string partName)
    {
        if (!partName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            return false;

        // Word: the body, plus anything running alongside it (headers, footers, notes, comments).
        if (partName.StartsWith("word/", StringComparison.OrdinalIgnoreCase))
        {
            var name = partName["word/".Length..];
            return name.StartsWith("document", StringComparison.OrdinalIgnoreCase) ||
                   name.StartsWith("header", StringComparison.OrdinalIgnoreCase) ||
                   name.StartsWith("footer", StringComparison.OrdinalIgnoreCase) ||
                   name.StartsWith("footnotes", StringComparison.OrdinalIgnoreCase) ||
                   name.StartsWith("endnotes", StringComparison.OrdinalIgnoreCase) ||
                   name.StartsWith("comments", StringComparison.OrdinalIgnoreCase);
        }

        // Excel: cell text lives in the shared string table, except for inline strings on the sheet.
        if (partName.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
        {
            return partName.Equals("xl/sharedStrings.xml", StringComparison.OrdinalIgnoreCase) ||
                   partName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) ||
                   partName.StartsWith("xl/comments", StringComparison.OrdinalIgnoreCase);
        }

        // PowerPoint: slide bodies and their speaker notes.
        return partName.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase) ||
               partName.StartsWith("ppt/notesSlides/notesSlide", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Appends the text runs of an Office XML part. Reading only the text elements matters: Word
    /// splits a single word across runs whenever formatting or spell-check state changes, so
    /// stripping tags wholesale would turn "Rechnung" into "Rech nung" and lose the search hit.
    /// Runs are joined as-is and only paragraph-level elements introduce a space.
    /// </summary>
    private static void AppendTextNodes(string xml, StringBuilder text)
    {
        foreach (var match in TextNodeRegex().EnumerateMatches(xml))
        {
            var element = xml.AsSpan(match.Index, match.Length);

            // A paragraph/row/cell boundary - the words either side are separate.
            if (element[1] == '/' || element[^2] == '/')
            {
                Append(text, " ");
                continue;
            }

            var close = element.IndexOf('>');
            var inner = element[(close + 1)..element.LastIndexOf("</", StringComparison.Ordinal)];
            Append(text, DecodeEntities(inner.ToString()));

            if (text.Length >= MaxChars)
                return;
        }

        Append(text, " ");
    }

    private static void Append(StringBuilder text, string value)
    {
        // Collapse the runs of separators that empty paragraphs and self-closing breaks produce.
        if (value == " " && (text.Length == 0 || text[^1] == ' '))
            return;

        text.Append(value);
    }

    private static string DecodeEntities(string value) => value.Contains('&')
        ? value.Replace("&lt;", "<").Replace("&gt;", ">").Replace("&quot;", "\"")
            .Replace("&apos;", "'").Replace("&amp;", "&")
        : value;

    /// <summary>Drops tags and collapses whitespace, so markup can be searched as words.</summary>
    private static string StripMarkup(string content) =>
        WhitespaceRegex().Replace(TagRegex().Replace(content, " "), " ").Trim();

    private static string Cap(string text) => text.Length > MaxChars ? text[..MaxChars] : text;

    [GeneratedRegex(@"<[^>]{0,4000}>", RegexOptions.Singleline)]
    private static partial Regex TagRegex();

    /// <summary>
    /// Office text elements (Word <c>w:t</c>, drawing/PowerPoint <c>a:t</c>, Excel <c>t</c>) plus the
    /// closing and self-closing elements that end a paragraph, table cell or shared string. Opening
    /// tags are deliberately not matched - only these three shapes reach <see cref="AppendTextNodes"/>.
    /// </summary>
    [GeneratedRegex("""
        </(?:w:p|w:tc|w:tr|a:p|si|is|c|row)>
        |<(?:w:br|w:cr|w:tab|a:br)(?:\s[^>]*)?/>
        |<(?:w:t|a:t|t)(?:\s[^>]*)?>[^<]*</(?:w:t|a:t|t)>
        """, RegexOptions.IgnorePatternWhitespace)]
    private static partial Regex TextNodeRegex();

    [GeneratedRegex(@"\s+")]
    private static partial Regex WhitespaceRegex();
}
