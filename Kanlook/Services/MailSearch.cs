using Kanlook.Models;

namespace Kanlook.Services;

/// <summary>Matches mails against the board's search box.</summary>
public sealed class MailSearch
{
    private readonly AttachmentIndex _attachments;

    public MailSearch(AttachmentIndex attachments)
    {
        _attachments = attachments;
    }

    /// <summary>
    /// True when the mail matches every whitespace-separated term, each of which may hit any field:
    /// subject, sender, body, categories, or an attachment's name or contents.
    /// </summary>
    public bool Matches(MailSummary mail, string[] terms)
    {
        if (terms.Length == 0)
            return true;

        var haystack = new[]
        {
            mail.Subject,
            mail.SenderName,
            mail.SenderEmail,
            mail.ToNames,
            mail.SearchBody,
            mail.Categories,
            _attachments.TextFor(mail.EntryId),
        };

        return terms.All(term => haystack.Any(field => field.Contains(term, StringComparison.OrdinalIgnoreCase)));
    }

    public static string[] ParseTerms(string? searchText) =>
        string.IsNullOrWhiteSpace(searchText)
            ? []
            : searchText.Split(' ', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
}
