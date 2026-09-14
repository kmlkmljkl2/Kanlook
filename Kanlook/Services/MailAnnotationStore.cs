using Kanlook.Models;

namespace Kanlook.Services;

/// <summary>
/// The user's notes and priority flags, keyed by Outlook entry id and persisted with the rest of the
/// app state. Writes save straight away: a note is typed once and has to survive a crash.
/// </summary>
public sealed class MailAnnotationStore
{
    private readonly Dictionary<string, MailAnnotation> _annotations;
    private readonly Action _save;

    public MailAnnotationStore(Dictionary<string, MailAnnotation> annotations, Action save)
    {
        _annotations = annotations;
        _save = save;
    }

    public string CommentFor(string entryId) =>
        _annotations.TryGetValue(entryId, out var annotation) ? annotation.Comment : "";

    public bool IsHighPriority(string entryId) =>
        _annotations.TryGetValue(entryId, out var annotation) && annotation.IsHighPriority;

    /// <summary>Returns true when the note actually changed, so callers can skip a needless repaint.</summary>
    public bool SetComment(string entryId, string comment) =>
        Saving(Write(entryId, comment.Trim(), a => a.Comment, (a, value) => a.Comment = value));

    /// <summary>
    /// Flags or clears a whole tile's messages at once. One save for the lot - a conversation of a
    /// dozen mails shouldn't rewrite the state file a dozen times.
    /// </summary>
    public bool SetHighPriority(IEnumerable<string> entryIds, bool isHighPriority)
    {
        var changed = false;
        foreach (var entryId in entryIds)
            changed |= Write(entryId, isHighPriority, a => a.IsHighPriority, (a, value) => a.IsHighPriority = value);

        return Saving(changed);
    }

    /// <summary>
    /// Drops annotations of mails that are gone. Their entry ids can never come back - Outlook issues
    /// a new one when an item is moved or deleted - so keeping them would only grow the state file.
    /// </summary>
    public void Forget(IEnumerable<string> entryIds)
    {
        var changed = false;
        foreach (var entryId in entryIds)
            changed |= _annotations.Remove(entryId);

        Saving(changed);
    }

    private bool Saving(bool changed)
    {
        if (changed)
            _save();

        return changed;
    }

    /// <summary>Applies one field, creating or dropping the entry as needed. True when it changed.</summary>
    private bool Write<T>(string entryId, T value, Func<MailAnnotation, T> read, Action<MailAnnotation, T> write)
    {
        if (!_annotations.TryGetValue(entryId, out var annotation))
        {
            // Nothing stored and nothing to store - don't create an entry just to hold the default.
            if (EqualityComparer<T>.Default.Equals(value, read(new MailAnnotation())))
                return false;

            annotation = new MailAnnotation();
            _annotations[entryId] = annotation;
        }
        else if (EqualityComparer<T>.Default.Equals(read(annotation), value))
        {
            return false;
        }

        write(annotation, value);

        if (annotation.IsEmpty)
            _annotations.Remove(entryId);

        return true;
    }
}
