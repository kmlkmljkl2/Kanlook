namespace Kanlook.Models;

/// <summary>
/// What the reading pane shows: the message body, and the attachments worth offering beside it.
/// The two come together because deciding between them takes both - a picture the body displays
/// inline is not something to hand the user as a file.
/// </summary>
/// <param name="Html">
/// The body, with any <c>cid:</c> references to inline images pointed at extracted files so they
/// actually render. Null when the id no longer refers to a mail item.
/// </param>
/// <param name="Attachments">Everything the body doesn't already show.</param>
public sealed record MailContent(string? Html, List<AttachmentInfo> Attachments);
