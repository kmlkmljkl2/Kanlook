using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace Kanlook.Behaviors;

/// <summary>
/// WPF's WebBrowser has no bindable "raw HTML string" property, so this attached
/// property hands the bound HTML to the control whenever it changes.
/// </summary>
public static class WebBrowserHtmlBehavior
{
    public static readonly DependencyProperty HtmlContentProperty = DependencyProperty.RegisterAttached(
        "HtmlContent",
        typeof(string),
        typeof(WebBrowserHtmlBehavior),
        new PropertyMetadata(null, OnHtmlContentChanged));

    public static string? GetHtmlContent(DependencyObject obj) => (string?)obj.GetValue(HtmlContentProperty);

    public static void SetHtmlContent(DependencyObject obj, string? value) => obj.SetValue(HtmlContentProperty, value);

    /// <summary>Emits the byte-order mark, which is what tells the browser how to read the rest.</summary>
    private static readonly UTF8Encoding Utf8 = new(encoderShouldEmitUTF8Identifier: true);

    /// <summary>
    /// The long form rather than HTML5's <c>&lt;meta charset&gt;</c>: the WebBrowser control runs in
    /// a legacy document mode, which only understands this one.
    /// </summary>
    private const string Utf8Meta = """<meta http-equiv="Content-Type" content="text/html; charset=utf-8">""";

    private const string EmptyDocument = "<html><body></body></html>";

    private static void OnHtmlContentChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is not WebBrowser browser)
            return;

        var html = e.NewValue as string;
        browser.NavigateToStream(ToUtf8Stream(string.IsNullOrEmpty(html) ? EmptyDocument : html));
    }

    /// <summary>
    /// Writes the document as UTF-8, and says so twice.
    ///
    /// NavigateToString would do the encoding part and nothing else: the bytes go out as UTF-8 with
    /// no mark and no declaration, so the browser falls back to whatever the mail itself claims -
    /// Outlook writes "us-ascii" or "iso-8859-1" into most of them - or, failing that, to the
    /// machine's ANSI codepage. Read back as Latin-1, a UTF-8 "ü" arrives as "Ã¼", which is why
    /// German mail came out mangled.
    ///
    /// The mark alone settles it under the encoding-sniffing rules, and the injected declaration is
    /// there because this control parses as an old version of IE, where "alone" is doing a lot of
    /// work. Ours goes in ahead of the mail's own so that whichever of the two the parser decides
    /// to believe, it believes this one.
    /// </summary>
    private static MemoryStream ToUtf8Stream(string html)
    {
        var preamble = Utf8.GetPreamble();
        var body = Utf8.GetBytes(WithUtf8Declaration(html));

        var stream = new MemoryStream(preamble.Length + body.Length);
        stream.Write(preamble);
        stream.Write(body);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// Puts our own charset declaration at the top of the document's head, or at the very top when
    /// it hasn't got one - the parser only sniffs the first stretch of the file either way. The
    /// mail's markup is otherwise left exactly as it came.
    /// </summary>
    private static string WithUtf8Declaration(string html)
    {
        var head = html.IndexOf("<head", StringComparison.OrdinalIgnoreCase);
        if (head < 0)
            return Utf8Meta + html;

        // Past the end of the <head ...> tag itself, which may carry attributes.
        var afterTag = html.IndexOf('>', head);
        return afterTag < 0
            ? Utf8Meta + html
            : html[..(afterTag + 1)] + Utf8Meta + html[(afterTag + 1)..];
    }
}
