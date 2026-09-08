using System.Windows;
using System.Windows.Controls;

namespace Kanlook.Behaviors;

/// <summary>
/// WPF's WebBrowser has no bindable "raw HTML string" property, so this attached
/// property calls NavigateToString whenever the bound HTML content changes.
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

    private static void OnHtmlContentChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is not WebBrowser browser)
            return;

        var html = e.NewValue as string;
        browser.NavigateToString(string.IsNullOrEmpty(html) ? "<html><body></body></html>" : html);
    }
}
