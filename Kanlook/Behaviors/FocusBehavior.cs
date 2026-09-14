using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace Kanlook.Behaviors;

/// <summary>
/// Moves the keyboard focus into an element the moment it becomes visible, so an editor a command
/// just revealed can be typed into without a click first. In a text box the caret lands at the end
/// rather than selecting everything, so reopening an existing note doesn't risk wiping it.
/// </summary>
public static class FocusBehavior
{
    public static readonly DependencyProperty FocusWhenVisibleProperty = DependencyProperty.RegisterAttached(
        "FocusWhenVisible",
        typeof(bool),
        typeof(FocusBehavior),
        new PropertyMetadata(false, OnFocusWhenVisibleChanged));

    public static bool GetFocusWhenVisible(DependencyObject obj) => (bool)obj.GetValue(FocusWhenVisibleProperty);

    public static void SetFocusWhenVisible(DependencyObject obj, bool value) =>
        obj.SetValue(FocusWhenVisibleProperty, value);

    private static void OnFocusWhenVisibleChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is not FrameworkElement element)
            return;

        element.IsVisibleChanged -= OnIsVisibleChanged;
        if (e.NewValue is true)
            element.IsVisibleChanged += OnIsVisibleChanged;
    }

    private static void OnIsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
    {
        if (e.NewValue is not true || sender is not FrameworkElement element)
            return;

        // The element is visible but not laid out yet, so focus has to wait for the render pass.
        element.Dispatcher.BeginInvoke(DispatcherPriority.Input, () =>
        {
            element.Focus();
            if (element is TextBox box)
                box.CaretIndex = box.Text.Length;
        });
    }
}
