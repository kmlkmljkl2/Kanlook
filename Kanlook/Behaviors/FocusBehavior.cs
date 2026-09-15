using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace Kanlook.Behaviors;

/// <summary>
/// Moves the keyboard focus into an element the moment it becomes visible, so an editor a command
/// just revealed can be typed into without a click first.
/// </summary>
public static class FocusBehavior
{
    /// <summary>
    /// Focus the element, leaving a text box's caret at the end. For an editor whose existing text
    /// is being added to - a note - where selecting everything would risk wiping it.
    /// </summary>
    public static readonly DependencyProperty FocusWhenVisibleProperty = DependencyProperty.RegisterAttached(
        "FocusWhenVisible",
        typeof(bool),
        typeof(FocusBehavior),
        new PropertyMetadata(false, OnRequestChanged));

    public static bool GetFocusWhenVisible(DependencyObject obj) => (bool)obj.GetValue(FocusWhenVisibleProperty);

    public static void SetFocusWhenVisible(DependencyObject obj, bool value) =>
        obj.SetValue(FocusWhenVisibleProperty, value);

    /// <summary>
    /// Focus the element and select its text. For an editor whose content is normally replaced
    /// wholesale - a rename box - where typing should overwrite the old name.
    /// </summary>
    public static readonly DependencyProperty SelectAllWhenVisibleProperty = DependencyProperty.RegisterAttached(
        "SelectAllWhenVisible",
        typeof(bool),
        typeof(FocusBehavior),
        new PropertyMetadata(false, OnRequestChanged));

    public static bool GetSelectAllWhenVisible(DependencyObject obj) =>
        (bool)obj.GetValue(SelectAllWhenVisibleProperty);

    public static void SetSelectAllWhenVisible(DependencyObject obj, bool value) =>
        obj.SetValue(SelectAllWhenVisibleProperty, value);

    private static void OnRequestChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is not FrameworkElement element)
            return;

        element.IsVisibleChanged -= OnIsVisibleChanged;
        if (GetFocusWhenVisible(element) || GetSelectAllWhenVisible(element))
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
            if (element is not TextBox box)
                return;

            if (GetSelectAllWhenVisible(element))
                box.SelectAll();
            else
                box.CaretIndex = box.Text.Length;
        });
    }
}
