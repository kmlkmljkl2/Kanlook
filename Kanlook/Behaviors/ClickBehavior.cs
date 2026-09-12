using System.Windows;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Media3D;

namespace Kanlook.Behaviors;

/// <summary>
/// Fires a command on a genuine click (press + release with minimal movement) instead of on
/// mouse-down, so starting a drag (press, then move) never triggers it - unlike Button (which
/// captures the mouse on press) or MouseBinding's "LeftClick" gesture (which matches on press).
/// </summary>
public static class ClickBehavior
{
    private const double MovementThreshold = 4.0;

    public static readonly DependencyProperty CommandProperty = DependencyProperty.RegisterAttached(
        "Command",
        typeof(ICommand),
        typeof(ClickBehavior),
        new PropertyMetadata(null, OnCommandChanged));

    public static ICommand? GetCommand(DependencyObject obj) => (ICommand?)obj.GetValue(CommandProperty);

    public static void SetCommand(DependencyObject obj, ICommand? value) => obj.SetValue(CommandProperty, value);

    private static readonly DependencyProperty MouseDownPositionProperty = DependencyProperty.RegisterAttached(
        "MouseDownPosition", typeof(Point?), typeof(ClickBehavior), new PropertyMetadata(null));

    private static void OnCommandChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is not UIElement element)
            return;

        element.PreviewMouseLeftButtonDown -= OnPreviewMouseLeftButtonDown;
        element.PreviewMouseLeftButtonUp -= OnPreviewMouseLeftButtonUp;

        if (e.NewValue is not null)
        {
            element.PreviewMouseLeftButtonDown += OnPreviewMouseLeftButtonDown;
            element.PreviewMouseLeftButtonUp += OnPreviewMouseLeftButtonUp;
        }
    }

    private static void OnPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is UIElement element)
            element.SetValue(MouseDownPositionProperty, e.GetPosition(element));
    }

    private static void OnPreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (sender is not UIElement element)
            return;

        var downPos = (Point?)element.GetValue(MouseDownPositionProperty);
        element.SetValue(MouseDownPositionProperty, null);
        if (downPos is null)
            return;

        // Buttons nested inside the element (e.g. a card's expander) own their clicks. Preview events
        // tunnel downwards, so without this the outer command would run before the button's does.
        if (OriginatesFromButton(e.OriginalSource, element))
            return;

        var upPos = e.GetPosition(element);
        var moved = (upPos - downPos.Value).Length;
        if (moved <= MovementThreshold)
            GetCommand(element)?.Execute(null);
    }

    private static bool OriginatesFromButton(object? originalSource, UIElement element)
    {
        var current = originalSource as DependencyObject;
        while (current is not null && !ReferenceEquals(current, element))
        {
            if (current is ButtonBase)
                return true;

            current = current is Visual or Visual3D
                ? VisualTreeHelper.GetParent(current)
                : LogicalTreeHelper.GetParent(current);
        }

        return false;
    }
}
