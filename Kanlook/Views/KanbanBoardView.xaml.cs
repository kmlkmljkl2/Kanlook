using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Kanlook.ViewModels;

namespace Kanlook.Views;

public partial class KanbanBoardView : UserControl
{
    public KanbanBoardView()
    {
        InitializeComponent();
    }

    private void ColumnTitle_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ClickCount == 2 && sender is FrameworkElement { DataContext: KanbanColumnViewModel column })
            column.BeginRenameCommand.Execute(null);
    }

    private void NameEditBox_LostFocus(object sender, RoutedEventArgs e)
    {
        if (sender is FrameworkElement { DataContext: KanbanColumnViewModel column })
            column.CommitRenameCommand.Execute(null);
    }

    private void NameEditBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key != Key.Enter)
            return;

        if (sender is FrameworkElement { DataContext: KanbanColumnViewModel column })
            column.CommitRenameCommand.Execute(null);

        Keyboard.ClearFocus();
    }
}
