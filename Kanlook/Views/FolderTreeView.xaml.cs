using System.Windows;
using System.Windows.Controls;
using Kanlook.ViewModels;

namespace Kanlook.Views;

public partial class FolderTreeView : UserControl
{
    public FolderTreeView()
    {
        InitializeComponent();
    }

    private void TreeView_OnSelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
    {
        if (DataContext is not MainViewModel vm)
            return;

        // The "Loading…" row of a folder that hasn't been opened yet stands for nothing to show.
        if (e.NewValue is MailFolderNodeVm { IsPlaceholder: true })
            return;

        vm.SelectedFolder = e.NewValue as MailFolderNodeVm;
    }

    private void SettingsButton_Click(object sender, RoutedEventArgs e)
    {
        if (DataContext is not MainViewModel vm)
            return;

        new SettingsWindow
        {
            DataContext = vm.Settings,
            Owner = Window.GetWindow(this),
        }.ShowDialog();
    }
}
