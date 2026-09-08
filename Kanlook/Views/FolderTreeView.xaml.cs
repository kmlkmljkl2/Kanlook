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
        if (DataContext is MainViewModel vm)
            vm.SelectedFolder = e.NewValue as MailFolderNodeVm;
    }
}
