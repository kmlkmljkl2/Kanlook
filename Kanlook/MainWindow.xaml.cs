using System.Windows;
using Kanlook.Services;
using Kanlook.ViewModels;

namespace Kanlook;

public partial class MainWindow : Window
{
    private readonly MainViewModel _viewModel;

    public MainWindow()
    {
        InitializeComponent();
        _viewModel = new MainViewModel(new OutlookService(), new BoardStateStore());
        DataContext = _viewModel;

        // Connecting to Outlook and reading the mailboxes happens after the window is up, so a
        // mailbox that takes a while to answer shows an empty sidebar rather than nothing at all.
        Loaded += async (_, _) => await _viewModel.InitializeAsync();
        Closed += (_, _) => _viewModel.Shutdown();
    }
}
