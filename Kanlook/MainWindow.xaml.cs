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
        _viewModel = new MainViewModel(new OutlookService());
        DataContext = _viewModel;
        Closed += (_, _) => _viewModel.Shutdown();
    }
}
