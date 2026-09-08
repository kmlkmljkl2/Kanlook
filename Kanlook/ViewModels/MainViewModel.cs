using System.Collections.ObjectModel;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Kanlook.Services;

namespace Kanlook.ViewModels;

public sealed partial class MainViewModel : ObservableObject
{
    private readonly IOutlookService _outlook;
    private readonly BoardStateStore _boardStore = new();

    public ObservableCollection<MailFolderNodeVm> RootFolders { get; } = [];

    [ObservableProperty]
    private MailFolderNodeVm? _selectedFolder;

    [ObservableProperty]
    private object? _currentContent;

    [ObservableProperty]
    private PreviewPaneViewModel? _preview;

    [ObservableProperty]
    private GridLength _previewColumnWidth = new(0);

    [ObservableProperty]
    private string? _connectionErrorMessage;

    [ObservableProperty]
    private bool _isLoadingFolder;

    public MainViewModel(IOutlookService outlook)
    {
        _outlook = outlook;

        try
        {
            _outlook.Connect();
            foreach (var root in _outlook.BuildFolderTree())
                RootFolders.Add(new MailFolderNodeVm(root));
        }
        catch (Exception ex)
        {
            ConnectionErrorMessage =
                $"Couldn't connect to Outlook. Make sure the classic desktop Outlook app is installed and try again. ({ex.Message})";
        }
    }

    partial void OnSelectedFolderChanged(MailFolderNodeVm? value)
    {
        if (value is null)
            return;

        IsLoadingFolder = true;
        try
        {
            var mails = _outlook.GetMailSummaries(value.StoreId, value.EntryId);

            CurrentContent = value.IsSharedMailbox
                ? new SharedFolderViewModel(value.Name, mails, ShowPreview)
                : new KanbanBoardViewModel(
                    value.Name,
                    FolderKeyHelper.BuildKey(value.StoreId, value.EntryId),
                    mails,
                    _boardStore,
                    ShowPreview);
        }
        catch (Exception ex)
        {
            ConnectionErrorMessage = $"Couldn't load folder '{value.Name}': {ex.Message}";
        }
        finally
        {
            IsLoadingFolder = false;
        }
    }

    private void ShowPreview(MailCardViewModel card)
    {
        Preview = new PreviewPaneViewModel(card.Summary, _outlook, ClosePreview);
        if (PreviewColumnWidth.Value <= 0)
            PreviewColumnWidth = new GridLength(420);
    }

    [RelayCommand]
    private void ClosePreview()
    {
        Preview = null;
        PreviewColumnWidth = new GridLength(0);
    }

    public void Shutdown() => _outlook.Dispose();
}
