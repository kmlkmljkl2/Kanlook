using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Kanlook.Models;

namespace Kanlook.ViewModels;

/// <summary>An attachment chip in the preview pane - click it to open the file.</summary>
public sealed partial class AttachmentViewModel : ObservableObject
{
    public string FileName { get; }
    public string SizeDisplay { get; }
    public string ToolTip => $"Open {FileName}";

    private readonly Action<AttachmentInfo> _onOpen;
    private readonly AttachmentInfo _info;

    public AttachmentViewModel(AttachmentInfo info, Action<AttachmentInfo> onOpen)
    {
        _info = info;
        _onOpen = onOpen;
        FileName = info.FileName;
        SizeDisplay = info.SizeDisplay;
    }

    [RelayCommand]
    private void Open() => _onOpen(_info);
}
