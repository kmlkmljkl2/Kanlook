using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using Kanlook.Models;

namespace Kanlook.ViewModels;

public sealed partial class SharedFolderViewModel : ObservableObject
{
    public string FolderName { get; }
    public ObservableCollection<MailCardViewModel> Cards { get; } = [];

    public SharedFolderViewModel(string folderName, List<MailSummary> mails, Action<MailCardViewModel> onCardSelected)
    {
        FolderName = folderName;
        foreach (var mail in mails)
            Cards.Add(new MailCardViewModel(mail, onCardSelected));
    }
}
