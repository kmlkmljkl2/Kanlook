namespace Kanlook.Services;

public static class FolderKeyHelper
{
    public static string BuildKey(string storeId, string folderEntryId) => $"{storeId}|{folderEntryId}";
}
