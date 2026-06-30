namespace HotPort.Infrastructure
{
    public interface IDialogService
    {
        bool TryOpenFile(string title, string filter, out string path, string? initialDirectory = null);
        bool TrySaveFile(string title, string filter, string? initialDirectory, string? fileName, out string path);
        bool TryOpenFolder(out string path);
        void ShowError(string message, string title);
        void ShowWarning(string message, string title);
    }
}
