namespace HotPort.Services
{
    public interface IFileDialogService
    {
        string? OpenFile(string title, string filter, string? initialDirectory = null);
        string? SaveFile(string title, string filter, string? initialDirectory = null, string? fileName = null);
        string? SelectFolder();
    }
}
