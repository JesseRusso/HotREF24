using Microsoft.Win32;
using Ookii.Dialogs.Wpf;

namespace HotPort.Services
{
    public class FileDialogService : IFileDialogService
    {
        public string? OpenFile(string title, string filter, string? initialDirectory = null)
        {
            var ofd = new OpenFileDialog
            {
                Title = title,
                Filter = filter
            };
            if (!string.IsNullOrEmpty(initialDirectory))
            {
                ofd.InitialDirectory = initialDirectory;
            }
            return ofd.ShowDialog() == true ? ofd.FileName : null;
        }

        public string? SaveFile(string title, string filter, string? initialDirectory = null, string? fileName = null)
        {
            var sfd = new SaveFileDialog
            {
                Title = title,
                Filter = filter,
                FileName = fileName
            };
            if (!string.IsNullOrEmpty(initialDirectory))
            {
                sfd.InitialDirectory = initialDirectory;
            }
            return sfd.ShowDialog() == true ? sfd.FileName : null;
        }

        public string? SelectFolder()
        {
            var fbd = new VistaFolderBrowserDialog();
            return fbd.ShowDialog() == true ? fbd.SelectedPath : null;
        }
    }
}
