using HotPort.ViewModels;
using HotPort.Views;
using Microsoft.Win32;
using Ookii.Dialogs.Wpf;
using System.Windows;

namespace HotPort.Infrastructure
{
    public class DialogService : IDialogService
    {
        public bool TryOpenFile(string title, string filter, out string path, string? initialDirectory = null)
        {
            var ofd = new OpenFileDialog { Title = title, Filter = filter };
            if (initialDirectory != null) ofd.InitialDirectory = initialDirectory;

            if (ofd.ShowDialog() == true)
            {
                path = ofd.FileName;
                return true;
            }
            path = string.Empty;
            return false;
        }

        public bool TrySaveFile(string title, string filter, string? initialDirectory, string? fileName, out string path)
        {
            var sfd = new SaveFileDialog
            {
                Title = title,
                Filter = filter,
                InitialDirectory = initialDirectory,
                FileName = fileName,
            };

            if (sfd.ShowDialog() == true)
            {
                path = sfd.FileName;
                return true;
            }
            path = string.Empty;
            return false;
        }

        public bool TryOpenFolder(out string path)
        {
            var fbd = new VistaFolderBrowserDialog();
            if (fbd.ShowDialog() == true)
            {
                path = fbd.SelectedPath;
                return true;
            }
            path = string.Empty;
            return false;
        }

        public void ShowError(string message, string title) =>
            MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Error);

        public void ShowWarning(string message, string title) =>
            MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Warning);

        public void ShowSettings()
        {
            var window = new SettingsWindow
            {
                Owner = Application.Current?.MainWindow,
                DataContext = new SettingsViewModel(this)
            };
            window.ShowDialog();
        }
    }
}
