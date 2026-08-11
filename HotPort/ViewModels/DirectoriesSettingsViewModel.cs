using HotPort.Infrastructure;
using HotPort.Properties;
using System.Windows.Input;

namespace HotPort.ViewModels
{
    /// <summary>
    /// Backs the "Directories" section of the settings window: the default
    /// template directory and the code library directory.
    /// </summary>
    internal class DirectoriesSettingsViewModel : ObservableObject
    {
        private readonly IDialogService _dialogs;

        private string _templateDir;
        private string _codeLibDir;

        public DirectoriesSettingsViewModel(IDialogService dialogs)
        {
            _dialogs = dialogs;
            _templateDir = Settings.Default.TemplateDir ?? string.Empty;
            _codeLibDir = Settings.Default.CodeLibDir ?? string.Empty;

            BrowseTemplateDirCommand = new RelayCommand(BrowseTemplateDir);
            BrowseCodeLibDirCommand = new RelayCommand(BrowseCodeLibDir);
        }

        public string TemplateDir
        {
            get => _templateDir;
            set => SetProperty(ref _templateDir, value);
        }

        public string CodeLibDir
        {
            get => _codeLibDir;
            set => SetProperty(ref _codeLibDir, value);
        }

        public ICommand BrowseTemplateDirCommand { get; }
        public ICommand BrowseCodeLibDirCommand { get; }

        private void BrowseTemplateDir()
        {
            if (_dialogs.TryOpenFolder(out string path))
                TemplateDir = path;
        }

        private void BrowseCodeLibDir()
        {
            if (_dialogs.TryOpenFolder(out string path))
                CodeLibDir = path;
        }

        /// <summary>
        /// Writes the edited paths into <see cref="Settings"/>. Does not persist —
        /// the owning window saves once.
        /// </summary>
        public void Apply()
        {
            Settings.Default.TemplateDir = TemplateDir;
            Settings.Default.CodeLibDir = CodeLibDir;
        }
    }
}
