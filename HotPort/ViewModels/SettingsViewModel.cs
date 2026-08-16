using HotPort.Infrastructure;
using HotPort.Properties;
using System.Reflection.Metadata;

namespace HotPort.ViewModels
{
    /// <summary>
    /// Aggregates every settings section shown in the settings window and applies
    /// them together with a single <see cref="Save"/>.
    /// </summary>
    internal class SettingsViewModel : ObservableObject
    {
        public DirectoriesSettingsViewModel Directories { get; }
        public WindowsDoorsSettingsViewModel Doors { get; }

        public SettingsViewModel(IDialogService dialogs)
        {
            Directories = new DirectoriesSettingsViewModel(dialogs);
            Doors = new WindowsDoorsSettingsViewModel();
        }

        /// <summary>Applies every section to <see cref="Settings"/> and persists once.</summary>
        public void Save()
        {
            Directories.Apply();
            Doors.Apply();
            Settings.Default.Save();
        }
    }
}
