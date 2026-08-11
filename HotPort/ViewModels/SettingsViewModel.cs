using HotPort.Infrastructure;
using HotPort.Properties;

namespace HotPort.ViewModels
{
    /// <summary>
    /// Aggregates every settings section shown in the settings window and applies
    /// them together with a single <see cref="Save"/>.
    /// </summary>
    internal class SettingsViewModel : ObservableObject
    {
        public DirectoriesSettingsViewModel Directories { get; }
        public DoorsSettingsViewModel Doors { get; }

        public SettingsViewModel(IDialogService dialogs)
        {
            Directories = new DirectoriesSettingsViewModel(dialogs);
            Doors = new DoorsSettingsViewModel();
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
