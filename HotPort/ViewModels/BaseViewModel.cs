using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace HotPort.ViewModels
{
    /// <summary>
    /// Provides a base implementation of <see cref="INotifyPropertyChanged"/> for view models.
    /// </summary>
    public abstract class BaseViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;

        /// <summary>
        /// Raises the <see cref="PropertyChanged"/> event for the specified property.
        /// </summary>
        /// <param name="propertyName">Name of the property. This value is optional and
        /// will be provided automatically when invoked from compilers that support CallerMemberName.</param>
        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
