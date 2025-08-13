using System.Windows;

namespace HotPort.Services
{
    public interface IMessageService
    {
        void ShowMessage(string message, string caption, MessageBoxButton buttons = MessageBoxButton.OK, MessageBoxImage icon = MessageBoxImage.None);
    }
}
