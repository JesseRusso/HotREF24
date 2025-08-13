using System.Windows;

namespace HotPort.Services
{
    public class MessageService : IMessageService
    {
        public void ShowMessage(string message, string caption, MessageBoxButton buttons = MessageBoxButton.OK, MessageBoxImage icon = MessageBoxImage.None)
        {
            MessageBox.Show(message, caption, buttons, icon);
        }
    }
}
