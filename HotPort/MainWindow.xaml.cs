using HotPort.Infrastructure;
using HotPort.Properties;
using HotPort.ViewModels;
using System;
using System.ComponentModel;
using System.Reflection;
using System.Windows;

namespace HotPort
{
    public partial class MainWindow
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new MainWindowViewModel(new DialogService());

            this.Left = Settings.Default.WindowLeft;
            this.Top = Settings.Default.WindowTop;

            // Fix right-aligned menus on systems with RTL menu drop alignment
            var menuDropAlignmentField = typeof(SystemParameters).GetField("_menuDropAlignment", BindingFlags.NonPublic | BindingFlags.Static);
            Action setAlignmentValue = () =>
            {
                if (SystemParameters.MenuDropAlignment && menuDropAlignmentField != null)
                    menuDropAlignmentField.SetValue(null, false);
            };
            setAlignmentValue();
            SystemParameters.StaticPropertyChanged += (sender, e) => setAlignmentValue();
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            Settings.Default.WindowLeft = this.Left;
            Settings.Default.WindowTop = this.Top;
            Settings.Default.Save();
        }
    }
}
