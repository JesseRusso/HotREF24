using System;
using System.ComponentModel;
using System.Reflection;
using System.Windows;
using HotPort.Properties;
using HotPort.ViewModels;

namespace HotPort
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        public MainWindow()
        {
            InitializeComponent();
            this.Left = Settings.Default.WindowLeft;
            this.Top = Settings.Default.WindowTop;

            var menuDropAlignmentField = typeof(SystemParameters).GetField("_menuDropAlignment", BindingFlags.NonPublic | BindingFlags.Static);
            Action setAlignmentValue = () =>
            {
                if (SystemParameters.MenuDropAlignment && menuDropAlignmentField != null) menuDropAlignmentField.SetValue(null, false);
            };
            setAlignmentValue();
            SystemParameters.StaticPropertyChanged += (sender, e) =>
            {
                setAlignmentValue();
            };
        }

        private void Window_Closing(object sender, CancelEventArgs e)
        {
            Settings.Default.WindowLeft = this.Left;
            Settings.Default.WindowTop = this.Top;

            if (DataContext is MainViewModel vm && vm.SaveSettingsCommand.CanExecute(null))
            {
                vm.SaveSettingsCommand.Execute(null);
            }
            else
            {
                Settings.Default.Save();
            }
        }
    }
}
