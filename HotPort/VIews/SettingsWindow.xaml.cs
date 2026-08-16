using HotPort.ViewModels;
using System.Windows;
using HotPort.Properties;
using System;
using HotPort.Infrastructure;

namespace HotPort.Views
{
    public partial class SettingsWindow
    {
        public SettingsWindow()
        {
            InitializeComponent();
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            ((SettingsViewModel)DataContext).Save();
            DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            // Edits live only in the section view-models until Save, so cancelling
            // simply discards them — Settings is left untouched.
            DialogResult = false;
        }
    }
}
