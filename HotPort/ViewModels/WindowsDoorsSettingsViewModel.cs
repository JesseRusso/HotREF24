using HotPort.Infrastructure;
using HotPort.Properties;
using System;

namespace HotPort.ViewModels
{
    /// <summary>
    /// Backs the "Doors" section of the settings window. The user edits dimensions
    /// in inches; the backing <see cref="Settings"/> values are stored in metres.
    /// </summary>
    internal class WindowsDoorsSettingsViewModel : ObservableObject
    {
        private const double MetresPerInch = 0.0254;

        private double _frontDoorWidthInches;
        private double _frontDoorHeightInches;
        private bool _frontTransom;
        private double _garageDoorWidthInches;
        private double _garageDoorHeightInches;
        private int _maxExcelWindowsRow;
        public WindowsDoorsSettingsViewModel()
        {
            _frontDoorWidthInches = MetresToInches(Settings.Default.FrontDoorWidth);
            _frontDoorHeightInches = MetresToInches(Settings.Default.FrontDoorHeight);
            _frontTransom = Settings.Default.FrontTransom;
            _garageDoorWidthInches = MetresToInches(Settings.Default.GarageDoorWidth);
            _garageDoorHeightInches = MetresToInches(Settings.Default.GarageDoorHeight);
            _maxExcelWindowsRow = Settings.Default.MaxWindowRow;
        }
        public int MaxExcelWindowsRow
        {
            get => _maxExcelWindowsRow;
            set => SetProperty(ref _maxExcelWindowsRow, value);
        }

        public double FrontDoorWidthInches
        {
            get => _frontDoorWidthInches;
            set => SetProperty(ref _frontDoorWidthInches, value);
        }

        public double FrontDoorHeightInches
        {
            get => _frontDoorHeightInches;
            set => SetProperty(ref _frontDoorHeightInches, value);
        }

        public bool FrontTransom
        {
            get => _frontTransom;
            set => SetProperty(ref _frontTransom, value);
        }

        public double GarageDoorWidthInches
        {
            get => _garageDoorWidthInches;
            set => SetProperty(ref _garageDoorWidthInches, value);
        }

        public double GarageDoorHeightInches
        {
            get => _garageDoorHeightInches;
            set => SetProperty(ref _garageDoorHeightInches, value);
        }

        /// <summary>
        /// Converts the edited inch values back to metres and writes them into
        /// <see cref="Settings"/>. Does not persist — the owning window saves once.
        /// </summary>
        public void Apply()
        {
            Settings.Default.FrontDoorWidth = FrontDoorWidthInches * MetresPerInch;
            Settings.Default.FrontDoorHeight = FrontDoorHeightInches * MetresPerInch;
            Settings.Default.FrontTransom = FrontTransom;
            Settings.Default.GarageDoorWidth = GarageDoorWidthInches * MetresPerInch;
            Settings.Default.GarageDoorHeight = GarageDoorHeightInches * MetresPerInch;
            Settings.Default.MaxWindowRow = MaxExcelWindowsRow;
        }

        private static double MetresToInches(double metres) => Math.Round(metres / MetresPerInch, 2);
    }
}
