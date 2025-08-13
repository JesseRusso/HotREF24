using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using System.Xml.Linq;
using HotPort.Properties;
using Microsoft.Win32;
using Ookii.Dialogs.Wpf;

namespace HotPort.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private XDocument? propHouse;
        private XDocument? newHouse;
        private string? templatePath;
        private string? excelFilePath;
        private string? proposedAddress;
        private string? directoryString;
        private XElement[]? profiles;
        private XElement? profile;
        private string worksheetText = "No worksheet selected";
        private string templateText = "No template selected";
        private string proposedFileText = "No file selected";
        private readonly ObservableCollection<string> zones;
        private int selectedZoneIndex;

        public string? TemplatePath { get => templatePath; set { templatePath = value; OnPropertyChanged(); } }
        public string? ExcelFilePath { get => excelFilePath; set { excelFilePath = value; OnPropertyChanged(); } }
        public string? ProposedAddress { get => proposedAddress; set { proposedAddress = value; OnPropertyChanged(); } }
        public string? DirectoryString { get => directoryString; set { directoryString = value; OnPropertyChanged(); } }
        public XElement[]? Profiles { get => profiles; set { profiles = value; OnPropertyChanged(); } }
        public XElement? Profile { get => profile; set { profile = value; OnPropertyChanged(); } }

        public string WorksheetText { get => worksheetText; set { worksheetText = value; OnPropertyChanged(); } }
        public string TemplateText { get => templateText; set { templateText = value; OnPropertyChanged(); } }
        public string ProposedFileText { get => proposedFileText; set { proposedFileText = value; OnPropertyChanged(); } }

        public ObservableCollection<string> Zones => zones;
        public int SelectedZoneIndex { get => selectedZoneIndex; set { selectedZoneIndex = value; OnPropertyChanged(); } }

        public ICommand SelectWorksheetCommand { get; }
        public ICommand SelectTemplateCommand { get; }
        public ICommand CreateProposedFileCommand { get; }
        public ICommand CreateRefCommand { get; }
        public ICommand SelectProposedFileCommand { get; }
        public ICommand SetDefaultDirectoryCommand { get; }

        public MainViewModel()
        {
            XDocument values = new XDocument(XDocument.Load(@".\\ReferenceProfiles.xml"));
            Profiles = (from el in values.Descendants("Zone")
                        select el).ToArray();
            zones = new ObservableCollection<string>(Profiles.Select(z => z.Attribute("name")?.Value ?? string.Empty));
            selectedZoneIndex = 0;

            SelectWorksheetCommand = new RelayCommand(_ => SelectWorksheet());
            SelectTemplateCommand = new RelayCommand(_ => SelectTemplate());
            CreateProposedFileCommand = new RelayCommand(_ => CreateProposedFile());
            CreateRefCommand = new RelayCommand(_ => CreateRef());
            SelectProposedFileCommand = new RelayCommand(_ => SelectProposedFile());
            SetDefaultDirectoryCommand = new RelayCommand(_ => SetDefaultDirectory());
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string? name = null) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        private void SelectWorksheet()
        {
            var ofd = new OpenFileDialog
            {
                Title = "Select worksheet",
                Filter = "Excel Files (*.xlsm) | *.xlsm"
            };

            if (ofd.ShowDialog() == true)
            {
                ExcelFilePath = ofd.FileName;
                WorksheetText = SplitAddress(ofd.FileName);
                ProposedAddress = WorksheetText;
            }
        }

        private void SelectProposedFile()
        {
            var ofd = new OpenFileDialog
            {
                Title = "Select Proposed File",
                Filter = "House Files (*.h2k)|*.h2k"
            };

            if (ofd.ShowDialog() == true)
            {
                propHouse = XDocument.Load(ofd.FileName);
                DirectoryString = Path.GetDirectoryName(ofd.FileName);
                ProposedAddress = SplitAddress(Path.GetFileName(ofd.FileName));
                ProposedFileText = ProposedAddress;
            }
        }

        private void SelectTemplate()
        {
            var ofd = new OpenFileDialog
            {
                Title = "Select HOT2000 builder template",
                Filter = "House Files(*.h2k) | *.h2k",
                InitialDirectory = Settings.Default.TemplateDir,
            };
            if (ofd.ShowDialog() == true)
            {
                TemplatePath = ofd.FileName;
                TemplateText = ofd.SafeFileName.Split('.').First();
            }
        }

        private void CreateRef()
        {
            if (propHouse == null)
            {
                MessageBox.Show("You must select a proposed file first.", "No file selected", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            else if (ExcelFilePath == null)
            {
                MessageBox.Show("You must select an Excel file first.", "No file selected", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            else
            {
                Mouse.OverrideCursor = Cursors.Wait;
                SetProfile();
                CreateRef cr = new CreateRef(propHouse, ExcelFilePath, Profile);
                cr.FindID(propHouse);
                propHouse = cr.Remover(propHouse);
                propHouse = cr.AddCode(propHouse);
                propHouse = cr.RChanger(propHouse);
                propHouse = cr.HeatingCooling(propHouse);
                propHouse = cr.ChangeACH(propHouse);
                propHouse = cr.AddFan(propHouse);
                propHouse = cr.Doors(propHouse);
                propHouse = cr.Windows(propHouse);
                propHouse = cr.HotWater(propHouse);
                Mouse.OverrideCursor = null;

                MessageBox.Show("Please save and check results", "REF changes made");
                SaveFileDialog sfd = new SaveFileDialog
                {
                    Filter = "House File|*.h2k",
                    DefaultExt = "h2k",
                    InitialDirectory = DirectoryString,
                    FileName = $"{ProposedAddress}-REFERENCE",
                };

                if (sfd.ShowDialog() == true)
                {
                    propHouse.Save(sfd.FileName);
                }
            }
        }

        private void CreateProposedFile()
        {
            if (TemplatePath == null)
            {
                MessageBox.Show("Select a builder template to modify",
                    "No builder template selected",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (ExcelFilePath == null)
            {
                MessageBox.Show("Select a spreadsheet to copy from.",
                    "No spreadsheet selected",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            Mouse.OverrideCursor = Cursors.Wait;
            XDocument template = new XDocument(XDocument.Load(TemplatePath));
            CreateProp cp = new CreateProp(ExcelFilePath, template);
            CreateProp.ChangeAddress(ProposedAddress);

            try { cp.ChangeEquipment(); }
            catch
            {
                Mouse.OverrideCursor = null;
                MessageBox.Show("There was an error. Check furnace and HRV values in excel.",
                    "Something went wrong",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            try { cp.CheckAC(); }
            catch
            {
                Mouse.OverrideCursor = null;
                MessageBox.Show("Unexpected value in A/C selection.",
                    "Something went wrong",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            try { cp.ChangeSpecs(); }
            catch
            {
                Mouse.OverrideCursor = null;
                MessageBox.Show("There was an error. Check for typos in intersections/corners and volume/highest ceiling.",
                    "Something went wrong",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            try { cp.ChangeWalls(); }
            catch
            {
                Mouse.OverrideCursor = null;
                MessageBox.Show("There was an error. Check for typos in above grade walls and check that the H2K template has all required elements.",
                    "Something went wrong",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            try { cp.CheckCeilings(); }
            catch
            {
                Mouse.OverrideCursor = null;
                MessageBox.Show("There was an error retrieving ceiling data from the template. Check that the H2K template has the required ceilings.",
                    "Something went wrong",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            try { cp.ChangeFloors(); }
            catch
            {
                Mouse.OverrideCursor = null;
                MessageBox.Show("Error either retrieving floor R values from template, or adding garage floor. Template should have 2 exposed floors with 'cant' and 'garage' in their names",
                    "Something went wrong",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            try { cp.ExtraFloors(); }
            catch
            {
                Mouse.OverrideCursor = null;
                MessageBox.Show("Unexpected value while adding floors. Have a typo in the EXPOSED FLOORS section?",
                    "Something went wrong",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            try { cp.ExtraCeilings(); }
            catch
            {
                Mouse.OverrideCursor = null;
                MessageBox.Show("Unexpected value while adding ceilings. Have a typo in the FLAT CEILINGS section?",
                    "Something went wrong",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            try { cp.CheckVaults(); }
            catch
            {
                Mouse.OverrideCursor = null;
                MessageBox.Show("Unexpected value while adding vaults. Have a typo in the VAULTS section?",
                    "Something went wrong",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            try { cp.ExtraWalls(); }
            catch
            {
                Mouse.OverrideCursor = null;
                MessageBox.Show("Unexpected value while adding walls. Have a typo in the ABOVE GRADE WALLS section?",
                    "Something went wrong",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            try { cp.ChangeBasment(); }
            catch
            {
                Mouse.OverrideCursor = null;
                MessageBox.Show("Error while changing basement. Check template has required basement elements then check spreadsheet for typos.",
                    "Something went wrong",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            try { cp.GasDHW(); }
            catch
            {
                Mouse.OverrideCursor = null;
                MessageBox.Show("Error while changing GAS DHW Check for typos in GAS DHW section.",
                    "Something went wrong",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            try { cp.ElectricDHW(); }
            catch
            {
                Mouse.OverrideCursor = null;
                MessageBox.Show("Error while changing ELECTRIC DHW Check for typos in ELECTRIC DHW section.",
                    "Something went wrong",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (Settings.Default.WindowsCheckbox)
            {
                cp.RemoveWindows();
                cp.ExtractWindows();
            }

            newHouse = CreateProp.GetHouse();
            Mouse.OverrideCursor = null;
            SaveFileDialog sfd = new SaveFileDialog
            {
                Title = "Save Generated Proposed House",
                Filter = " H2K files (*.h2k)| *.h2k",
                InitialDirectory = Path.GetDirectoryName(ExcelFilePath),
                FileName = $"{ProposedAddress}-PROPOSED",
            };

            if (sfd.ShowDialog() == true)
            {
                newHouse.Save(sfd.FileName, SaveOptions.None);
            }
            template = null;
            TemplatePath = null;
            TemplateText = "No template selected";
        }

        private void SetProfile()
        {
            if (Profiles != null && Profiles.Length > SelectedZoneIndex)
            {
                Profile = Profiles[SelectedZoneIndex];
            }
        }

        private string SplitAddress(string stringToSplit)
        {
            string[] pathStrings = Path.GetFileName(stringToSplit).Split('-');

            if (pathStrings.Length > 2)
            {
                string newAddress = pathStrings[0];
                for (int i = 1; i < pathStrings.Length - 1; i++)
                {
                    newAddress += $"-{pathStrings[i]}";
                }
                return newAddress;
            }
            else
            {
                return pathStrings[0].ToString();
            }
        }

        private void SetDefaultDirectory()
        {
            VistaFolderBrowserDialog fbd = new VistaFolderBrowserDialog();

            if (fbd.ShowDialog() == true)
            {
                Settings.Default.TemplateDir = fbd.SelectedPath;
                Settings.Default.Save();
            }
        }
    }
}
