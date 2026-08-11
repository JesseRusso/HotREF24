using HotPort.Infrastructure;
using HotPort.Models;
using CreatePropModel = HotPort.Models.CreateProp;
using CreateRefModel = HotPort.Models.CreateRef;
using HotPort.Properties;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows.Input;
using System.Xml.Linq;

namespace HotPort.ViewModels
{
    internal class MainWindowViewModel : ObservableObject
    {
        private readonly IDialogService _dialogs;

        // --- Child ViewModels ---
        public EnerguideViewModel Energuide { get; }

        // --- Commands ---
        public ICommand WorksheetCommand { get; }
        public ICommand SelectPropFileCommand { get; }
        public ICommand TemplateCommand { get; }
        public ICommand CreateRefCommand { get; }
        public ICommand CreatePropCommand { get; }
        public ICommand DefaultDirectoryCommand { get; }
        public ICommand CodeLibDirectoryCommand { get; }

        // --- Zone profiles loaded from ReferenceProfiles.xml ---
        private readonly XElement[] _profiles;

        public ObservableCollection<string> ZoneNames { get; } = new ObservableCollection<string>();

        private int _selectedZoneIndex;
        public int SelectedZoneIndex
        {
            get => _selectedZoneIndex;
            set => SetProperty(ref _selectedZoneIndex, value);
        }

        public XElement SelectedProfile => _profiles[SelectedZoneIndex];

        // --- Excel worksheet ---
        private string? _excelFilePath;
        public string? ExcelFilePath
        {
            get => _excelFilePath;
            set => SetProperty(ref _excelFilePath, value);
        }

        private string _worksheetLabel = "No worksheet selected";
        public string WorksheetLabel
        {
            get => _worksheetLabel;
            set => SetProperty(ref _worksheetLabel, value);
        }

        // --- Builder template ---
        private string? _templatePath;
        public string? TemplatePath
        {
            get => _templatePath;
            set => SetProperty(ref _templatePath, value);
        }

        private string _templateLabel = "No template selected";
        public string TemplateLabel
        {
            get => _templateLabel;
            set => SetProperty(ref _templateLabel, value);
        }

        // --- Proposed h2k file ---
        private XDocument? _propHouse;
        public XDocument? PropHouse
        {
            get => _propHouse;
            set => SetProperty(ref _propHouse, value);
        }

        private string _proposedFileLabel = "No file selected";
        public string ProposedFileLabel
        {
            get => _proposedFileLabel;
            set => SetProperty(ref _proposedFileLabel, value);
        }

        private string? _proposedAddress;
        public string? ProposedAddress
        {
            get => _proposedAddress;
            set => SetProperty(ref _proposedAddress, value);
        }

        private string? _directoryString;
        public string? DirectoryString
        {
            get => _directoryString;
            set => SetProperty(ref _directoryString, value);
        }

        public MainWindowViewModel(IDialogService dialogs)
        {
            _dialogs = dialogs;
            Energuide = new EnerguideViewModel(dialogs);

            XDocument values = XDocument.Load(@".\ReferenceProfiles.xml");
            _profiles = values.Descendants("Zone").ToArray();
            foreach (XElement zone in _profiles)
                ZoneNames.Add(zone.Attribute("name")?.Value ?? string.Empty);
            _selectedZoneIndex = 0;

            WorksheetCommand = new RelayCommand(SelectWorksheet);
            SelectPropFileCommand = new RelayCommand(SelectPropFile);
            TemplateCommand = new RelayCommand(SelectTemplate);
            CreateRefCommand = new RelayCommand(CreateRef);
            CreatePropCommand = new RelayCommand(CreateProp);
            DefaultDirectoryCommand = new RelayCommand(SelectDefaultDirectory);
            CodeLibDirectoryCommand = new RelayCommand(SelectCodeLibDirectory);
        }

        private void SelectWorksheet()
        {
            if (_dialogs.TryOpenFile("Select worksheet", "Excel Files (*.xlsm) | *.xlsm", out string path))
            {
                ExcelFilePath = path;
                string address = SplitAddress(path);
                WorksheetLabel = address;
                ProposedAddress = address;
                Energuide.OnWorksheetLoaded(path);
            }
        }

        private void SelectPropFile()
        {
            if (_dialogs.TryOpenFile("Select Proposed File", "House Files (*.h2k)|*.h2k", out string path))
            {
                PropHouse = XDocument.Load(path);
                DirectoryString = Path.GetDirectoryName(path);
                ProposedAddress = SplitAddress(Path.GetFileName(path));
                ProposedFileLabel = ProposedAddress;
            }
        }

        private void SelectTemplate()
        {
            if (_dialogs.TryOpenFile(
                    "Select HOT2000 builder template",
                    "House Files(*.h2k) | *.h2k",
                    out string path,
                    Settings.Default.TemplateDir))
            {
                TemplatePath = path;
                TemplateLabel = Path.GetFileName(path).Split('.').First();
            }
        }

        private void CreateRef()
        {
            if (PropHouse == null)
            {
                _dialogs.ShowWarning("You must select a proposed file first.", "No file selected");
                return;
            }
            if (ExcelFilePath == null)
            {
                _dialogs.ShowWarning("You must select an Excel file first.", "No file selected");
                return;
            }

            var cr = new CreateRefModel(PropHouse, ExcelFilePath, SelectedProfile);
            cr.FindID(PropHouse);
            PropHouse = cr.Remover(PropHouse);
            PropHouse = cr.AddCode(PropHouse);
            PropHouse = cr.RChanger(PropHouse);
            PropHouse = cr.HeatingCooling(PropHouse);
            PropHouse = cr.ChangeACH(PropHouse);
            PropHouse = cr.AddFan(PropHouse);

            try { PropHouse = cr.Doors(PropHouse); }
            catch (Exception ex)
            {
                _dialogs.ShowError(ex.Message, "Error getting door width");
                ExcelHelper.CloseCachedDocuments();
                return;
            }

            try { PropHouse = cr.Windows(PropHouse); }
            catch (Exception ex)
            {
                _dialogs.ShowError(ex.Message, "Error getting window size");
                ExcelHelper.CloseCachedDocuments();
                return;
            }

            PropHouse = cr.HotWater(PropHouse);

            if (_dialogs.TrySaveFile(
                    null,
                    "House File|*.h2k",
                    DirectoryString,
                    $"{ProposedAddress}-REFERENCE",
                    out string savePath))
            {
                PropHouse.Save(savePath);
            }

            ExcelHelper.CloseCachedDocuments();
        }

        private void CreateProp()
        {
            if (TemplatePath == null)
            {
                _dialogs.ShowWarning("Select a builder template to modify", "No builder template selected");
                return;
            }
            if (ExcelFilePath == null)
            {
                _dialogs.ShowWarning("Select a spreadsheet to copy from.", "No spreadsheet selected");
                return;
            }

            XDocument template = new XDocument(XDocument.Load(TemplatePath));
            var cp = new CreatePropModel(ExcelFilePath, template);
            cp.ChangeAddress(ProposedAddress);

            try { cp.CreateHouse(); }
            catch (Exception ex)
            {
                _dialogs.ShowError(ex.Message, "Invalid data");
                ExcelHelper.CloseCachedDocuments();
                return;
            }

            if (Settings.Default.WindowsCheckbox)
            {
                cp.RemoveWindows();
                cp.ExtractWindows();
            }

            XDocument newHouse = cp.GetHouse();

            if (_dialogs.TrySaveFile(
                    "Save Generated Proposed House",
                    " H2K files (*.h2k)| *.h2k",
                    Path.GetDirectoryName(ExcelFilePath),
                    $"{ProposedAddress}-PROPOSED",
                    out string savePath))
            {
                newHouse.Save(savePath, SaveOptions.None);
            }

            TemplatePath = null;
            TemplateLabel = "No template selected";
            ExcelHelper.CloseCachedDocuments();
        }

        private void SelectDefaultDirectory()
        {
            if (_dialogs.TryOpenFolder(out string path))
            {
                Settings.Default.TemplateDir = path;
                Settings.Default.Save();
            }
        }

        private void SelectCodeLibDirectory()
        {
            if (_dialogs.TryOpenFolder(out string path))
            {
                Settings.Default.CodeLibDir = path;
                Settings.Default.Save();
            }
        }

        public static string SplitAddress(string filePath)
        {
            string[] parts = Path.GetFileName(filePath).Split('-');
            if (parts.Length > 2)
            {
                string address = parts[0];
                for (int i = 1; i < parts.Length - 1; i++)
                    address += $"-{parts[i]}";
                return address;
            }
            return parts[0];
        }
    }
}
