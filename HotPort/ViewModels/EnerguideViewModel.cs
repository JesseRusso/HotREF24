using HotPort.Infrastructure;
using HotPort.Models;
using HotPort.Properties;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows.Input;
using System.Xml.Linq;

namespace HotPort.ViewModels
{
    internal class EnerguideViewModel : ObservableObject
    {
        private readonly IDialogService _dialogs;
        private CodeLibrary? _codeLibrary;
        private string? _excelFilePath;

        // --- COD file status ---
        private string _codLabel = "No code library loaded";
        public string CodLabel
        {
            get => _codLabel;
            set => SetProperty(ref _codLabel, value);
        }

        // --- Wall codes (all three dropdowns share the same source list) ---
        public ObservableCollection<CodeEntry> WallCodes { get; } = new();

        private CodeEntry? _selectedMainWallCode;
        public CodeEntry? SelectedMainWallCode
        {
            get => _selectedMainWallCode;
            set => SetProperty(ref _selectedMainWallCode, value);
        }

        private CodeEntry? _selectedNoСladWallCode;
        public CodeEntry? SelectedNoCladWallCode
        {
            get => _selectedNoСladWallCode;
            set => SetProperty(ref _selectedNoСladWallCode, value);
        }

        private CodeEntry? _selectedTallWallCode;
        public CodeEntry? SelectedTallWallCode
        {
            get => _selectedTallWallCode;
            set => SetProperty(ref _selectedTallWallCode, value);
        }

        // --- Floor header ---
        public ObservableCollection<CodeEntry> FloorHeaderCodes { get; } = new();

        private CodeEntry? _selectedFloorHeaderCode;
        public CodeEntry? SelectedFloorHeaderCode
        {
            get => _selectedFloorHeaderCode;
            set => SetProperty(ref _selectedFloorHeaderCode, value);
        }

        // --- Ceilings ---
        public ObservableCollection<CodeEntry> CeilingCodes { get; } = new();

        private CodeEntry? _selectedCeilingCode;
        public CodeEntry? SelectedCeilingCode
        {
            get => _selectedCeilingCode;
            set => SetProperty(ref _selectedCeilingCode, value);
        }

        public ObservableCollection<CodeEntry> VaultCodes { get; } = new();

        private CodeEntry? _selectedVaultCode;
        public CodeEntry? SelectedVaultCode
        {
            get => _selectedVaultCode;
            set => SetProperty(ref _selectedVaultCode, value);
        }

        // --- Floors ---
        public ObservableCollection<CodeEntry> ExposedFloorCodes { get; } = new();

        private CodeEntry? _selectedExposedFloorCode;
        public CodeEntry? SelectedExposedFloorCode
        {
            get => _selectedExposedFloorCode;
            set => SetProperty(ref _selectedExposedFloorCode, value);
        }

        public ObservableCollection<CodeEntry> GarageFloorCodes { get; } = new();

        private CodeEntry? _selectedGarageFloorCode;
        public CodeEntry? SelectedGarageFloorCode
        {
            get => _selectedGarageFloorCode;
            set => SetProperty(ref _selectedGarageFloorCode, value);
        }

        // --- Foundation ---
        public ObservableCollection<CodeEntry> FloorsAboveCodes { get; } = new();

        private CodeEntry? _selectedFloorsAboveCode;
        public CodeEntry? SelectedFloorsAboveCode
        {
            get => _selectedFloorsAboveCode;
            set => SetProperty(ref _selectedFloorsAboveCode, value);
        }

        public ObservableCollection<CodeEntry> InteriorWallCodes { get; } = new();

        private CodeEntry? _selectedInteriorWallCode;
        public CodeEntry? SelectedInteriorWallCode
        {
            get => _selectedInteriorWallCode;
            set => SetProperty(ref _selectedInteriorWallCode, value);
        }

        public ObservableCollection<CodeEntry> PonyWallCodes { get; } = new();

        private CodeEntry? _selectedPonyWallCode;
        public CodeEntry? SelectedPonyWallCode
        {
            get => _selectedPonyWallCode;
            set => SetProperty(ref _selectedPonyWallCode, value);
        }

        // --- Template ---
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

        // --- Commands ---
        public ICommand SelectCodCommand { get; }
        public ICommand SelectTemplateCommand { get; }
        public ICommand CreateEnerguideCommand { get; }

        public EnerguideViewModel(IDialogService dialogs)
        {
            _dialogs = dialogs;
            SelectCodCommand = new RelayCommand(SelectCodManually);
            SelectTemplateCommand = new RelayCommand(SelectTemplate);
            CreateEnerguideCommand = new RelayCommand(CreateEnerguide);
        }

        // Called by MainWindowViewModel whenever the worksheet path changes
        public void OnWorksheetLoaded(string excelFilePath)
        {
            _excelFilePath = excelFilePath;
            TryAutoLoadCod(excelFilePath);
            TryAutoLoadTemplate(excelFilePath);
        }

        private void TryAutoLoadCod(string excelFilePath)
        {
            string codLibDir = Settings.Default.CodeLibDir;
            if (string.IsNullOrEmpty(codLibDir) || !Directory.Exists(codLibDir))
                return;

            string builderName = ExcelHelper.GetCellValue(excelFilePath, "Calc", "K1").Replace(" ", "");
            if (string.IsNullOrEmpty(builderName))
                return;

            string[] matches = Directory.GetFiles(codLibDir, "*.COD", SearchOption.TopDirectoryOnly)
                .Where(f => Path.GetFileName(f).Contains(builderName, System.StringComparison.OrdinalIgnoreCase))
                .ToArray();

            if (matches.Length == 1)
                LoadCod(matches[0]);
        }

        private void SelectCodManually()
        {
            if (_dialogs.TryOpenFile("Select code library", "COD Files (*.COD)|*.COD", out string path, Settings.Default.CodeLibDir))
                LoadCod(path);
        }

        private void LoadCod(string filePath)
        {
            _codeLibrary = new CodeLibrary(filePath);
            CodLabel = Path.GetFileName(filePath);

            Populate(WallCodes, _codeLibrary.GetWallCodes());
            Populate(FloorHeaderCodes, _codeLibrary.GetFloorHeaderCodes());
            Populate(CeilingCodes, _codeLibrary.GetCeilingCodes());
            Populate(VaultCodes, _codeLibrary.GetVaultCodes());
            Populate(ExposedFloorCodes, _codeLibrary.ExposedFloorCodes());
            Populate(GarageFloorCodes, _codeLibrary.GarageFloorCodes());
            Populate(FloorsAboveCodes, _codeLibrary.GetFloorsAboveCodes());
            Populate(InteriorWallCodes, _codeLibrary.GetInteriorWallCodes());
            Populate(PonyWallCodes, _codeLibrary.GetPonyWallCodes());

            if (_excelFilePath != null)
                InferSelections(_excelFilePath);
        }

        private void TryAutoLoadTemplate(string excelFilePath)
        {
            string templateDir = Settings.Default.TemplateDir;
            if (string.IsNullOrEmpty(templateDir) || !Directory.Exists(templateDir))
                return;
            string builderName = ExcelHelper.GetCellValue(excelFilePath, "Calc", "K1").Replace(" ", "");
            if (string.IsNullOrEmpty(builderName))
                return;
            string[] matches = Directory.GetFiles(templateDir, "*.h2k", SearchOption.TopDirectoryOnly)
                .Where(f => Path.GetFileName(f).Contains(builderName, System.StringComparison.OrdinalIgnoreCase))
                .ToArray();
            if (matches.Length > 0)
            {
                TemplatePath = matches[0];
                TemplateLabel = System.IO.Path.GetFileNameWithoutExtension(matches[0]);
            }
        }

        private void InferSelections(string excelFilePath)
        {
            // Read all wall rows (B21:B34 for spacing, D21:D34 for siding)
            var wallRows = new List<(string spacing, string siding)>();
            for (int row = 21; row <= 34; row++)
            {
                string spacing = ExcelHelper.GetCellValue(excelFilePath, "Calc", $"B{row}");
                string siding = ExcelHelper.GetCellValue(excelFilePath, "Calc", $"D{row}");
                if (!string.IsNullOrEmpty(spacing) || !string.IsNullOrEmpty(siding))
                    wallRows.Add((spacing, siding));
            }

            // Main wall: first row where siding is not "n/a" and spacing is not 12 OC
            var mainRow = wallRows.FirstOrDefault(r =>
                !string.IsNullOrEmpty(r.siding) &&
                !r.siding.Equals("n/a", System.StringComparison.OrdinalIgnoreCase) &&
                !r.spacing.Contains("12") && !r.spacing.Contains("8"));
            if (mainRow != default)
                SelectedMainWallCode = Match(WallCodes, _codeLibrary!.InferWallCode(mainRow.spacing, mainRow.siding));

            // No-clad wall: first row where siding is "n/a"
            var noCladRow = wallRows.FirstOrDefault(r =>
                r.siding.Equals("n/a", System.StringComparison.OrdinalIgnoreCase));
            if (noCladRow != default)
                SelectedNoCladWallCode = Match(WallCodes, _codeLibrary!.InferWallCode(noCladRow.spacing, "NoClad"));

            // Tall wall: first row where spacing contains "12"
            var tallRow = wallRows.FirstOrDefault(r => r.spacing.Contains("12"));
            if (tallRow != default)
                SelectedTallWallCode = Match(WallCodes, _codeLibrary!.InferWallCode(tallRow.spacing, tallRow.siding));

            // Floor header: use main wall siding
            string mainSiding = mainRow != default ? mainRow.siding : string.Empty;
            if (!string.IsNullOrEmpty(mainSiding))
                SelectedFloorHeaderCode = Match(FloorHeaderCodes, _codeLibrary!.InferFloorHeaderCode(mainSiding));
            else SelectedFloorHeaderCode = Match(FloorHeaderCodes, _codeLibrary!.InferFloorHeaderCode(""));

                SelectedCeilingCode = Match(CeilingCodes, _codeLibrary!.InferCeilingCode());
            SelectedVaultCode = Match(VaultCodes, _codeLibrary!.InferVaultCode());
            SelectedExposedFloorCode = Match(ExposedFloorCodes, _codeLibrary!.InferExposedFloorCode());
            SelectedGarageFloorCode = Match(GarageFloorCodes, _codeLibrary!.InferGarageFloorCode());
            SelectedFloorsAboveCode = Match(FloorsAboveCodes, _codeLibrary!.InferFloorsAboveCode());
            SelectedInteriorWallCode = Match(InteriorWallCodes, _codeLibrary!.InferInteriorWallCode());
            SelectedPonyWallCode = Match(PonyWallCodes, _codeLibrary!.InferPonyWallCode(mainSiding));
        }

        // Finds the collection entry whose label matches the inferred entry
        private static CodeEntry? Match(ObservableCollection<CodeEntry> collection, CodeEntry? inferred) =>
            inferred == null ? null : collection.FirstOrDefault(e => e.Label == inferred.Label);

        private static void Populate(ObservableCollection<CodeEntry> collection, IReadOnlyList<CodeEntry> entries)
        {
            collection.Clear();
            foreach (var entry in entries.OrderBy(e => e.Label))
                collection.Add(entry);
        }

        private void SelectTemplate()
        {
            if (_dialogs.TryOpenFile(
                    "Select HOT2000 builder template",
                    "House Files (*.h2k)|*.h2k",
                    out string path,
                    Settings.Default.TemplateDir))
            {
                TemplatePath = path;
                TemplateLabel = System.IO.Path.GetFileNameWithoutExtension(path);
            }
        }

        private void CreateEnerguide()
        {
            if (_templatePath == null)
            {
                _dialogs.ShowWarning("Select a builder template first.", "No template selected");
                return;
            }
            if (_excelFilePath == null)
            {
                _dialogs.ShowWarning("Select a worksheet first.", "No worksheet selected");
                return;
            }
            if (_codeLibrary == null)
            {
                _dialogs.ShowWarning("Select a code library first.", "No code library loaded");
                return;
            }

            XDocument template = XDocument.Load(_templatePath);

            var energuide = new Energuide(
                template,
                _excelFilePath,
                SelectedMainWallCode,
                SelectedNoCladWallCode,
                SelectedTallWallCode,
                SelectedFloorHeaderCode,
                SelectedCeilingCode,
                SelectedVaultCode,
                _codeLibrary.InferCathedralCode(),
                _codeLibrary.InferFlatCode(),
                SelectedExposedFloorCode,
                SelectedGarageFloorCode,
                SelectedFloorsAboveCode,
                SelectedInteriorWallCode,
                SelectedPonyWallCode);

            string address = MainWindowViewModel.SplitAddress(System.IO.Path.GetFileName(_excelFilePath));
            energuide.ChangeAddress(address);
            energuide.Generate();
            string defaultName = address + "-P";
            if (_dialogs.TrySaveFile(
                    "Save Energuide File",
                    "House Files (*.h2k)|*.h2k",
                    System.IO.Path.GetDirectoryName(_excelFilePath),
                    defaultName,
                    out string savePath))
            {
                energuide.House.Save(savePath, SaveOptions.None);
            }

            ExcelHelper.CloseCachedDocuments();
        }
    }
}
