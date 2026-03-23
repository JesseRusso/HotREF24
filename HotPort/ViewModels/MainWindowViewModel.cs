using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using HotPort.Infrastructure;
using HotPort.Properties;
using System.Windows.Input;

namespace HotPort.ViewModels
{
    public sealed class MainWindowViewModel : ObservableObject
    {
        private const string NoWorksheetSelected = "No worksheet selected";
        private const string NoTemplateSelected = "No template selected";
        private const string NoProposedFileSelected = "No file selected";

        private readonly RelayCommand createProposedCommand;
        private readonly RelayCommand createReferenceCommand;
        private readonly Action<bool> includeWindowsChanged;
        private readonly Action<int> selectedZoneIndexChanged;

        private string? worksheetPath;
        private string? templatePath;
        private string? proposedFilePath;
        private string worksheetDisplayText = NoWorksheetSelected;
        private string templateDisplayText = NoTemplateSelected;
        private string proposedFileDisplayText = NoProposedFileSelected;
        private int selectedZoneIndex;
        private bool includeWindows;

        public MainWindowViewModel(
            IEnumerable<string?> zoneNames,
            Action selectWorksheet,
            Action selectTemplate,
            Action createProposed,
            Action selectProposedFile,
            Action createReference,
            Action selectDefaultTemplateDirectory,
            Action<bool> includeWindowsChanged,
            Action<int> selectedZoneIndexChanged)
        {
            ZoneNames = new ObservableCollection<string>(CreateZoneList(zoneNames));
            this.includeWindowsChanged = includeWindowsChanged;
            this.selectedZoneIndexChanged = selectedZoneIndexChanged;

            includeWindows = Settings.Default.WindowsCheckbox;

            SelectWorksheetCommand = new RelayCommand(selectWorksheet);
            SelectTemplateCommand = new RelayCommand(selectTemplate);
            createProposedCommand = new RelayCommand(createProposed, CanCreateProposed);
            SelectProposedFileCommand = new RelayCommand(selectProposedFile);
            createReferenceCommand = new RelayCommand(createReference, CanCreateReference);
            SelectDefaultTemplateDirectoryCommand = new RelayCommand(selectDefaultTemplateDirectory);
        }

        public ObservableCollection<string> ZoneNames { get; }

        public ICommand SelectWorksheetCommand { get; }

        public ICommand SelectTemplateCommand { get; }

        public ICommand CreateProposedCommand => createProposedCommand;

        public ICommand SelectProposedFileCommand { get; }

        public ICommand CreateReferenceCommand => createReferenceCommand;

        public ICommand SelectDefaultTemplateDirectoryCommand { get; }

        public string? WorksheetPath
        {
            get => worksheetPath;
            set
            {
                if (SetProperty(ref worksheetPath, value))
                {
                    WorksheetDisplayText = string.IsNullOrWhiteSpace(value) ? NoWorksheetSelected : value;
                    NotifyCommandStateChanged();
                }
            }
        }

        public string? TemplatePath
        {
            get => templatePath;
            set
            {
                if (SetProperty(ref templatePath, value))
                {
                    TemplateDisplayText = string.IsNullOrWhiteSpace(value) ? NoTemplateSelected : value;
                    NotifyCommandStateChanged();
                }
            }
        }

        public string? ProposedFilePath
        {
            get => proposedFilePath;
            set
            {
                if (SetProperty(ref proposedFilePath, value))
                {
                    ProposedFileDisplayText = string.IsNullOrWhiteSpace(value) ? NoProposedFileSelected : value;
                    NotifyCommandStateChanged();
                }
            }
        }

        public string WorksheetDisplayText
        {
            get => worksheetDisplayText;
            private set => SetProperty(ref worksheetDisplayText, value);
        }

        public string TemplateDisplayText
        {
            get => templateDisplayText;
            private set => SetProperty(ref templateDisplayText, value);
        }

        public string ProposedFileDisplayText
        {
            get => proposedFileDisplayText;
            private set => SetProperty(ref proposedFileDisplayText, value);
        }

        public int SelectedZoneIndex
        {
            get => selectedZoneIndex;
            set
            {
                if (SetProperty(ref selectedZoneIndex, value))
                {
                    OnPropertyChanged(nameof(SelectedZoneName));
                    selectedZoneIndexChanged(value);
                }
            }
        }

        public bool IncludeWindows
        {
            get => includeWindows;
            set
            {
                if (SetProperty(ref includeWindows, value))
                {
                    includeWindowsChanged(value);
                }
            }
        }

        public string? SelectedZoneName =>
            SelectedZoneIndex >= 0 && SelectedZoneIndex < ZoneNames.Count
                ? ZoneNames[SelectedZoneIndex]
                : null;

        public void ResetTemplateSelection()
        {
            TemplatePath = null;
        }

        private bool CanCreateProposed()
        {
            return !string.IsNullOrWhiteSpace(WorksheetPath)
                && !string.IsNullOrWhiteSpace(TemplatePath);
        }

        private bool CanCreateReference()
        {
            return !string.IsNullOrWhiteSpace(WorksheetPath)
                && !string.IsNullOrWhiteSpace(ProposedFilePath);
        }

        private void NotifyCommandStateChanged()
        {
            createProposedCommand.RaiseCanExecuteChanged();
            createReferenceCommand.RaiseCanExecuteChanged();
        }

        private static List<string> CreateZoneList(IEnumerable<string?> zoneNames)
        {
            List<string> zones = new();

            foreach (string? zoneName in zoneNames)
            {
                if (!string.IsNullOrWhiteSpace(zoneName))
                {
                    zones.Add(zoneName);
                }
            }

            return zones;
        }
    }
}
