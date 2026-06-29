using System;
using System.IO;
using HotPort.Properties;
using System.Windows;
using Microsoft.Win32;
using System.ComponentModel;
using System.Windows.Input;
using System.Xml.Linq;
using System.Linq;
using Ookii.Dialogs.Wpf;
using System.Reflection;
using HotPort.Models;

namespace HotPort
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow 

    {
        XDocument? propHouse;
        private XDocument? newHouse;
        public string? templatePath;
        string? excelFilePath;
        private string? proposedAddress;
        private string? directoryString;
        public string? ExcelPath { get; private set; }
        private XElement[] profiles;
        private XElement profile;

        public MainWindow()
        {
            InitializeComponent();
            this.Left = Settings.Default.WindowLeft;
            this.Top = Settings.Default.WindowTop;
            XDocument values = new XDocument(XDocument.Load(@".\ReferenceProfiles.xml"));

            profiles = (from el in values.Descendants("Zone")
                        select el).ToArray();
            foreach (XElement zone in profiles)
            {
                ZoneSelectBox.Items.Add(zone.Attribute("name")?.Value);
            }
            ZoneSelectBox.SelectedIndex = 0;

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
        private void WorksheetButton_Click(object sender, RoutedEventArgs e)
        {
            string excelAddress;
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Select worksheet";
            ofd.Filter = "Excel Files (*.xlsm) | *.xlsm";

            if (ofd.ShowDialog() == true)
            {
                worksheetTextBlock.Text = "";
                excelFilePath = ofd.FileName.ToString();
                excelAddress = SplitAddress(excelFilePath);
                ExcelPath = excelFilePath;
                worksheetTextBlock.Text = excelAddress;
                proposedAddress = excelAddress;
            }
        }
        private void SelectPropFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Select Proposed File";
            ofd.Filter = "House Files (*.h2k)|*.h2k";

            if (ofd.ShowDialog() == true)
            {
                propHouse = XDocument.Load(ofd.FileName);
                directoryString = Path.GetDirectoryName(ofd.FileName);
                proposedAddress = SplitAddress(Path.GetFileName(ofd.FileName));
                ProposedFileTextBlock.Text = proposedAddress;
            }
        }

        private void CreateRefButton_Click(object sender, RoutedEventArgs e)
        {
            if (propHouse == null)
            {
                MessageBox.Show("You must select a proposed file first.", "No file selected", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            else if (excelFilePath == null)
            {
                MessageBox.Show("You must select an Excel file first.", "No file selected", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            else
            {
                Cursor cursor = Cursors.Wait;
                SetProfile();
                CreateRef cr = new CreateRef(propHouse, excelFilePath, profile);
                cr.FindID(propHouse);
                propHouse = cr.Remover(propHouse);
                propHouse = cr.AddCode(propHouse);
                propHouse = cr.RChanger(propHouse);
                propHouse = cr.HeatingCooling(propHouse);
                propHouse = cr.ChangeACH(propHouse);
                propHouse = cr.AddFan(propHouse);
                try
                {
                    propHouse = cr.Doors(propHouse);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, 
                        "Error getting door width",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                try
                {
                    propHouse = cr.Windows(propHouse);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, 
                        "Error getting window size",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                propHouse = cr.HotWater(propHouse);
                Cursor = Cursors.Arrow;

                SaveFileDialog sfd = new SaveFileDialog
                {
                    Filter = "House File|*.h2k",
                    DefaultExt = "h2k",
                    InitialDirectory = directoryString,
                    FileName = $"{proposedAddress}-REFERENCE",
                };

                if (sfd.ShowDialog() == true)
                {
                    propHouse.Save(sfd.FileName);
                }
                ExcelHelper.CloseCachedDocuments();
            }
        }
        private void CreatePropButton_Click(object sender, RoutedEventArgs e)
        {
            if (templatePath == null)
            {
                MessageBox.Show("Select a builder template to modify",
                    "No builder template selected",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (excelFilePath == null)
            {
                MessageBox.Show("Select a spreadsheet to copy from.",
                    "No spreadsheet selected",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            Cursor = Cursors.Wait;
            XDocument template = new XDocument(XDocument.Load(templatePath));
            CreateProp cp = new CreateProp(excelFilePath, template);
            CreateProp.ChangeAddress(proposedAddress);

            try { cp.CreateHouse(); }
            catch(Exception ex)
            {
                Cursor = Cursors.Arrow;
                MessageBox.Show(ex.Message, "Invalid data", MessageBoxButton.OK, MessageBoxImage.Error);
                ExcelHelper.CloseCachedDocuments();
                return;
            }

            if (Properties.Settings.Default.WindowsCheckbox)
            {
                cp.RemoveWindows();
                cp.ExtractWindows();
            }
            
            newHouse = CreateProp.GetHouse();
            Cursor = Cursors.Arrow;
            SaveFileDialog sfd = new SaveFileDialog
            {
                Title = "Save Generated Proposed House",
                Filter = " H2K files (*.h2k)| *.h2k",
                InitialDirectory = Path.GetDirectoryName(excelFilePath),
                FileName = $"{proposedAddress}-PROPOSED",
            };

            if (sfd.ShowDialog() == true)
            {
                newHouse.Save(sfd.FileName, SaveOptions.None);
            }
            template = null;
            templatePath = null;
            TemplateTextBlock.Text = "No template selected";
            ExcelHelper.CloseCachedDocuments();
        }
        private void TemplateButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                Title = "Select HOT2000 builder template",
                Filter = "House Files(*.h2k) | *.h2k",
                InitialDirectory = Settings.Default.TemplateDir,

            };
            if (ofd.ShowDialog() == true)
            {
                templatePath = ofd.FileName;
                TemplateTextBlock.Text = ofd.SafeFileName.Split(".").First();
            }
        }
        private void Window_Closing(object sender, CancelEventArgs e)
        {
            Settings.Default.WindowLeft = this.Left;
            Settings.Default.WindowTop = this.Top;
            Settings.Default.Save();
        }
        private void DefaultDirectory_Click(object sender, RoutedEventArgs e)
        {
            VistaFolderBrowserDialog fbd = new VistaFolderBrowserDialog();

            if (fbd.ShowDialog() == true)
            {
                Settings.Default.TemplateDir = fbd.SelectedPath;
                Settings.Default.Save();
            }
        }
        private void SetProfile()
        {
            profile = profiles[ZoneSelectBox.SelectedIndex];
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
    }}



