using System.Windows;
using HotPort.Services;
using HotPort.ViewModels;
using Xunit;

namespace HotPort.Tests
{
    public class StubFileDialogService : IFileDialogService
    {
        public string? OpenFilePath { get; set; }
        public string? SaveFilePath { get; set; }
        public string? FolderPath { get; set; }
        public string? OpenFile(string title, string filter, string? initialDirectory = null) => OpenFilePath;
        public string? SaveFile(string title, string filter, string? initialDirectory = null, string? fileName = null) => SaveFilePath;
        public string? SelectFolder() => FolderPath;
    }

    public class StubMessageService : IMessageService
    {
        public void ShowMessage(string message, string caption, MessageBoxButton buttons = MessageBoxButton.OK, MessageBoxImage icon = MessageBoxImage.None)
        {
        }
    }

    public class MainViewModelTests
    {
        [Fact]
        public void SelectWorksheetCommand_UsesFileDialogService()
        {
            var fileService = new StubFileDialogService { OpenFilePath = "C\\temp\\test.xlsm" };
            var vm = new MainViewModel(fileService, new StubMessageService());
            vm.SelectWorksheetCommand.Execute(null);
            Assert.Equal("C\\temp\\test.xlsm", vm.ExcelFilePath);
        }
    }
}
