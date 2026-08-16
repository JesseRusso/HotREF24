using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace HotPort.Views
{
    public partial class WindowsDoorsSettingsView : UserControl
    {
        // Allows the text that would result from the keystroke to remain a valid
        // (possibly partial) non-negative decimal number, e.g. "", "37", "94.5".
        private static readonly Regex NumericPattern = new Regex(@"^\d*\.?\d*$");

        public WindowsDoorsSettingsView()
        {
            InitializeComponent();
            DataObject.AddPastingHandler(this, OnPaste);
        }

        private void Numeric_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !IsValidNumericResult((TextBox)sender, e.Text);
        }

        private void OnPaste(object sender, DataObjectPastingEventArgs e)
        {
            if (e.SourceDataObject.GetData(DataFormats.UnicodeText) is string pasted
                && Keyboard.FocusedElement is TextBox box)
            {
                if (!IsValidNumericResult(box, pasted))
                    e.CancelCommand();
            }
            else
            {
                e.CancelCommand();
            }
        }

        // Builds the text the box would hold after the input replaces the current
        // selection, then checks it against the numeric pattern.
        private static bool IsValidNumericResult(TextBox box, string input)
        {
            string proposed = box.Text
                .Remove(box.SelectionStart, box.SelectionLength)
                .Insert(box.SelectionStart, input);
            return NumericPattern.IsMatch(proposed);
        }
    }
}
