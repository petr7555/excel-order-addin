using System.Windows.Forms;
using ExcelOrderAddIn.Displays;

namespace Tests
{
    public class TestDisplay : IDisplay
    {
        public string LastDisplayedMessage { get; private set; }

        public DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            LastDisplayedMessage = text;
            return DialogResult.Yes;
        }
    }
}
