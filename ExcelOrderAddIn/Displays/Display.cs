using System.Windows.Forms;

namespace ExcelOrderAddIn.Displays
{
    public class Display : IDisplay
    {
        public DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            return MessageBox.Show(text, caption, buttons, icon);
        }
    }
}
