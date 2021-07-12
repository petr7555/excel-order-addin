using System.Windows.Forms;

namespace ExcelOrderAddIn.Displays
{
    public interface IDisplay
    {
        DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon);
    }
}
