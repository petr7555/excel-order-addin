// ReSharper disable once RedundantUsingDirective
using Microsoft.Office.Tools.Ribbon;

namespace ExcelOrderAddIn
{
    // ReSharper disable once ClassNeverInstantiated.Global
    public partial class Ribbon
    {
        private void openSidebarBtn_Click(object sender, RibbonControlEventArgs e)
        {
            var userControl = new UserControl();
            var taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(userControl, "Order Add-In");
            taskPane.Width = 450;
            taskPane.Visible = true;
        }
    }
}
