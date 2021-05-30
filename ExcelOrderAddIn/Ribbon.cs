using Microsoft.Office.Tools.Ribbon;
using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelOrderAddIn
{
    public partial class Ribbon
    {
        private void openSidebarBtn_Click(object sender, RibbonControlEventArgs e)
        {
            UserControl userControl = new UserControl();
            var taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(userControl, "Order Add-In");
            taskPane.Width = 450;
            taskPane.Visible = true;
        }
    }
}
