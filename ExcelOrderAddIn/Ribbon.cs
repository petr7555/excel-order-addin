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

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            var ws = UserControl.CreateNewWorksheet();
            Styling.Apply(ws.Range["A1"], Styling.Style.CALCULATION);
            Styling.Apply(ws.Range["A2", "A4"], Styling.Style.INPUT);
            Styling.Apply(ws.Range["A5"], Styling.Style.HEADER);
            Styling.Apply(ws.Range["A6"], Styling.Style.RED_TEXT);
            Styling.Apply(ws.Range["A7"], Styling.Style.BOLD_TEXT);
        }
    }
}
