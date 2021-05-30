using Microsoft.Office.Tools.Ribbon;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelOrderAddIn
{
    public partial class UserControl : System.Windows.Forms.UserControl
    {
        private readonly IEnumerable<WorksheetItem> WorksheetItems;

        public UserControl()
        {
            InitializeComponent();

            WorksheetItems = GetWorksheetItems();

            table1ComboBox.Items.AddRange(WorksheetItems.ToArray());
            table2ComboBox.Items.AddRange(WorksheetItems.ToArray());
        }

        private IEnumerable<WorksheetItem> GetWorksheetItems()
        {
            foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
            {
                yield return new WorksheetItem(worksheet);
            }
        }
    }
}
