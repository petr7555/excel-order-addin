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
            table3ComboBox.Items.AddRange(WorksheetItems.ToArray());
        }

        private IEnumerable<WorksheetItem> GetWorksheetItems()
        {
            foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
            {
                yield return new WorksheetItem(worksheet);
            }
        }

        private void table1ComboBox_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            AddColumns(table1ComboBox, idCol1ComboBox);
        }

        private void table2ComboBox_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            AddColumns(table2ComboBox, idCol2ComboBox);
        }

        private void table3ComboBox_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            AddColumns(table3ComboBox, idCol3ComboBox);
        }

        private void AddColumns(System.Windows.Forms.ComboBox tableComboBox, System.Windows.Forms.ComboBox idColComboBox)
        {
            idColComboBox.Items.Clear();
            var items = (tableComboBox.SelectedItem as WorksheetItem).GetColumns().ToArray();
            idColComboBox.Items.AddRange(items);
        }
    }
}
