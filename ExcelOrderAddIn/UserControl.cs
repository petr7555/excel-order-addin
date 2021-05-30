using Microsoft.Office.Tools.Ribbon;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
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

        private void createBtn_Click(object sender, System.EventArgs e)
        {
            if (ValidateChildren(ValidationConstraints.Enabled))
            {
                MessageBox.Show("msg", "Success", MessageBoxButtons.OK);
            }
        }

        private void validateComboBox(ComboBox comboBox, System.ComponentModel.CancelEventArgs e)
        {
            if (comboBox.SelectedIndex == -1)
            {
                e.Cancel = true;
                errorProvider.SetError(comboBox, "Select a value.");
            }
            else
            {
                e.Cancel = false;
                errorProvider.SetError(comboBox, null);
            }
        }

        private void table1ComboBox_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            validateComboBox(table1ComboBox, e);
        }

        private void table2ComboBox_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            validateComboBox(table2ComboBox, e);
        }

        private void table3ComboBox_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            validateComboBox(table3ComboBox, e);
        }

        private void idCol1ComboBox_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            validateComboBox(idCol1ComboBox, e);
        }

        private void idCol2ComboBox_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            validateComboBox(idCol2ComboBox, e);
        }

        private void idCol3ComboBox_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            validateComboBox(idCol3ComboBox, e);
        }
    }
}
