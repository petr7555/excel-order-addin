using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelOrderAddIn
{
    public partial class UserControl : System.Windows.Forms.UserControl
    {
        private IEnumerable<WorksheetItem> WorksheetItems;

        public UserControl()
        {
            InitializeComponent();

            RefreshItems();
            Globals.ThisAddIn.Application.SheetChange += Application_SheetChange;
        }

        private void Application_SheetChange(object Sh, Excel.Range Target)
        {
            RefreshItems();
        }

        private void RefreshItems()
        {
            WorksheetItems = GetWorksheetItems();

            RefreshTableComboBox(table1ComboBox, 0);
            RefreshTableComboBox(table2ComboBox, 1);
            RefreshTableComboBox(table3ComboBox, 2);
        }

        private void RefreshTableComboBox(ComboBox comboBox, int preferredIndex)
        {
            var prevIndex = comboBox.SelectedIndex;
            comboBox.Items.Clear();
            comboBox.Items.AddRange(WorksheetItems.ToArray());
            if (comboBox.Items.Count >= prevIndex + 1)
            {
                comboBox.SelectedIndex = prevIndex;
            }
            if (comboBox.SelectedIndex == -1)
            {
                comboBox.SelectedIndex = Math.Min(preferredIndex, comboBox.Items.Count - 1);
            }
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
            RefreshIdColComboBox(table1ComboBox, idCol1ComboBox);
        }

        private void table2ComboBox_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            RefreshIdColComboBox(table2ComboBox, idCol2ComboBox);
        }

        private void table3ComboBox_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            RefreshIdColComboBox(table3ComboBox, idCol3ComboBox);
        }

        private void RefreshIdColComboBox(ComboBox tableComboBox, ComboBox idColComboBox)
        {
            var prevIndex = idColComboBox.SelectedIndex;
            idColComboBox.Items.Clear();
            var items = (tableComboBox.SelectedItem as WorksheetItem).Worksheet.GetColumnNames().ToArray();
            idColComboBox.Items.AddRange(items);
            if (idColComboBox.Items.Count >= prevIndex + 1)
            {
                idColComboBox.SelectedIndex = prevIndex;
            }
            if (idColComboBox.SelectedIndex == -1 && idColComboBox.Items.Count > 0)
            {
                idColComboBox.SelectedIndex = 0;
            }
        }

        private void createBtn_Click(object sender, System.EventArgs e)
        {
            if (ValidateChildren(ValidationConstraints.Enabled))
            {

                var table1 = Table.FromComboBoxes(table1ComboBox, idCol1ComboBox);
                var table2 = Table.FromComboBoxes(table2ComboBox, idCol2ComboBox);
                var table3 = Table.FromComboBoxes(table3ComboBox, idCol3ComboBox);

                var newWorksheet = CreateNewWorksheet();

                var joined = table1.Join(table2).Join(table3);

                joined.PrintToWorksheet(newWorksheet);
                

                MessageBox.Show("Done");
            }
        }

        private Excel.Worksheet CreateNewWorksheet()
        {
            Excel.Worksheet newWorksheet;
            newWorksheet = Globals.ThisAddIn.Application.Worksheets.Add();
            var newName = "New Order";
            var i = 2;
            while (Globals.ThisAddIn.Application.Worksheets.OfType<Excel.Worksheet>().Any(ws => ws.Name == newName))
            {
                newName = $"New Order {i++}";
            }
            newWorksheet.Name = newName;
            return newWorksheet;
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

        private void UserControl_Enter(object sender, System.EventArgs e)
        {
            MessageBox.Show("yes");
        }
    }
}
