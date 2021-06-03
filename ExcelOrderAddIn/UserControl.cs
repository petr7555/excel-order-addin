using System;
using System.Collections.Generic;
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
            Globals.ThisAddIn.Application.SheetActivate += Application_SheetActivate;

            InitializeImageFolderPicker();
        }

        private void InitializeImageFolderPicker()
        {
            imgFolderTextBox.Text = Properties.Settings.Default.ImgFolder;
            folderBrowserDialog.SelectedPath = Properties.Settings.Default.ImgFolder;
        }

        private void Application_SheetActivate(object Sh)
        {
            RefreshItems();
        }

        private void Application_SheetChange(object Sh, Excel.Range Target)
        {
            RefreshItems();
        }

        private void RefreshItems()
        {
            WorksheetItems = GetWorksheetItems();

            RefreshTableComboBox(table1ComboBox, Properties.Settings.Default.Table1, 0);
            RefreshTableComboBox(table2ComboBox, Properties.Settings.Default.Table2, 1);
            RefreshTableComboBox(table3ComboBox, Properties.Settings.Default.Table3, 2);
        }

        private IEnumerable<WorksheetItem> GetWorksheetItems()
        {
            foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
            {
                yield return new WorksheetItem(worksheet);
            }
        }

        private void RefreshTableComboBox(ComboBox comboBox, string preferredTable, int preferredIndex)
        {
            comboBox.Items.Clear();
            comboBox.Items.AddRange(WorksheetItems.ToArray());

            var allWorksheetsNames = comboBox.Items.OfType<WorksheetItem>().Select(item => item.Name);
            if (allWorksheetsNames.Contains(preferredTable))
            {
                comboBox.SelectedIndex = allWorksheetsNames.ToList().IndexOf(preferredTable);
            }
            // TODO zjistit, co je pohodlnější
            //if (comboBox.SelectedIndex == -1)
            //{
            //    comboBox.SelectedIndex = Math.Min(preferredIndex, comboBox.Items.Count - 1);
            //}
        }

        /**
         * 
         * Main logic of plugin
         * 
         * 
         * TODO add asynchronous progress bar
         * 
         */
        private void createBtn_Click(object sender, System.EventArgs e)
        {
            if (ValidateChildren(ValidationConstraints.Enabled))
            {

                Globals.ThisAddIn.Application.ScreenUpdating = false;

                var table1 = Table.FromComboBoxes(table1ComboBox, idCol1ComboBox);
                var table2 = Table.FromComboBoxes(table2ComboBox, idCol2ComboBox);
                var table3 = Table.FromComboBoxes(table3ComboBox, idCol3ComboBox);

                var newWorksheet = CreateNewWorksheet();

                var joined = table1.Join(table2).Join(table3);

                joined.RemoveUnavailableProducts();

                joined.InsertColumns();

                joined.RenameColumns();

                joined.SelectColumns();

                int topOffset = 2;
                joined.PrintTotalPriceTable(newWorksheet, topOffset);
                joined.PrintToWorksheet(newWorksheet, topOffset);

                joined.InsertImages(newWorksheet, topOffset, imgFolderTextBox.Text);

                Globals.ThisAddIn.Application.ScreenUpdating = true;

                MessageBox.Show($"{joined.Data.GetLength(0)} rows created.", "Success!");
            }
        }

        public static Excel.Worksheet CreateNewWorksheet()
        {
            Excel.Worksheet newWorksheet;
            newWorksheet = Globals.ThisAddIn.Application.Worksheets.Add();
            var newName = FindNewName();
            newWorksheet.Name = newName;
            return newWorksheet;
        }

        private static string FindNewName()
        {
            var newName = "New Order";
            var i = 2;
            while (Globals.ThisAddIn.Application.Worksheets.OfType<Excel.Worksheet>().Any(ws => ws.Name == newName))
            {
                newName = $"New Order {i++}";
            }
            return newName;
        }

        private void ValidateComboBox(ComboBox comboBox, System.ComponentModel.CancelEventArgs e)
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
            ValidateComboBox(table1ComboBox, e);
        }

        private void table2ComboBox_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            ValidateComboBox(table2ComboBox, e);
        }

        private void table3ComboBox_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            ValidateComboBox(table3ComboBox, e);
        }

        private void idCol1ComboBox_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            ValidateComboBox(idCol1ComboBox, e);
        }

        private void idCol2ComboBox_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            ValidateComboBox(idCol2ComboBox, e);
        }

        private void idCol3ComboBox_Validating(object sender, System.ComponentModel.CancelEventArgs e)
        {
            ValidateComboBox(idCol3ComboBox, e);
        }

        private void table1ComboBox_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            Properties.Settings.Default.Table1 = (table1ComboBox.SelectedItem as WorksheetItem).Name;

            RefreshIdColComboBox(table1ComboBox, idCol1ComboBox, Properties.Settings.Default.IdCol1);
        }

        private void table2ComboBox_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            Properties.Settings.Default.Table2 = (table2ComboBox.SelectedItem as WorksheetItem).Name;

            RefreshIdColComboBox(table2ComboBox, idCol2ComboBox, Properties.Settings.Default.IdCol2);
        }

        private void table3ComboBox_SelectedIndexChanged(object sender, System.EventArgs e)
        {
            Properties.Settings.Default.Table3 = (table3ComboBox.SelectedItem as WorksheetItem).Name;

            RefreshIdColComboBox(table3ComboBox, idCol3ComboBox, Properties.Settings.Default.IdCol3);
        }

        private void RefreshIdColComboBox(ComboBox tableComboBox, ComboBox idColComboBox, string preferredIdColumn)
        {
            idColComboBox.Items.Clear();
            var items = (tableComboBox.SelectedItem as WorksheetItem).Worksheet.GetColumnNames();
            idColComboBox.Items.AddRange(items.ToArray());

            if (items.Contains(preferredIdColumn))
            {
                idColComboBox.SelectedIndex = items.IndexOf(preferredIdColumn);
            }

            if (idColComboBox.SelectedIndex == -1 && idColComboBox.Items.Count > 0)
            {
                idColComboBox.SelectedIndex = 0;
            }
        }

        private void deleteGeneratedSheetsBtn_Click(object sender, EventArgs e)
        {
            var sheetName = "New Order";
            var generatedSheets = Globals.ThisAddIn.Application.Worksheets.OfType<Excel.Worksheet>().Where(ws => ws.Name.StartsWith(sheetName));
            var count = generatedSheets.Count();

            Globals.ThisAddIn.Application.Application.DisplayAlerts = false;
            foreach (var worksheet in generatedSheets)
            {
                worksheet.Delete();
            }
            Globals.ThisAddIn.Application.Application.DisplayAlerts = true;

            MessageBox.Show($"{count} sheets have been deleted.", "Success!");
        }

        private void idCol1ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.IdCol1 = idCol1ComboBox.SelectedItem as string;
        }

        private void idCol2ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.IdCol2 = idCol2ComboBox.SelectedItem as string;
        }

        private void idCol3ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.IdCol3 = idCol3ComboBox.SelectedItem as string;
        }

        private void selectImgFolderBtn_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                imgFolderTextBox.Text = folderBrowserDialog.SelectedPath;
                Properties.Settings.Default.ImgFolder = folderBrowserDialog.SelectedPath;
            }
        }
    }
}
