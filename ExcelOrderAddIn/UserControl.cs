﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using ExcelOrderAddIn.Displays;
// ReSharper disable once RedundantUsingDirective
using ExcelOrderAddIn.Extensions;
using ExcelOrderAddIn.Logging;
using ExcelOrderAddIn.Model;
using ExcelOrderAddIn.Properties;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelOrderAddIn
{
    public partial class UserControl : System.Windows.Forms.UserControl
    {
        private const string GeneratedSheetName = "New Order";

        private IEnumerable<WorksheetItem> _worksheetItems;
        private readonly ILogger _logger = new Logger(100u);
        private static readonly IDisplay Display = new Display();

        public UserControl()
        {
            InitializeComponent();

            RefreshItems();

            InitializeImageFolderPicker();
        }

        private void InitializeImageFolderPicker()
        {
            imgFolderTextBox.Text = Settings.Default.ImgFolder;
            folderBrowserDialog.SelectedPath = Settings.Default.ImgFolder;
        }

        private void RefreshItems()
        {
            _worksheetItems = GetWorksheetItems();

            RefreshTableComboBox(table1ComboBox, Settings.Default.Table1, 0);
            RefreshTableComboBox(table2ComboBox, Settings.Default.Table2, 1);
            RefreshTableComboBox(table3ComboBox, Settings.Default.Table3, 2);
        }

        private static IEnumerable<WorksheetItem> GetWorksheetItems()
        {
            foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.Worksheets)
            {
                yield return new WorksheetItem(worksheet);
            }
        }

        private void RefreshTableComboBox(ComboBox comboBox, string preferredTable, int preferredIndex)
        {
            comboBox.Items.Clear();
            comboBox.Items.AddRange(_worksheetItems.ToArray());

            var allWorksheetsNames = comboBox.Items.OfType<WorksheetItem>().Select(item => item.Name);
            if (allWorksheetsNames.Contains(preferredTable))
            {
                comboBox.SelectedIndex = allWorksheetsNames.ToList().IndexOf(preferredTable);
            }

            if (comboBox.SelectedIndex == -1)
            {
                comboBox.SelectedIndex = Math.Min(preferredIndex, comboBox.Items.Count - 1);
            }
        }

        /**
         * Main logic of add-in
         */
        private async void createBtn_Click(object sender, EventArgs e)
        {
            UpdateProgress(0, "");

            if (!ValidateChildren(ValidationConstraints.Enabled)) return;

            Globals.ThisAddIn.Application.Interactive = false;

            try
            {
                // Start the timer that updates logs
                timer.Enabled = true;

                var table1 = Table.FromComboBoxes(_logger, Display, table1ComboBox, idCol1ComboBox);
                var table2 = Table.FromComboBoxes(_logger, Display, table2ComboBox, idCol2ComboBox);
                var table3 = Table.FromComboBoxes(_logger, Display, table3ComboBox, idCol3ComboBox);

                _logger.Info("Creating new worksheet.");
                var newWorksheet = CreateNewWorksheet();
                _logger.Info("Worksheet created.");

                var joined = table1.Join(table2).Join(table3);

                joined.CheckAvailableColumns();

                joined.RemoveUnavailableProducts();

                joined.InsertColumns();

                joined.RenameColumns();

                joined.SelectColumns();

                const int topOffset = 2;
                joined.PrintTotalPriceTable(newWorksheet, topOffset);

                UpdateProgress(40, "Printing data to worksheet...");
                await joined.PrintToWorksheet(newWorksheet, topOffset);

                UpdateProgress(80, "Inserting images...");
                await joined.InsertImages(newWorksheet, topOffset, imgFolderTextBox.Text);

                UpdateProgress(100, "Done.");
                var countOfRows = joined.Data.GetLength(0);
                Display.Show($"{countOfRows} row{(countOfRows == 1 ? "" : "s")} created.", "Success!",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                UpdateProgress(100, "Failed.");
                Display.Show($"Some error occurred. Details:\n{ex}", "Error!", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                Globals.ThisAddIn.Application.Interactive = true;
                timer.Enabled = false;
                // Print remaining logs
                UpdateLogWindow();
            }
        }

        private void UpdateProgress(int percent, string status)
        {
            BeginInvoke(new Action(() =>
            {
                progressBar.Value = percent;
                progressBarLabel.Text = status;
            }));
        }

        private static Excel.Worksheet CreateNewWorksheet()
        {
            var newWorksheet = Globals.ThisAddIn.Application.Worksheets.Add();
            var newName = FindNewName();
            newWorksheet.Name = newName;
            return newWorksheet;
        }

        private static string FindNewName()
        {
            var newName = GeneratedSheetName;
            var i = 2;
            while (Globals.ThisAddIn.Application.Worksheets.OfType<Excel.Worksheet>().Any(ws => ws.Name == newName))
            {
                newName = $"{GeneratedSheetName} {i++}";
            }

            return newName;
        }

        private void ValidateComboBox(ComboBox comboBox, CancelEventArgs e)
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

        private void table1ComboBox_Validating(object sender, CancelEventArgs e)
        {
            ValidateComboBox(table1ComboBox, e);
        }

        private void table2ComboBox_Validating(object sender, CancelEventArgs e)
        {
            ValidateComboBox(table2ComboBox, e);
        }

        private void table3ComboBox_Validating(object sender, CancelEventArgs e)
        {
            ValidateComboBox(table3ComboBox, e);
        }

        private void idCol1ComboBox_Validating(object sender, CancelEventArgs e)
        {
            ValidateComboBox(idCol1ComboBox, e);
        }

        private void idCol2ComboBox_Validating(object sender, CancelEventArgs e)
        {
            ValidateComboBox(idCol2ComboBox, e);
        }

        private void idCol3ComboBox_Validating(object sender, CancelEventArgs e)
        {
            ValidateComboBox(idCol3ComboBox, e);
        }

        private void table1ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            Settings.Default.Table1 = (table1ComboBox.SelectedItem as WorksheetItem).Name;

            RefreshIdColComboBox(table1ComboBox, idCol1ComboBox, Settings.Default.IdCol1);
            Validate();
        }

        private void table2ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            Settings.Default.Table2 = (table2ComboBox.SelectedItem as WorksheetItem).Name;

            RefreshIdColComboBox(table2ComboBox, idCol2ComboBox, Settings.Default.IdCol2);
            Validate();
        }

        private void table3ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            Settings.Default.Table3 = (table3ComboBox.SelectedItem as WorksheetItem).Name;

            RefreshIdColComboBox(table3ComboBox, idCol3ComboBox, Settings.Default.IdCol3);
            Validate();
        }

        private static void RefreshIdColComboBox(ComboBox tableComboBox, ComboBox idColComboBox,
            string preferredIdColumn)
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
            const string generatedSheetName = GeneratedSheetName;
            var generatedSheets = Globals.ThisAddIn.Application.Worksheets.OfType<Excel.Worksheet>()
                .Where(ws => ws.Name.StartsWith(generatedSheetName));
            DeleteWorksheets(generatedSheets);
        }

        private void deleteNotGeneratedSheetsBtn_Click(object sender, EventArgs e)
        {
            const string generatedSheetName = GeneratedSheetName;
            var notGeneratedSheets = Globals.ThisAddIn.Application.Worksheets.OfType<Excel.Worksheet>()
                .Where(ws => !ws.Name.StartsWith(generatedSheetName));
            DeleteWorksheets(notGeneratedSheets);
        }

        private static void DeleteWorksheets(IEnumerable<Excel.Worksheet> worksheets)
        {
            var count = worksheets.Count();

            if (count == 0)
            {
                Display.Show("No sheets to be deleted.", "No sheets", MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            if (Display.Show(
                $"Do you really want to delete the following {count} sheet{(count == 1 ? "" : "s")}: {string.Join(", ", worksheets.Select(ws => ws.Name))}?",
                "Confirm deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;

            var allSheetsCount = Globals.ThisAddIn.Application.Worksheets.OfType<Excel.Worksheet>().Count();

            if (count == allSheetsCount)
            {
                Display.Show($"Cannot delete all sheets.", "Invalid operation", MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            Globals.ThisAddIn.Application.Application.DisplayAlerts = false;
            foreach (var worksheet in worksheets)
            {
                worksheet.Delete();
            }

            Globals.ThisAddIn.Application.Application.DisplayAlerts = true;

            Display.Show($"{count} sheet{(count == 1 ? "" : "s")} {(count == 1 ? "has" : "have")} been deleted.",
                "Success!",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void idCol1ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            Settings.Default.IdCol1 = idCol1ComboBox.SelectedItem as string;
            Validate();
        }

        private void idCol2ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            Settings.Default.IdCol2 = idCol2ComboBox.SelectedItem as string;
            Validate();
        }

        private void idCol3ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            Settings.Default.IdCol3 = idCol3ComboBox.SelectedItem as string;
            Validate();
        }

        private void selectImgFolderBtn_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog.ShowDialog() != DialogResult.OK) return;

            imgFolderTextBox.Text = folderBrowserDialog.SelectedPath;
            Settings.Default.ImgFolder = folderBrowserDialog.SelectedPath;
        }

        private void refreshBtn_Click(object sender, EventArgs e)
        {
            RefreshItems();
        }

        private void TimeUpdateLogWindow_Tick(object sender, EventArgs e)
        {
            UpdateLogWindow();
        }

        private void UpdateLogWindow()
        {
            BeginInvoke(new Action(() =>
            {
                scrollingRichTextBox.Rtf = _logger.GetLogAsRichText(false);
                scrollingRichTextBox.ScrollToBottom();
            }));
        }
    }
}
