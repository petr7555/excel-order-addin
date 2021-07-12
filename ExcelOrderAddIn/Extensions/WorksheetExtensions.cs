using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelOrderAddIn.Displays;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelOrderAddIn.Extensions
{
    public static class WorksheetExtensions
    {
        private static readonly IDisplay Display = new Display();

        public static int NCols(this Excel.Worksheet worksheet)
        {
            var n = 1;
            while (worksheet.Cells[1, n].Value2 != null)
            {
                n++;
            }

            return n - 2;
        }

        // Excluding input table header
        public static int NRows(this Excel.Worksheet worksheet)
        {
            var n = 2;
            while (worksheet.Cells[n, 1].Value2 != null)
            {
                n++;
            }

            return n - 3;
        }

        public static IList<string> GetColumnNames(this Excel.Worksheet worksheet)
        {
            if (!worksheet.Exists())
            {
                Display.Show("The selected worksheet does not exist anymore, please refresh.", "Refresh!",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return new List<string>();
            }

            var i = 1;
            object column;
            var result = new List<string>();
            while ((column = worksheet.Cells[1, i++].Value2) != null)
            {
                result.Add(column.ToString());
            }

            return result;
        }

        public static bool Exists(this Excel.Worksheet worksheet)
        {
            try
            {
                // accessing 'Name' property of non-existent worksheet will throw an exception
                var unused = worksheet.Name;
            }
            catch (COMException)
            {
                return false;
            }

            return true;
        }
    }
}
