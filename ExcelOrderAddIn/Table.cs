using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelOrderAddIn
{
    class Table
    {
        private IList<string> Columns;
        public object[][] Data = new object[0][];
        private string IdCol;

        private int NCols { get => Columns.Count; }
        private int NRows { get => Data.Count(); }

        private int IdColIdx { get => Columns.IndexOf(IdCol); }

        private Table()
        {          
        }

        public Table(IList<string> columns, object[][] data, string idCol)
        {
            Columns = columns;
            Data = data;
            IdCol = idCol;
        }

        public Table Join(Table rightTable)
        {
            var leftIdColIdx = IdColIdx;
            var rightIdColIdx = rightTable.IdColIdx;

            // Join data on id columns, merge all columns (including the id column)
            var joinedData = Data
                .Join(rightTable.Data,
                leftRow => leftRow.ElementAt(leftIdColIdx),
                rightRow => rightRow.ElementAt(rightIdColIdx),
                (leftRow, rightRow) => leftRow.Concat(rightRow));

            // Columns in the right table that are already in the left table
            // Should be removed from the resulting table
            var removedCols = Columns.Intersect(rightTable.Columns).ToList();

            // Indices of those columns in the original table
            // In the joined table, the index is shifted by the number of columns in the left table
            var removedColsIndices = removedCols.Select(col => rightTable.Columns.IndexOf(col) + NCols);

            // Remove columns from the joined table on the found indices
            var filteredData = joinedData
                .Select(row => row.Where((value, index) => !removedColsIndices.Contains(index)));

            // Convert to proper type
            var resultData = filteredData.Select(row => row.ToArray()).ToArray();

            var newCols = Columns.Union(rightTable.Columns).ToList();

            return new Table(newCols, resultData, IdCol);
        }

        internal void PrintToWorksheet(Excel.Worksheet worksheet)
        {
            if (NCols == 0)
            {
                return;
            }

            // header
            var headerStartCell = worksheet.Cells[1, 1] as Excel.Range;
            var headerEndCell = worksheet.Cells[1, NCols] as Excel.Range;
            worksheet.Range[headerStartCell, headerEndCell].Value2 = Columns.ToExcelMultidimArray();

            if (NRows == 0)
            {
                return;
            }

            // skip header
            var dataStartCell = worksheet.Cells[2, 1] as Excel.Range;
            var dataEndCell = worksheet.Cells[NRows + 1, NCols] as Excel.Range;
            worksheet.Range[dataStartCell, dataEndCell].Value2 = Data.ToExcelMultidimArray();
        }

        internal static Table FromComboBoxes(ComboBox tableComboBox, ComboBox idColComboBox)
        {
            var worksheet = (tableComboBox.SelectedItem as WorksheetItem).Worksheet;
            var idCol = idColComboBox.SelectedItem as string;

            var table = new Table
            {
                IdCol = idCol,
                Columns = worksheet.GetColumnNames()
            };

            var nCols = worksheet.NCols();
            var nRows = worksheet.NRows();

            if (table.NCols == 0 || nRows == 0)
            {
                return table;
            }

            // skip header
            var dataStartCell = worksheet.Cells[2, 1] as Excel.Range;
            var dataEndCell = worksheet.Cells[nRows + 1, nCols] as Excel.Range;
            table.Data = (worksheet.Range[dataStartCell, dataEndCell].Value2 as object[,]).FromExcelMultidimArray();

            return table;
        }
    }
}
