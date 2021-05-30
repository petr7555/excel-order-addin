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
        private object[,] Header = new object[0, 0];
        private object[,] Data = new object[0, 0];
        private IList<string> ColumnNames;
        private string IdCol;

        private int NCols { get => Header.GetLength(1); }
        private int NRows { get => Data.GetLength(0); }


        private Table()
        {          
        }

        public Table Join(Table other)
        {

            //return new Table(Rows.Join(other.Rows,
            //       r1 => r1.ColumnsAndValues[leftId],
            //       r2 => r2.ColumnsAndValues[rightId],
            //       (r1, r2) => {
            //           return new Row(r1.ColumnsAndValues.ToList().Select(kv => kv.Key).Concat(r2.ColumnsAndValues.ToList().Select(kv => kv.Key)).ToList(),
            //                                r1.ColumnsAndValues.ToList().Select(kv => kv.Value).Concat(r2.ColumnsAndValues.ToList().Select(kv => kv.Value)).ToList());
            //           }
            //   ).ToList()
            //);

            throw new NotImplementedException();
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

            worksheet.Range[headerStartCell, headerEndCell].Value2 = Header;

            if (NRows == 0)
            {
                return;
            }

            // skip header
            var dataStartCell = worksheet.Cells[2, 1] as Excel.Range;
            var dataEndCell = worksheet.Cells[NRows + 1, NCols] as Excel.Range;
            worksheet.Range[dataStartCell, dataEndCell].Value2 = Data;

        }

        internal static Table FromComboBoxes(ComboBox tableComboBox, ComboBox idColComboBox)
        {
            var worksheet = (tableComboBox.SelectedItem as WorksheetItem).Worksheet;
            var idCol = idColComboBox.SelectedItem as string;

            var table = new Table
            {
                IdCol = idCol,
                ColumnNames = worksheet.GetColumnNames()
            };

            var nCols = worksheet.NCols();
            var nRows = worksheet.NRows();

            if (nCols == 0)
            {
                return table;
            }

            // header
            var headerStartCell = worksheet.Cells[1, 1] as Excel.Range;
            var headerEndCell = worksheet.Cells[1, nCols] as Excel.Range;
            table.Header = worksheet.Range[headerStartCell, headerEndCell].Value2;

            if (nRows == 0)
            {
                return table;
            }

            // skip header
            var dataStartCell = worksheet.Cells[2, 1] as Excel.Range;
            var dataEndCell = worksheet.Cells[nRows + 1, nCols] as Excel.Range;
            table.Data = worksheet.Range[dataStartCell, dataEndCell].Value2;

            return table;
        }
    }
}
