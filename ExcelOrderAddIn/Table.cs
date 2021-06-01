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

            var newCols = Columns.Union(rightTable.Columns).ToList();

            return new Table(newCols, filteredData.ToJaggedArray(), IdCol);
        }

        internal void PrintToWorksheet(Excel.Worksheet worksheet, int topOffset = 0)
        {
            PrintRawDataToWorksheet(worksheet, topOffset);

            FormatData(worksheet);
        }

        internal void PrintRawDataToWorksheet(Excel.Worksheet worksheet, int topOffset)
        {
            if (NCols == 0)
            {
                return;
            }

            // header
            var headerStartCell = worksheet.Cells[1 + topOffset, 1] as Excel.Range;
            var headerEndCell = worksheet.Cells[1 + topOffset, NCols] as Excel.Range;
            var headerRange = worksheet.Range[headerStartCell, headerEndCell];
            headerRange.Value2 = Columns.ToExcelMultidimArray();
            Styling.Apply(headerRange, Styling.Style.HEADER);


            if (NRows == 0)
            {
                return;
            }

            // skip header
            var dataStartCell = worksheet.Cells[2 + topOffset, 1] as Excel.Range;
            var dataEndCell = worksheet.Cells[NRows + 1 + topOffset, NCols] as Excel.Range;
            worksheet.Range[dataStartCell, dataEndCell].Value2 = Data.ToExcelMultidimArray();
        }

        internal void SelectColumns()
        {
            var resultColumns = new List<string>() { 
                "Image",
                "Product",
                "Item",
                "EAN",
                "Description",
                "Colli (pcs in carton)",
                "NEW",
                "EXW CZ",
                "Order",
                "Total order",
                "RRP",
                "In stock",
                "Stock comming",
                "Note for stock",
                "Brand",
                "Category",
                "Product type",
                "Theme",
                "Country of origin",
            };

            var newOrderOfIndices = resultColumns.Select(col => Columns.IndexOf(col));

            Data = Data.
                Select(row => {
                    var reorderedValues = new object[resultColumns.Count];
                    for (int i = 0; i < resultColumns.Count; i++)
                    {
                        reorderedValues[i] = row[newOrderOfIndices.ElementAt(i)];
                    }
                    return reorderedValues;
                }).ToArray();
            Columns = resultColumns;
        }

        internal void RenameColumns()
        {
            var dict = new Dictionary<string, string>
                {
                    { "Produkt", "Product" },
                    { "Katalogové číslo", "Item" },
                    { "Popis alternativní", "Description" },
                {"Balení karton (ks)", "Colli (pcs in carton)" },
                    {"Cena", "EXW CZ" },
                { "Cena DMOC EUR", "RRP"},
                     {"K dispozici", "In stock" },
                 {"Bude k dispozici", "Stock comming" },
               {"Výrobce", "Brand" },
               {"Údaj 2", "Category" },
               {"Údaj 1", "Product type" },
               {"Země původu", "Country of origin" },
                };

            Columns = Columns.Select(col =>
            {
                if (dict.ContainsKey(col))
                {
                    return dict[col];
                }
                return col;
            }).ToList();
        }

        internal void InsertColumns()
        {
            InsertImageColumn();
            InsertNewColumn();
            InsertOrderColumn();
            InsertTotalOrderColumn();
            InsertNoteForStockColumn();
            InsertThemeColumn();
        }

        private void InsertThemeColumn()
        {
            InsertEmptyColumn("Theme");
        }

        private void InsertNoteForStockColumn()
        {
            InsertEmptyColumn("Note for stock");
        }

        private void InsertTotalOrderColumn()
        {
            InsertEmptyColumn("Total order");
        }

        private void InsertOrderColumn()
        {
            InsertEmptyColumn("Order");
        }

        private void InsertNewColumn()
        {
            InsertEmptyColumn("NEW");
        }

        private void InsertImageColumn()
        {
            InsertEmptyColumn("Image");
        }

        private void InsertEmptyColumn(string columnName)
        {
            Columns.Add(columnName);
            Data = Data
               .Select(row => row.Append(null))
               .ToJaggedArray();
        }

        internal void FormatData(Excel.Worksheet worksheet)
        {
            var usedRange = worksheet.UsedRange;
            usedRange.Columns.AutoFit();
            // TODO
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
