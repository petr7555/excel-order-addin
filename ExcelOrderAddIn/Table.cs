using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelOrderAddIn
{
    class Table
    {
        // TODO Could be configurable
        private const int ImgColHeight = 76;
        private const int ImgColWidth = 13;
        private const int ImgSize = 72; // = 96 pixels
        private const string PriceFormat = "#,###,###.00 €";
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

            //FormatData(worksheet);
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
            var dataRange = worksheet.Range[dataStartCell, dataEndCell];
            dataRange.Value2 = Data.ToExcelMultidimArray();

            // Set row height so that images fit
            dataRange.RowHeight = ImgColHeight;

            // Autofit all columns
            var usedRange = worksheet.UsedRange;
            usedRange.Columns.AutoFit();

            // Make 'Image' column wider
            worksheet.Columns[1].ColumnWidth = ImgColWidth;

            // Format 'Order'
            ApplyStyleToColumn(worksheet, topOffset, Styling.Style.INPUT, "Order");

            // Format 'Total order'
            FormatTotalOrder(worksheet, topOffset);

            // Format 'Stock comming'
            ApplyStyleToColumn(worksheet, topOffset, Styling.Style.YELLOW, "Stock comming");

            // Format 'Note for stock'
            ApplyStyleToColumn(worksheet, topOffset, Styling.Style.YELLOW, "Note for stock");

        }

        private void FormatTotalOrder(Excel.Worksheet worksheet, int topOffset)
        {
            ApplyStyleToColumn(worksheet, topOffset, Styling.Style.CALCULATION, "Total order");
            // 
            var totalOrderIndex = Columns.IndexOf("Total order") + 1;
            var priceIndex = Columns.IndexOf("EXW CZ") + 1;
            var orderIndex = Columns.IndexOf("Order") + 1;
            
            for (int i = 0; i < NRows; i++)
            {
                var row = topOffset + 2 + i;
                worksheet.Cells[i + 2 + topOffset, totalOrderIndex].Formula =
                $"={(priceIndex).ToLetter()}{row}*" +
                $"{(orderIndex).ToLetter()}{row}";
            }
            var range = GetColumnRange(worksheet, topOffset, "Total order");
            range.NumberFormat = PriceFormat;
        }

        private void ApplyStyleToColumn(Excel.Worksheet worksheet, int topOffset, Styling.Style style, string columnName)
        {
            var range = GetColumnRange(worksheet, topOffset, columnName);
            Styling.Apply(range, style);
        }

        private Excel.Range GetColumnRange(Excel.Worksheet worksheet, int topOffset, string columnName)
        {
            var colIndex = Columns.IndexOf(columnName) + 1;
            var startCell = worksheet.Cells[2 + topOffset, colIndex] as Excel.Range;
            var endCell = worksheet.Cells[NRows + 1 + topOffset, colIndex] as Excel.Range;
            return worksheet.Range[startCell, endCell]; ;
        }

        internal void PrintTotalPriceTable(Excel.Worksheet worksheet, int topOffset)
        {
            // Index of 'Order' column in Excel's 'starting from 1 system'
            var orderColIndex = Columns.IndexOf("Order") + 1;

            var titleCell = worksheet.Cells[1, orderColIndex - 1];
            titleCell.Value2 = "Total order";
            Styling.Apply(titleCell, Styling.Style.HEADER);

            var unitsCell = worksheet.Cells[1, orderColIndex];
            Styling.Apply(unitsCell, Styling.Style.CALCULATION);
            unitsCell.Formula = $"=SUM(" +
                $"{orderColIndex.ToLetter()}{topOffset + 2}:" +
                $"{orderColIndex.ToLetter()}{topOffset + 1 + NRows})";

            // Assumes that 'Total order' follows directly after 'Order'
            var totalPriceCell = worksheet.Cells[1, orderColIndex + 1];
            Styling.Apply(totalPriceCell, Styling.Style.CALCULATION);
            totalPriceCell.NumberFormat = PriceFormat;
            totalPriceCell.Formula = $"=SUM(" +
                $"{(orderColIndex + 1).ToLetter()}{topOffset + 2}:" +
                $"{(orderColIndex + 1).ToLetter()}{topOffset + 1 + NRows})";
        }

        /**
         * Assumes that 'Image' column is first.
         * Assumes 'Katalogové číslo' is translated as 'Item'.
         * Only one selection rule applies now:
         *  - image name == value in 'Item' column
         */
        internal void InsertImages(Excel.Worksheet worksheet, int topOffset, string imgFolder)
        {
            var defaultRowSize = 15;

            var imgNames = Data
                .Select(row => row[Columns.IndexOf("Item")] as string);

            var imgIdx = 0;
            foreach (var imgName in imgNames)
            {           
                if (FindImagePath(imgFolder, imgName, out var imgPath))
                {
                    worksheet.Shapes.AddPicture(imgPath, MsoTriState.msoFalse, MsoTriState.msoCTrue, 0, ((topOffset+1) * defaultRowSize) + (ImgColHeight * imgIdx) + (ImgColHeight-ImgSize) /2, ImgSize, ImgSize);
                }
                imgIdx++;
            }
        }

        /**
         * Returns true if image is found and sets imgPath.
         * Returns false if the image is not found, imgPath is set to empty string and should not be used.
         */
        private bool FindImagePath(string imgFolder, string imgName, out string imgPath)
        {
            var extensions = new string[] { "jpg", "png", "jpeg" };

            foreach(var extension in extensions)
            {
                var possiblePath = Path.Combine(imgFolder, $"{imgName}.{extension}");
                if (File.Exists(possiblePath))
                {
                    imgPath = possiblePath;
                    return true;
                }
            }

            imgPath = "";
            return false;
        }

        internal void RemoveUnavailableProducts()
        {
            Data = Data
                .Where(row => !(
                (Convert.ToInt32(row[Columns.IndexOf("Bude k dispozici")]) == 0 &&
                (Convert.ToString(row[Columns.IndexOf("Údaj sklad 1")]).Contains("ukončeno") ||
                Convert.ToString(row[Columns.IndexOf("Údaj sklad 1")]).Contains("doprodej")
                )) || Convert.ToString(row[Columns.IndexOf("Údaj sklad 1")]).Contains("POS")

                ))
                .ToJaggedArray();
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
                "Will be available",
                "Stock comming",
                "Note for stock",
                "Brand",
                "Category",
                "Product type",
                "Theme",
                "Country of origin",
            };

            var newOrderOfIndices = resultColumns
                .Select(col => Columns.IndexOf(col));

            Data = Data
                .Select(row => newOrderOfIndices.Select(index => row[index]))
                .ToJaggedArray();

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
                { "Bude bude", "Will be available"},
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

            InsertBudeBude();
        }

        private void InsertBudeBude()
        {

            Columns.Add("Bude bude");
            Data = Data
               .Select(row => row.Append(
                   Convert.ToInt32(row[Columns.IndexOf("Bude k dispozici")]) +
                   Convert.ToInt32(row[Columns.IndexOf("OBJEDNÁNO")]) -
                   Convert.ToInt32(row[Columns.IndexOf("DODAT")])
                   ))
               .ToJaggedArray();
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

        //internal void FormatData(Excel.Worksheet worksheet)
        //{
        //    var usedRange = worksheet.UsedRange;
        //    usedRange.Columns.AutoFit();

        //    // TODO
        //}

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
