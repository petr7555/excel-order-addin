// ReSharper disable once RedundantUsingDirective

using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelOrderAddIn.Extensions;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelOrderAddIn.Model
{
    internal class Table
    {
        private enum ColumnImportance
        {
            Mandatory,
            Optional,
        }

        // TODO Could be configurable
        private const int ImgColHeight = 76;
        private const int ImgColWidth = 13;
        private const int ImgSize = 72; // = 96 pixels

        // It is good to figure these out by recording VB macro in Excel.
        // Quotes are escaped by doubling them both in VB and regex.
        private const string AccountingFormat =
            @"_([$€-x-euro2] * #,##0.00_);_([$€-x-euro2] * (#,##0.00);_([$€-x-euro2] * ""-""??_);_(@_)";

        private const string IntegerFormat = "0";
        private const string TextFormat = "@";

        // Column names
        private const string Produkt = "Produkt";
        private const string KatalogoveCislo = "Katalogové číslo";
        private const string PopisAlternativni = "Popis alternativní";
        private const string BaleníKartonKs = "Balení karton (ks)";
        private const string Cena = "Cena";
        private const string CenaDmocEur = "Cena DMOC EUR";
        private const string KDispozici = "K dispozici";
        private const string BudeKDispozici = "Bude k dispozici";
        private const string Vyrobce = "Výrobce";
        private const string Udaj1 = "Údaj 1";
        private const string Udaj2 = "Údaj 2";
        private const string UdajSklad1 = "Údaj sklad 1";
        private const string ZemePuvodu = "Země původu";
        private const string Product = "Product";
        private const string Item = "Item";
        private const string Description = "Description";
        private const string ColliPcsInCarton = "Colli (pcs in carton)";
        private const string ExwCz = "EXW CZ";
        private const string Rrp = "RRP";
        private const string InStock = "In stock";
        private const string StockComing = "Stock coming";
        private const string Brand = "Brand";
        private const string Category = "Category";
        private const string ProductType = "Product type";
        private const string CountryOfOrigin = "Country of origin";
        private const string Objednano = "OBJEDNÁNO";
        private const string Dodat = "DODAT";
        private const string Image = "Image";
        private const string Ean = "EAN";
        private const string New = "NEW";
        private const string Order = "Order";
        private const string TotalOrder = "Total order";
        private const string WillBeAvailable = "Will be available";
        private const string NoteForStock = "Note for stock";
        private const string Theme = "Theme";

        private IList<string> _columns;
        public object[][] Data = new object[0][];
        private string _idCol;

        private int NCols => _columns.Count;

        private int NRows => Data.Length;

        private int IdColIdx => _columns.IndexOf(_idCol);

        private Table()
        {
        }

        private Table(IList<string> columns, object[][] data, string idCol)
        {
            _columns = columns;
            Data = data;
            _idCol = idCol;
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
            var removedCols = _columns.Intersect(rightTable._columns).ToList();

            // Indices of those columns in the original table
            // In the joined table, the index is shifted by the number of columns in the left table
            var removedColsIndices = removedCols.Select(col => rightTable._columns.IndexOf(col) + NCols);

            // Remove columns from the joined table on the found indices
            var filteredData = joinedData
                .Select(row => row.Where((value, index) => !removedColsIndices.Contains(index)));

            var newCols = _columns.Union(rightTable._columns).ToList();

            return new Table(newCols, filteredData.ToJaggedArray(), _idCol);
        }

        internal async Task PrintToWorksheet(Excel.Worksheet worksheet, int topOffset = 0)
        {
            await Task.Run(() =>
            {
                if (NCols == 0)
                {
                    return;
                }

                // insert header
                var headerStartCell = worksheet.Cells[topOffset + 1, 1] as Excel.Range;
                var headerEndCell = worksheet.Cells[topOffset + 1, NCols] as Excel.Range;
                var headerRange = worksheet.Range[headerStartCell, headerEndCell];
                headerRange.Value2 = _columns.ToExcelMultidimArray();
                Styling.Apply(headerRange, Styling.Style.Header);

                // insert data
                var dataStartCell = worksheet.Cells[topOffset + 2, 1] as Excel.Range;
                var dataEndCell = worksheet.Cells[topOffset + 1 + Math.Max(NRows, 1), NCols] as Excel.Range;
                var dataRange = worksheet.Range[dataStartCell, dataEndCell];
                dataRange.Value2 = Data.ToExcelMultidimArray();

                // Auto-fit all columns
                worksheet.UsedRange.Columns.AutoFit();

                // Set row height so that images fit
                dataRange.RowHeight = ImgColHeight;

                FormatImageColumn(worksheet);
                FormatEanColumn(worksheet, topOffset);
                FormatColliColumn(worksheet, topOffset);
                FormatNewColumn(worksheet, topOffset);
                FormatExwCzColumn(worksheet, topOffset);
                FormatOrderColumn(worksheet, topOffset);
                FormatTotalOrderColumn(worksheet, topOffset);
                FormatRrpColumn(worksheet, topOffset);
                FormatInStockColumn(worksheet, topOffset);
                FormatWillBeAvailableColumn(worksheet, topOffset);
                FormatStockComingColumn(worksheet, topOffset);
                FormatNoteForStockColumn(worksheet, topOffset);

                AddBorder(headerRange);
                AddBorder(dataRange);
            });
        }

        private void AddBorder(Excel.Range range)
        {
            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
        }

        private void FormatWillBeAvailableColumn(Excel.Worksheet worksheet, int topOffset)
        {
            ApplyStyleToColumn(worksheet, topOffset, Styling.Style.Yellow, WillBeAvailable);
            FormatColumnAsInteger(worksheet, topOffset, StockComing);
        }

        private void FormatInStockColumn(Excel.Worksheet worksheet, int topOffset)
        {
            ApplyStyleToColumn(worksheet, topOffset, Styling.Style.SalmonBold, InStock);
            FormatColumnAsInteger(worksheet, topOffset, InStock);
        }

        private void FormatRrpColumn(Excel.Worksheet worksheet, int topOffset)
        {
            ApplyStyleToColumn(worksheet, topOffset, Styling.Style.BoldText, Rrp);
            FormatColumnAsAccounting(worksheet, topOffset, Rrp);
        }

        private void FormatExwCzColumn(Excel.Worksheet worksheet, int topOffset)
        {
            FormatColumnAsAccounting(worksheet, topOffset, ExwCz);
        }

        private void FormatNewColumn(Excel.Worksheet worksheet, int topOffset)
        {
            ApplyStyleToColumn(worksheet, topOffset, Styling.Style.RedBoldText, New);

            var colIndex = _columns.IndexOf(New) + 1;
            var headerCell = worksheet.Cells[topOffset + 1, colIndex];
            Styling.Apply(headerCell, Styling.Style.RedBoldHeaderText);
        }

        private void FormatColliColumn(Excel.Worksheet worksheet, int topOffset)
        {
            FormatColumnAsInteger(worksheet, topOffset, ColliPcsInCarton);
        }

        private void FormatNoteForStockColumn(Excel.Worksheet worksheet, int topOffset)
        {
            ApplyStyleToColumn(worksheet, topOffset, Styling.Style.Yellow, NoteForStock);
            FormatColumnAsText(worksheet, topOffset, NoteForStock);
        }

        private void FormatStockComingColumn(Excel.Worksheet worksheet, int topOffset)
        {
            ApplyStyleToColumn(worksheet, topOffset, Styling.Style.Yellow, StockComing);
            FormatColumnAsInteger(worksheet, topOffset, StockComing);
        }

        private static void FormatImageColumn(Excel.Worksheet worksheet)
        {
            // Make column wider
            worksheet.Columns[1].ColumnWidth = ImgColWidth;
        }

        private void FormatOrderColumn(Excel.Worksheet worksheet, int topOffset)
        {
            ApplyStyleToColumn(worksheet, topOffset, Styling.Style.Input, Order);
        }

        private void FormatEanColumn(Excel.Worksheet worksheet, int topOffset)
        {
            FormatColumnAsInteger(worksheet, topOffset, Ean);
        }

        private void FormatTotalOrderColumn(Excel.Worksheet worksheet, int topOffset)
        {
            ApplyStyleToColumn(worksheet, topOffset, Styling.Style.Calculation, TotalOrder);
            InsertTotalOrderFormula(worksheet, topOffset);
            FormatColumnAsAccounting(worksheet, topOffset, TotalOrder);
        }

        private void FormatColumnAsAccounting(Excel.Worksheet worksheet, int topOffset, string columnName)
        {
            var range = GetColumnRange(worksheet, topOffset, columnName);
            range.NumberFormat = AccountingFormat;
        }

        private void FormatColumnAsInteger(Excel.Worksheet worksheet, int topOffset, string columnName)
        {
            var range = GetColumnRange(worksheet, topOffset, columnName);
            range.NumberFormat = IntegerFormat;
        }

        private void FormatColumnAsText(Excel.Worksheet worksheet, int topOffset, string columnName)
        {
            var range = GetColumnRange(worksheet, topOffset, columnName);
            range.NumberFormat = TextFormat;
        }

        internal void CheckAvailableColumns()
        {
            var importanceDict = new Dictionary<string, ColumnImportance>
            {
                {Produkt, ColumnImportance.Mandatory},
                {KatalogoveCislo, ColumnImportance.Mandatory},
                {PopisAlternativni, ColumnImportance.Mandatory},
                {BaleníKartonKs, ColumnImportance.Mandatory},
                {Cena, ColumnImportance.Mandatory},
                {CenaDmocEur, ColumnImportance.Mandatory},
                {KDispozici, ColumnImportance.Mandatory},
                {BudeKDispozici, ColumnImportance.Mandatory},
                {Vyrobce, ColumnImportance.Mandatory},
                {UdajSklad1, ColumnImportance.Mandatory},
                {Udaj1, ColumnImportance.Mandatory},
                {Udaj2, ColumnImportance.Mandatory},
                {Objednano, ColumnImportance.Mandatory},
                {Dodat, ColumnImportance.Mandatory},
                {ZemePuvodu, ColumnImportance.Optional},
            };

            var notFoundColumns = importanceDict
                .Where(x => _columns.IndexOf(x.Key) == -1 && x.Value == ColumnImportance.Mandatory).ToList();
            if (notFoundColumns.Count > 0)
            {
                throw new InvalidDataException(
                    $"Data do not contain the following columns: {string.Join(", ", notFoundColumns)}.");
            }
        }
        
        private void InsertTotalOrderFormula(Excel.Worksheet worksheet, int topOffset)
        {
            var totalOrderIndex = _columns.IndexOf(TotalOrder) + 1;
            var priceIndex = _columns.IndexOf(ExwCz) + 1;
            var orderIndex = _columns.IndexOf(Order) + 1;

            Parallel.For(0, NRows, i =>
            {
                var row = topOffset + 2 + i;
                worksheet.Cells[row, totalOrderIndex].Formula =
                    $"={priceIndex.ToLetter()}{row}*" +
                    $"{orderIndex.ToLetter()}{row}";
            });
        }

        private void ApplyStyleToColumn(Excel.Worksheet worksheet, int topOffset, Styling.Style style,
            string columnName)
        {
            var range = GetColumnRange(worksheet, topOffset, columnName);
            Styling.Apply(range, style);
        }

        private Excel.Range GetColumnRange(Excel.Worksheet worksheet, int topOffset, string columnName)
        {
            var colIndex = _columns.IndexOf(columnName) + 1;
            var startCell = worksheet.Cells[topOffset + 2, colIndex] as Excel.Range;
            var endCell = worksheet.Cells[topOffset + 1 + Math.Max(NRows, 1), colIndex] as Excel.Range;
            return worksheet.Range[startCell, endCell];
        }

        internal void PrintTotalPriceTable(Excel.Worksheet worksheet, int topOffset)
        {
            // Index of 'Order' column in Excel's 'starting from 1 system'
            var orderColIndex = _columns.IndexOf(Order) + 1;

            var titleCell = worksheet.Cells[1, orderColIndex - 1];
            titleCell.Value2 = TotalOrder;
            Styling.Apply(titleCell, Styling.Style.Header);

            var unitsCell = worksheet.Cells[1, orderColIndex];
            Styling.Apply(unitsCell, Styling.Style.Calculation);
            unitsCell.Formula = "=SUM(" +
                                $"{orderColIndex.ToLetter()}{topOffset + 2}:" +
                                $"{orderColIndex.ToLetter()}{topOffset + 1 + NRows})";

            // Assumes that 'Total order' follows directly after 'Order'
            var totalPriceCell = worksheet.Cells[1, orderColIndex + 1];
            Styling.Apply(totalPriceCell, Styling.Style.Calculation);
            totalPriceCell.NumberFormat = AccountingFormat;
            totalPriceCell.Formula = "=SUM(" +
                                     $"{(orderColIndex + 1).ToLetter()}{topOffset + 2}:" +
                                     $"{(orderColIndex + 1).ToLetter()}{topOffset + 1 + NRows})";

            AddBorder(worksheet.Range[titleCell, totalPriceCell]);
        }

        /**
         * Assumes that 'Image' column is first.
         * Assumes 'Katalogové číslo' is translated as 'Item'.
         * Only one selection rule applies now:
         *  - image name == value in 'Item' column
         */
        internal async Task InsertImages(Excel.Worksheet worksheet, int topOffset, string imgFolder)
        {
            await Task.Run(() =>
            {
                const int defaultRowSize = 15;

                // image names are values in the 'Item' column
                var imgNames = Data
                    .Select(row => row[_columns.IndexOf(Item)] as string);

                var imgIdx = 0;
                foreach (var imgName in imgNames)
                {
                    if (FindImagePath(imgFolder, imgName, out var imgPath))
                    {
                        worksheet.Shapes.AddPicture(imgPath, MsoTriState.msoFalse, MsoTriState.msoCTrue, 0,
                            (topOffset + 1) * defaultRowSize + ImgColHeight * imgIdx + (ImgColHeight - ImgSize) / 2,
                            ImgSize, ImgSize);
                    }

                    imgIdx++;
                }
            });
        }

        /**
         * Returns true if image is found and sets imgPath.
         * Returns false if the image is not found, imgPath is set to empty string and should not be used.
         */
        private static bool FindImagePath(string imgFolder, string imgName, out string imgPath)
        {
            var extensions = new[] {"jpg", "png", "jpeg"};

            foreach (var extension in extensions)
            {
                var possiblePath = Path.Combine(imgFolder, $"{imgName}.{extension}");
                if (!File.Exists(possiblePath)) continue;
                imgPath = possiblePath;
                return true;
            }

            imgPath = "";
            return false;
        }

        internal void RemoveUnavailableProducts()
        {
            Data = Data
                .Where(row => !(
                    Convert.ToInt32(row[_columns.IndexOf(BudeKDispozici)]) == 0 &&
                    (Convert.ToString(row[_columns.IndexOf(UdajSklad1)]).Contains("ukončeno") ||
                     Convert.ToString(row[_columns.IndexOf(UdajSklad1)]).Contains("doprodej")
                    ) || Convert.ToString(row[_columns.IndexOf(UdajSklad1)]).Contains("POS")
                ))
                .ToJaggedArray();
        }

        internal void SelectColumns()
        {
            var allResultColumns = new List<string>
            {
                Image,
                Product,
                Item,
                Ean,
                Description,
                ColliPcsInCarton,
                New,
                ExwCz,
                Order,
                TotalOrder,
                Rrp,
                InStock,
                WillBeAvailable,
                StockComing,
                NoteForStock,
                Brand,
                Category,
                ProductType,
                Theme,
                CountryOfOrigin,
            };

            var availableResultColumns = allResultColumns
                .Where(col => _columns.IndexOf(col) != -1)
                .ToList();

            var newOrderOfIndices = availableResultColumns
                .Select(col => _columns.IndexOf(col));

            Data = Data
                .Select(row => newOrderOfIndices.Select(index => row[index]))
                .ToJaggedArray();

            _columns = availableResultColumns;
        }

        internal void RenameColumns()
        {
            var translationDict = new Dictionary<string, string>
            {
                {Produkt, Product},
                {KatalogoveCislo, Item},
                {PopisAlternativni, Description},
                {BaleníKartonKs, ColliPcsInCarton},
                {Cena, ExwCz},
                {CenaDmocEur, Rrp},
                {KDispozici, InStock},
                {BudeKDispozici, StockComing},
                {Vyrobce, Brand},
                {Udaj2, Category},
                {Udaj1, ProductType},
                {ZemePuvodu, CountryOfOrigin},
            };

            _columns = _columns.Select(col => translationDict.ContainsKey(col) ? translationDict[col] : col).ToList();
        }

        internal void InsertColumns()
        {
            InsertImageColumn();
            InsertNewColumn();
            InsertOrderColumn();
            InsertTotalOrderColumn();
            InsertNoteForStockColumn();
            InsertThemeColumn();
            InsertWillBeAvailableColumn();
        }

        /**
         * "Bude bude"
         */
        private void InsertWillBeAvailableColumn()
        {
            _columns.Add(WillBeAvailable);
            Data = Data
                .Select(row => row.Append(
                    Convert.ToInt32(row[_columns.IndexOf(BudeKDispozici)]) +
                    Convert.ToInt32(row[_columns.IndexOf(Objednano)]) -
                    Convert.ToInt32(row[_columns.IndexOf(Dodat)])
                ))
                .ToJaggedArray();
        }

        private void InsertThemeColumn()
        {
            InsertEmptyColumn(Theme);
        }

        private void InsertNoteForStockColumn()
        {
            InsertEmptyColumn(NoteForStock);
        }

        private void InsertTotalOrderColumn()
        {
            InsertEmptyColumn(TotalOrder);
        }

        private void InsertOrderColumn()
        {
            InsertEmptyColumn(Order);
        }

        private void InsertNewColumn()
        {
            InsertEmptyColumn(New);
        }

        private void InsertImageColumn()
        {
            InsertEmptyColumn(Image);
        }

        private void InsertEmptyColumn(string columnName)
        {
            _columns.Add(columnName);
            Data = Data
                .Select(row => row.Append(null))
                .ToJaggedArray();
        }

        internal static Table FromComboBoxes(ComboBox tableComboBox, ComboBox idColComboBox)
        {
            var worksheet = ((WorksheetItem) tableComboBox.SelectedItem).Worksheet;
            var idCol = idColComboBox.SelectedItem as string;

            var table = new Table
            {
                _idCol = idCol,
                _columns = worksheet.GetColumnNames()
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
