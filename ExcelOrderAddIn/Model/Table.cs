// ReSharper disable once RedundantUsingDirective

using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelOrderAddIn.Exceptions;
using ExcelOrderAddIn.Extensions;
using ExcelOrderAddIn.Logging;
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
        private const string Popis = "Popis";
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
        private const string Description2 = "Description 2";
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
        internal object[][] Data = new object[0][];
        private readonly string _idCol;
        private readonly ILogger _logger;

        private int NCols => _columns.Count;

        private int NRows => Data.Length;

        private int IdColIdx => _columns.IndexOf(_idCol);

        /**
         * internal and not private for tests
         */
        internal Table(ILogger logger, IList<string> columns, string idCol)
        {
            _logger = logger;
            _columns = columns;
            _idCol = idCol;
        }

        /**
         * internal and not private for tests
         */
        internal Table(ILogger logger, IList<string> columns, string idCol, object[][] data) : this(logger, columns,
            idCol)
        {
            Data = data;
        }

        internal Table Join(Table rightTable)
        {
            var leftIdColIdx = IdColIdx;
            var rightIdColIdx = rightTable.IdColIdx;

            // Join data on id columns, merge all columns (including the id column)
            var joinedData = Data
                .Join(rightTable.Data,
                    leftRow => leftRow.ElementAt(leftIdColIdx).ToString(),
                    rightRow => rightRow.ElementAt(rightIdColIdx).ToString(),
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

            return new Table(_logger, newCols, _idCol, filteredData.ToJaggedArray());
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
            if (ColumnIsMissing(WillBeAvailable)) return;

            ApplyStyleToColumn(worksheet, topOffset, Styling.Style.Yellow, WillBeAvailable);
            FormatColumnAsInteger(worksheet, topOffset, StockComing);
        }

        private bool ColumnIsMissing(string columnName)
        {
            return _columns.IndexOf(columnName) == -1;
        }

        private void FormatInStockColumn(Excel.Worksheet worksheet, int topOffset)
        {
            if (ColumnIsMissing(InStock)) return;

            ApplyStyleToColumn(worksheet, topOffset, Styling.Style.SalmonBold, InStock);
            FormatColumnAsInteger(worksheet, topOffset, InStock);
        }

        private void FormatRrpColumn(Excel.Worksheet worksheet, int topOffset)
        {
            if (ColumnIsMissing(Rrp)) return;

            ApplyStyleToColumn(worksheet, topOffset, Styling.Style.BoldText, Rrp);
            FormatColumnAsAccounting(worksheet, topOffset, Rrp);
        }

        private void FormatExwCzColumn(Excel.Worksheet worksheet, int topOffset)
        {
            if (ColumnIsMissing(ExwCz)) return;

            FormatColumnAsAccounting(worksheet, topOffset, ExwCz);
        }

        private void FormatNewColumn(Excel.Worksheet worksheet, int topOffset)
        {
            if (ColumnIsMissing(New)) return;

            ApplyStyleToColumn(worksheet, topOffset, Styling.Style.RedBoldText, New);

            var colIndex = GetColumnIndex(New) + 1;
            var headerCell = worksheet.Cells[topOffset + 1, colIndex];
            Styling.Apply(headerCell, Styling.Style.RedBoldHeaderText);
        }

        private void FormatColliColumn(Excel.Worksheet worksheet, int topOffset)
        {
            if (ColumnIsMissing(ColliPcsInCarton)) return;

            FormatColumnAsInteger(worksheet, topOffset, ColliPcsInCarton);
        }

        private void FormatNoteForStockColumn(Excel.Worksheet worksheet, int topOffset)
        {
            if (ColumnIsMissing(NoteForStock)) return;

            ApplyStyleToColumn(worksheet, topOffset, Styling.Style.Yellow, NoteForStock);
            FormatColumnAsText(worksheet, topOffset, NoteForStock);
        }

        private void FormatStockComingColumn(Excel.Worksheet worksheet, int topOffset)
        {
            if (ColumnIsMissing(StockComing)) return;

            ApplyStyleToColumn(worksheet, topOffset, Styling.Style.Yellow, StockComing);
            FormatColumnAsInteger(worksheet, topOffset, StockComing);
        }

        private void FormatImageColumn(Excel.Worksheet worksheet)
        {
            if (ColumnIsMissing(Image)) return;

            // Make column wider
            worksheet.Columns[1].ColumnWidth = ImgColWidth;
        }

        private void FormatOrderColumn(Excel.Worksheet worksheet, int topOffset)
        {
            if (ColumnIsMissing(Order)) return;

            ApplyStyleToColumn(worksheet, topOffset, Styling.Style.Input, Order);
        }

        private void FormatEanColumn(Excel.Worksheet worksheet, int topOffset)
        {
            if (ColumnIsMissing(Ean)) return;

            FormatColumnAsInteger(worksheet, topOffset, Ean);
        }

        private void FormatTotalOrderColumn(Excel.Worksheet worksheet, int topOffset)
        {
            if (ColumnIsMissing(TotalOrder)) return;

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
            // Further methods rely on some columns to exist.
            // Make sure the column is not needed for other methods before making it optional.
            // Or handle its absence in the other methods.
            // TODO Could be configurable
            var importanceDict = new Dictionary<string, ColumnImportance>
            {
                // MANDATORY
                {Produkt, ColumnImportance.Mandatory},
                {KatalogoveCislo, ColumnImportance.Mandatory},
                {Cena, ColumnImportance.Mandatory},
                {CenaDmocEur, ColumnImportance.Mandatory},
                {KDispozici, ColumnImportance.Mandatory},
                {Vyrobce, ColumnImportance.Mandatory},
                {Udaj1, ColumnImportance.Mandatory},
                {Udaj2, ColumnImportance.Mandatory},
                // OPTIONAL
                {UdajSklad1, ColumnImportance.Optional},
                {PopisAlternativni, ColumnImportance.Optional},
                {Popis, ColumnImportance.Optional},
                {BaleníKartonKs, ColumnImportance.Optional},
                {BudeKDispozici, ColumnImportance.Optional},
                {Objednano, ColumnImportance.Optional},
                {Dodat, ColumnImportance.Optional},
                {ZemePuvodu, ColumnImportance.Optional},
            };

            var notFoundColumns = importanceDict
                .Where(x => _columns.IndexOf(x.Key) == -1 && x.Value == ColumnImportance.Mandatory).ToList();
            if (notFoundColumns.Count > 0)
            {
                throw new InvalidDataException(
                    $"Data do not contain the following columns: {string.Join(", ", notFoundColumns.Select(col => col.Key))}.");
            }
        }

        private void InsertTotalOrderFormula(Excel.Worksheet worksheet, int topOffset)
        {
            var totalOrderIndex = GetColumnIndex(TotalOrder) + 1;
            var priceColLetter = (GetColumnIndex(ExwCz) + 1).ToLetter();
            var orderColLetter = (GetColumnIndex(Order) + 1).ToLetter();

            Parallel.For(0, NRows, i =>
            {
                var row = topOffset + 2 + i;
                worksheet.Cells[row, totalOrderIndex].Formula =
                    $"={priceColLetter}{row}*" +
                    $"{orderColLetter}{row}";
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
            var colIndex = GetColumnIndex(columnName) + 1;
            var startCell = worksheet.Cells[topOffset + 2, colIndex] as Excel.Range;
            var endCell = worksheet.Cells[topOffset + 1 + Math.Max(NRows, 1), colIndex] as Excel.Range;
            return worksheet.Range[startCell, endCell];
        }

        internal void PrintTotalPriceTable(Excel.Worksheet worksheet, int topOffset)
        {
            // Index of 'Order' column in Excel's 'starting from 1 system'
            var orderColIndex = GetColumnIndex(Order) + 1;

            var titleCell = worksheet.Cells[1, orderColIndex - 1];
            titleCell.Value2 = TotalOrder;
            Styling.Apply(titleCell, Styling.Style.Header);

            var unitsCell = worksheet.Cells[1, orderColIndex];
            Styling.Apply(unitsCell, Styling.Style.Calculation);
            unitsCell.Formula = "=SUM(" +
                                $"{orderColIndex.ToLetter()}{topOffset + 2}:" +
                                $"{orderColIndex.ToLetter()}{topOffset + 1 + NRows})";

            // Assumes that 'Total order' follows directly after 'Order'
            // Casting to Excel.Range makes AccountingFormat work
            var totalPriceCell = worksheet.Cells[1, orderColIndex + 1] as Excel.Range;
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

                // For some reason, row height is actually smaller than what is set.
                // E.g. when row height is 76, the second image needs to start at 75.75 (76 - 0.25) from top, instead of 76.
                const float weirdExcelShift = 0.25f;

                // image names are values in the 'Item' column
                var imgNames = Data
                    .Select(row => row[GetColumnIndex(Item)].ToString());

                var imgIdx = 0;
                foreach (var imgName in imgNames)
                {
                    if (FindImagePath(imgFolder, imgName, out var imgPath))
                    {
                        worksheet.Shapes.AddPicture(imgPath, MsoTriState.msoFalse, MsoTriState.msoCTrue, 0,
                            (topOffset + 1) * defaultRowSize + (ImgColHeight - weirdExcelShift) * imgIdx +
                            (ImgColHeight - ImgSize) / 2,
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

        private int GetColumnIndex(string columnName)
        {
            var columnIdx = _columns.IndexOf(columnName);
            if (columnIdx == -1)
            {
                throw new ProgrammerErrorException(
                    $"Data do not contain \"{columnName}\" column, this should have been checked before.");
            }

            return columnIdx;
        }

        internal void RemoveUnavailableProducts()
        {
            if (WarnIfColumnIsMissing(BudeKDispozici, "unavailable products won't be removed")) return;
            if (WarnIfColumnIsMissing(UdajSklad1, "unavailable products won't be removed")) return;

            var budeKDispoziciIdx = GetColumnIndex(BudeKDispozici);
            var udajSklad1Idx = GetColumnIndex(UdajSklad1);

            Data = Data
                .Where(row => !(
                    (Convert.ToInt32(row[budeKDispoziciIdx]) == 0 &&
                     (Convert.ToString(row[udajSklad1Idx]).Contains("ukončeno") ||
                      Convert.ToString(row[udajSklad1Idx]).Contains("doprodej")
                     )) || Convert.ToString(row[udajSklad1Idx]).Contains("POS")
                ))
                .ToJaggedArray();
        }

        /**
         * Selects columns that should be in the final order.
         * Unavailable columns are skipped.
         */
        internal void SelectColumns()
        {
            // TODO Could be configurable
            var allResultColumns = new List<string>
            {
                Image,
                Product,
                Item,
                Ean,
                Description,
                Description2,
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
            // TODO Could be configurable
            var translationDict = new Dictionary<string, string>
            {
                {Produkt, Product},
                {KatalogoveCislo, Item},
                {PopisAlternativni, Description},
                {Popis, Description2},
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

        private bool WarnIfColumnIsMissing(string columnName, string effect)
        {
            if (!ColumnIsMissing(columnName)) return false;
            MessageBox.Show(
                $"Data do not contain \"{columnName}\" column, {effect}.",
                "Missing column", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return true;
        }

        /**
         * So called "Bude bude" column.
         */
        private void InsertWillBeAvailableColumn()
        {
            bool WarnIfColumnIsMissingWithMsg(string columnName)
            {
                return WarnIfColumnIsMissing(columnName, "\"Will be available column\" won't be added");
            }

            if (WarnIfColumnIsMissingWithMsg(BudeKDispozici)) return;
            if (WarnIfColumnIsMissingWithMsg(Objednano)) return;
            if (WarnIfColumnIsMissingWithMsg(Dodat)) return;

            _columns.Add(WillBeAvailable);
            Data = Data
                .Select(row => row.Append(
                    Convert.ToInt32(row[GetColumnIndex(BudeKDispozici)]) +
                    Convert.ToInt32(row[GetColumnIndex(Objednano)]) -
                    Convert.ToInt32(row[GetColumnIndex(Dodat)])
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

        internal static Table FromComboBoxes(ILogger logger, ComboBox tableComboBox, ComboBox idColComboBox)
        {
            var worksheet = ((WorksheetItem) tableComboBox.SelectedItem).Worksheet;
            var idCol = idColComboBox.SelectedItem as string;

            var table = new Table
            (
                logger,
                worksheet.GetColumnNames(),
                idCol
            );

            var nCols = worksheet.NCols();
            var nRows = worksheet.NRows();

            if (table.NCols == 0 || nRows == 0)
            {
                return table;
            }

            // skip input table header
            var dataStartCell = worksheet.Cells[2, 1] as Excel.Range;
            var dataEndCell = worksheet.Cells[nRows + 1, nCols] as Excel.Range;
            table.Data = (worksheet.Range[dataStartCell, dataEndCell].Value2 as object[,]).FromExcelMultidimArray();

            return table;
        }
    }
}
