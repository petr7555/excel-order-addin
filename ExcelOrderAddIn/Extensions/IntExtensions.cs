using System;

namespace ExcelOrderAddIn.Extensions
{
    public static class IntExtensions
    {
        /**
         * Convert column number (starting from 1) to Excel's column letter numbering,
         * i.e. 1 -> A, 27 -> AA.
         * Credits to https://stackoverflow.com/a/182924.
         */
        public static string ToLetter(this int columnNumber)
        {
            var dividend = columnNumber;
            var columnName = "";

            while (dividend > 0)
            {
                var modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }
    }
}
