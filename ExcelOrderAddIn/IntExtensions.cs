using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelOrderAddIn
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
            int dividend = columnNumber;
            string columnName = "";
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }
    }
}
