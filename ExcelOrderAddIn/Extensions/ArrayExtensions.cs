using System.Collections.Generic;
using System.Linq;

namespace ExcelOrderAddIn
{
    public static class ArrayExtensions
    {
        private static IEnumerable<T> GetColumn<T>(this T[,] array, int column)
        {
            for (var i = 0; i < array.GetLength(0); i++)
            {
                yield return array[i, column];
            }
        }

        public static IEnumerable<IEnumerable<T>> GetColumns<T>(this T[,] array)
        {
            for (var i = 0; i < array.GetLength(1); i++)
            {
                yield return array.GetColumn(i);
            }
        }

        /**
         * Excel arrays start at 1.
         * Convert into normal matrix starting at 0.
         */
        public static T[][] FromExcelMultidimArray<T>(this T[,] array)
        {
            var rows = array.GetLength(0);
            var cols = array.GetLength(1);

            var result = new T[rows][];
            for (var i = 0; i < rows; i++)
            {
                result[i] = new T[cols];

                for (var j = 0; j < cols; j++)
                {
                    result[i][j] = array[i + 1, j + 1];
                }
            }

            return result;
        }

        /**
         * Values do not have to start at 1, as opposed to the data read.
         */
        public static T[,] ToExcelMultidimArray<T>(this T[][] array)
        {
            var rows = array.GetLength(0);

            if (rows == 0)
            {
                return new T[0, 0];
            }

            var cols = array[0].GetLength(0);

            var result = new T[rows, cols];

            for (var i = 0; i < rows; i++)
            {
                for (var j = 0; j < cols; j++)
                {
                    result[i, j] = array[i][j];
                }
            }

            return result;
        }

        public static T[,] ToExcelMultidimArray<T>(this IList<T> array)
        {
            var cols = array.Count;

            var result = new T[1, cols];

            for (var i = 0; i < cols; i++)
            {
                result[0, i] = array[i];
            }

            return result;
        }

        public static T[][] ToJaggedArray<T>(this IEnumerable<IEnumerable<T>> array)
        {
            return array.Select(row => row.ToArray()).ToArray();
        }
    }
}
