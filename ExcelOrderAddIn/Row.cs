using System;
using System.Collections.Generic;

namespace ExcelOrderAddIn
{
    class Row
    {
        public Dictionary<string, object> ColumnsAndValues { get; set; } = new Dictionary<string, object>();

        public Row(List<string> columnNames, List<object> values)
        {
            if (columnNames.Count != values.Count)
            {
                throw new ArgumentException("Number of column names must be equal to the number of values in the row.");
            }

            for (int i = 0; i < columnNames.Count; i++)
            {
                var colName = columnNames[i];
                var value = values[i];

                if (!ColumnsAndValues.ContainsKey(colName))
                {
                    ColumnsAndValues.Add(colName, value);
                }
            }
        }
    }
}