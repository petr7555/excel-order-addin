using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelOrderAddIn
{
    class Column
    {
        public Excel.Worksheet Worksheet { get; set; }

        public Column(Excel.Worksheet worksheet)
        {
            Worksheet = worksheet;
        }

        public override string ToString()
        {
            return Worksheet.Name;
        }
    }
}
