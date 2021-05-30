using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelOrderAddIn
{
    class WorksheetItem
    {
        public Excel.Worksheet Worksheet { get; set; }

        public WorksheetItem(Excel.Worksheet worksheet)
        {
            Worksheet = worksheet;
        }

        public override string ToString()
        {
            return Worksheet.Name;
        }
    }
}
