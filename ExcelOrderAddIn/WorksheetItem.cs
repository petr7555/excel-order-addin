using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelOrderAddIn
{
    class WorksheetItem
    {
        public Excel.Worksheet Worksheet { get; set; }
        public string Name { get => Worksheet.Name; }
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
