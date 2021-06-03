using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelOrderAddIn
{
    class WorksheetItem
    {
        public Excel.Worksheet Worksheet { get; set; }
        public string Name { get; private set; }
        public WorksheetItem(Excel.Worksheet worksheet)
        {
            Worksheet = worksheet;
            Name = Worksheet.Name;
        }

        public override string ToString()
        {
            return Name;
        }
    }
}
