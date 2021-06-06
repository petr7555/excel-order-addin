using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelOrderAddIn.Model
{
    internal class WorksheetItem
    {
        public Excel.Worksheet Worksheet { get; }
        public string Name { get; }

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
