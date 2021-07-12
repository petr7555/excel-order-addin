using ExcelOrderAddIn.Logging;

namespace Tests
{
    public class TestLogger : ILogger
    {
        public string GetLogAsRichText(bool includeEntryNumbers)
        {
            return "test";
        }

        public void Info(string text)
        {
        }

        public void Error(string text)
        {
        }

        public void Clear()
        {
        }
    }
}
