namespace ExcelOrderAddIn.Logging
{
    public interface ILogger
    {
        string GetLogAsRichText(bool includeEntryNumbers);
        void Info(string text);
        void Error(string text);
        void Clear();
    }
}
