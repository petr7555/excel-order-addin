using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using ExcelOrderAddIn.DateTime;

namespace ExcelOrderAddIn.Logging
{
    /**
     * Credits: https://stackoverflow.com/a/55540909/9290771
     */
    public class Logger : ILogger
    {
        private readonly Queue<LogEntry> _log;
        private uint _entryNumber;
        private readonly uint _maxEntries;
        private readonly object _logLock = new object();
        private readonly Color _defaultColor = Color.Black;
        private readonly IDateTime _dateTime;

        private class LogEntry
        {
            public uint EntryId;
            public System.DateTime EntryTimeStamp;
            public string EntryText;
            public Color EntryColor;
        }

        private struct ColorTableItem
        {
            public uint Index;
            public string RichColor;
        }

        /**
         * Create an instance of the Logger class which stores maximumEntries log entries.
         */
        public Logger(uint maximumEntries, IDateTime dateTime)
        {
            _log = new Queue<LogEntry>();
            _maxEntries = maximumEntries;
            _dateTime = dateTime;
        }

        /**
         * Create an instance of the Logger class which stores maximumEntries log entries.
         */
        public Logger(uint maximumEntries) : this(maximumEntries, new MyDateTime())
        {
        }

        /**
         * Retrieve the contents of the log as rich text, suitable for populating a System.Windows.Forms.RichTextBox.Rtf property.
         * includeEntryNumbers - option to prepend line numbers to each entry.
         */
        public string GetLogAsRichText(bool includeEntryNumbers)
        {
            lock (_logLock)
            {
                var sb = new StringBuilder();

                var uniqueColors = BuildRichTextColorTable();
                sb.AppendLine($@"{{\rtf1{{\colortbl;{string.Join("", uniqueColors.Select(d => d.Value.RichColor))}}}");

                foreach (var entry in _log)
                {
                    if (includeEntryNumbers)
                    {
                        sb.Append($"\\cf1 {entry.EntryId}. ");
                    }

                    sb.Append(
                        $"\\cf1 {entry.EntryTimeStamp.ToShortDateString()} {entry.EntryTimeStamp.ToLongTimeString()}: ");

                    var richColor = $"\\cf{uniqueColors[entry.EntryColor].Index + 1}";
                    sb.Append($"{richColor} {entry.EntryText}\\par").AppendLine();
                }

                return sb.ToString();
            }
        }

        /**
         * Adds text as a INFO log entry.
         */
        public void Info(string text)
        {
            AddToLog(text, _defaultColor);
        }

        /**
         * Adds text as an ERROR log entry.
         */
        public void Error(string text)
        {
            AddToLog(text, Color.Red);
        }

        /**
         * Adds text as a log entry, and specifies a color to display it in.
         */
        private void AddToLog(string text, Color entryColor)
        {
            lock (_logLock)
            {
                if (_entryNumber >= uint.MaxValue)
                {
                    _entryNumber = 0;
                }

                _entryNumber++;
                var logEntry = new LogEntry
                    {EntryId = _entryNumber, EntryTimeStamp = _dateTime.Now, EntryText = text, EntryColor = entryColor};
                _log.Enqueue(logEntry);

                if (_log.Count > _maxEntries)
                {
                    _log.Dequeue();
                }
            }
        }

        /**
         * Clears the entire log.
         */
        public void Clear()
        {
            lock (_logLock)
            {
                _log.Clear();
            }
        }

        private Dictionary<Color, ColorTableItem> BuildRichTextColorTable()
        {
            var uniqueColors = new Dictionary<Color, ColorTableItem>();
            var index = 0u;

            uniqueColors.Add(_defaultColor,
                new ColorTableItem {Index = index++, RichColor = ColorToRichColorString(_defaultColor)});

            foreach (var c in _log.Select(l => l.EntryColor).Distinct().Where(c => c != _defaultColor))
                uniqueColors.Add(c, new ColorTableItem {Index = index++, RichColor = ColorToRichColorString(c)});

            return uniqueColors;
        }

        private static string ColorToRichColorString(Color c)
        {
            return $"\\red{c.R}\\green{c.G}\\blue{c.B};";
        }
    }
}
