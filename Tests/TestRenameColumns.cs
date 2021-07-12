using System.Collections.Generic;
using System.IO;
using ExcelOrderAddIn.Logging;
using NUnit.Framework;
using ExcelOrderAddIn.Model;

namespace Tests
{
    public class TestRenameColumns
    {
        private static readonly ILogger Logger = new TestLogger();

        [Test]
        public void RenameColumns()
        {
            var originalColumns = new List<string>
            {
                "Produkt",
                "Bude k dispozici",
                "Bez překladu",
                "Výrobce",
            };
            
            var renamedColumns = new List<string>
            {
                "Product",
                "Stock coming",
                "Bez překladu",
                "Brand",
            };

            var table = new Table(Logger, originalColumns, "Produkt");

            table.RenameColumns();

            Assert.AreEqual(renamedColumns, table.Columns);
        }
    }
}
