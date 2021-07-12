using System.Collections.Generic;
using System.IO;
using ExcelOrderAddIn.Logging;
using NUnit.Framework;
using ExcelOrderAddIn.Model;

namespace Tests
{
    public class TestSelectColumns
    {
        private static readonly ILogger Logger = new TestLogger();

        [Test]
        public void SelectColumns()
        {
            var columns = new List<string>
            {
                "Product",
                "Stock coming",
                "Some extra column",
                "Brand",
                "Country of origin",
                "Another extra column",
            };

            var data = new[]
            {
                new object[] {"Carlos", "20", "extra", "Squishmallows", "China", "another extra info"},
                new object[] {"Tatiana", "1", "extra", "Squishmallows", "China", "another extra info"},
                new object[] {"Henry", "0", "extra", "Squishmallows", "China", "another extra info"},
                new object[] {"Gordon", "3", "extra", "Squishmallows", "China", "another extra info"},
            };

            var table = new Table(Logger, columns, "Product", data);

            table.SelectColumns();

            var expectedData = new[]
            {
                new object[] {"Carlos", "20", "Squishmallows", "China"},
                new object[] {"Tatiana", "1", "Squishmallows", "China"},
                new object[] {"Henry", "0", "Squishmallows", "China"},
                new object[] {"Gordon", "3", "Squishmallows", "China"},
            };

            var expectedColumns = new List<string>
            {
                "Product",
                "Stock coming",
                "Brand",
                "Country of origin",
            };
            
            Assert.AreEqual(expectedData, table.Data);
            Assert.AreEqual(expectedColumns, table.Columns);
        }
    }
}
