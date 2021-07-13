using System.Collections.Generic;
using ExcelOrderAddIn.Displays;
using ExcelOrderAddIn.Logging;
using NUnit.Framework;
using ExcelOrderAddIn.Model;

namespace Tests
{
    public class TestSelectColumns
    {
        private static readonly ILogger Logger = new TestLogger();
        private static readonly IDisplay Display = new TestDisplay();

        [Test]
        public void SelectsColumns()
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

            var table = new Table(Logger, Display, columns, "Product", data);

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

        [Test]
        public void ReordersColumns()
        {
            var columns = new List<string>
            {
                "Stock coming",
                "Product",
                "Country of origin",
                "Brand",
            };

            var data = new[]
            {
                new object[] {"20", "Carlos", "China", "Squishmallows"},
                new object[] {"1", "Tatiana", "China", "Squishmallows"},
                new object[] {"0", "Henry", "China", "Squishmallows"},
                new object[] {"3", "Gordon", "China", "Squishmallows"},
            };

            var table = new Table(Logger, Display, columns, "Product", data);

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
