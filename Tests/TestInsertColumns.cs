using System.Collections.Generic;
using System.Linq;
using ExcelOrderAddIn.Logging;
using NUnit.Framework;
using ExcelOrderAddIn.Model;
using Tests.Stubs;

namespace Tests
{
    public class TestInsertColumns
    {
        private static readonly ILogger Logger = new LoggerForTests();
        private static readonly TestDisplay Display = new TestDisplay();

        [Test]
        public void InsertsColumnsIncludingWillBeAvailableColumn()
        {
            var columns = new List<string>
            {
                "Produkt",
                "Bude k dispozici",
                "OBJEDNÁNO",
                "DODAT"
            };

            var data = new[]
            {
                new object[] {"Carlos", "20", "5", "3"},
                new object[] {"Tatiana", "1", "0", "3"},
                new object[] {"Henry", "0", "2", "0"},
                new object[] {"Gordon", "3", "5", "3"},
            };

            var table = new Table(Logger, Display, columns, "Produkt", data);

            table.InsertColumns();

            var expectedData = new[]
            {
                new object[] {"Carlos", "20", "5", "3", null, null, null, null, null, null, 20 + 5 - 3},
                new object[] {"Tatiana", "1", "0", "3", null, null, null, null, null, null, 1 + 0 - 3},
                new object[] {"Henry", "0", "2", "0", null, null, null, null, null, null, 0 + 2 - 0},
                new object[] {"Gordon", "3", "5", "3", null, null, null, null, null, null, 3 + 5 - 3},
            };

            var expectedColumns = new List<string>
            {
                "Produkt",
                "Bude k dispozici",
                "OBJEDNÁNO",
                "DODAT",
                "Image",
                "NEW",
                "Order",
                "Total order",
                "Note for stock",
                "Theme",
                "Will be available",
            };

            Assert.AreEqual(expectedData, table.Data);
            Assert.AreEqual(expectedColumns, table.Columns);
        }

        private static void DoesNotInsertWillBeAvailableColumnWhenUnderlyingColumnIsMissing(IList<string> columns,
            string missingColumnName)
        {
            var data = new[]
            {
                new object[] {"Carlos", "5", "3"},
                new object[] {"Tatiana", "0", "3"},
                new object[] {"Henry", "2", "0"},
                new object[] {"Gordon", "5", "3"},
            };

            var table = new Table(Logger, Display, columns, "Produkt", data);

            table.InsertColumns();

            var expectedData = new[]
            {
                new object[] {"Carlos", "5", "3", null, null, null, null, null, null},
                new object[] {"Tatiana", "0", "3", null, null, null, null, null, null},
                new object[] {"Henry", "2", "0", null, null, null, null, null, null},
                new object[] {"Gordon", "5", "3", null, null, null, null, null, null},
            };

            var expectedColumns = columns.Concat(new List<string>
            {
                "Image",
                "NEW",
                "Order",
                "Total order",
                "Note for stock",
                "Theme",
            }).ToList();

            Assert.AreEqual(expectedData, table.Data);
            Assert.AreEqual(expectedColumns, table.Columns);
            Assert.AreEqual(
                $"Data do not contain \"{missingColumnName}\" column, \"Will be available column\" won't be added.",
                Display.LastDisplayedMessage);
        }

        [Test]
        public void DoesNotInsertWillBeAvailableColumnWhenWillBeAvailableColumnIsMissing()
        {
            var columns = new List<string>
            {
                "Produkt",
                "OBJEDNÁNO",
                "DODAT"
            };

            DoesNotInsertWillBeAvailableColumnWhenUnderlyingColumnIsMissing(columns, "Bude k dispozici");
        }

        [Test]
        public void DoesNotInsertWillBeAvailableColumnWhenOrderedColumnIsMissing()
        {
            var columns = new List<string>
            {
                "Produkt",
                "Bude k dispozici",
                "DODAT"
            };

            DoesNotInsertWillBeAvailableColumnWhenUnderlyingColumnIsMissing(columns, "OBJEDNÁNO");
        }

        [Test]
        public void DoesNotInsertWillBeAvailableColumnWhenDeliverColumnIsMissing()
        {
            var columns = new List<string>
            {
                "Produkt",
                "Bude k dispozici",
                "OBJEDNÁNO"
            };

            DoesNotInsertWillBeAvailableColumnWhenUnderlyingColumnIsMissing(columns, "DODAT");
        }
    }
}
