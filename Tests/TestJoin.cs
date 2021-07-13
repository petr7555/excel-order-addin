using System.Collections.Generic;
using ExcelOrderAddIn.Displays;
using ExcelOrderAddIn.Logging;
using NUnit.Framework;
using ExcelOrderAddIn.Model;
using Tests.Stubs;

namespace Tests
{
    public class TestJoin
    {
        private static readonly ILogger Logger = new TestLogger();
        private static readonly IDisplay Display = new TestDisplay();

        [Test]
        public void JoinTables()
        {
            var columnsLeft = new List<string>
            {
                "A",
                "B",
                "C",
            };

            var columnsRight = new List<string>
            {
                "D",
                "E",
            };

            var dataLeft = new[]
            {
                new object[] {"Carlos", "B1", "C1"},
                new object[] {"Tatiana", "B2", "C2"},
                new object[] {"Henry", "B3", "C3"},
                new object[] {"Gordon", "B4", "C4"},
            };

            var dataRight = new[]
            {
                new object[] {"D1", "Carlos"},
                new object[] {"D2", "Tatiana"},
                new object[] {"D3", "Henry"},
                new object[] {"D4", "Gordon"},
            };

            var tableLeft = new Table(Logger, Display, columnsLeft, "A", dataLeft);
            var tableRight = new Table(Logger, Display, columnsRight, "E", dataRight);

            var joinedTable = tableLeft.Join(tableRight);

            var expectedData = new[]
            {
                new object[] {"Carlos", "B1", "C1", "D1", "Carlos"},
                new object[] {"Tatiana", "B2", "C2", "D2", "Tatiana"},
                new object[] {"Henry", "B3", "C3", "D3", "Henry"},
                new object[] {"Gordon", "B4", "C4", "D4", "Gordon"},
            };

            var expectedColumns = new List<string>
            {
                "A",
                "B",
                "C",
                "D",
                "E"
            };

            Assert.AreEqual(expectedData, joinedTable.Data);
            Assert.AreEqual(expectedColumns, joinedTable.Columns);
        }

        [Test]
        public void JoinRemovesIdColWhenSameName()
        {
            var columnsLeft = new List<string>
            {
                "A",
                "B",
                "C",
            };

            var columnsRight = new List<string>
            {
                "A",
                "D",
            };

            var dataLeft = new[]
            {
                new object[] {"Carlos", "B1", "C1"},
                new object[] {"Tatiana", "B2", "C2"},
                new object[] {"Henry", "B3", "C3"},
                new object[] {"Gordon", "B4", "C4"},
            };

            var dataRight = new[]
            {
                new object[] {"Carlos", "D1"},
                new object[] {"Tatiana", "D2"},
                new object[] {"Henry", "D3"},
                new object[] {"Gordon", "D4"},
            };

            var tableLeft = new Table(Logger, Display, columnsLeft, "A", dataLeft);
            var tableRight = new Table(Logger, Display, columnsRight, "A", dataRight);

            var joinedTable = tableLeft.Join(tableRight);

            var expectedData = new[]
            {
                new object[] {"Carlos", "B1", "C1", "D1"},
                new object[] {"Tatiana", "B2", "C2", "D2"},
                new object[] {"Henry", "B3", "C3", "D3"},
                new object[] {"Gordon", "B4", "C4", "D4"},
            };

            var expectedColumns = new List<string>
            {
                "A",
                "B",
                "C",
                "D",
            };

            Assert.AreEqual(expectedData, joinedTable.Data);
            Assert.AreEqual(expectedColumns, joinedTable.Columns);
        }

        [Test]
        public void JoinRemovesColumnsOfSameName()
        {
            var columnsLeft = new List<string>
            {
                "A",
                "B",
                "C",
            };

            var columnsRight = new List<string>
            {
                "D",
                "B",
            };

            var dataLeft = new[]
            {
                new object[] {"Carlos", "B1", "C1"},
                new object[] {"Tatiana", "B2", "C2"},
                new object[] {"Henry", "B3", "C3"},
                new object[] {"Gordon", "B4", "C4"},
            };

            var dataRight = new[]
            {
                new object[] {"Carlos", "D1"},
                new object[] {"Tatiana", "D2"},
                new object[] {"Henry", "D3"},
                new object[] {"Gordon", "D4"},
            };

            var tableLeft = new Table(Logger, Display, columnsLeft, "A", dataLeft);
            var tableRight = new Table(Logger, Display, columnsRight, "D", dataRight);

            var joinedTable = tableLeft.Join(tableRight);

            var expectedData = new[]
            {
                new object[] {"Carlos", "B1", "C1", "Carlos"},
                new object[] {"Tatiana", "B2", "C2", "Tatiana"},
                new object[] {"Henry", "B3", "C3", "Henry"},
                new object[] {"Gordon", "B4", "C4", "Gordon"},
            };

            var expectedColumns = new List<string>
            {
                "A",
                "B",
                "C",
                "D",
            };

            Assert.AreEqual(expectedData, joinedTable.Data);
            Assert.AreEqual(expectedColumns, joinedTable.Columns);
        }

        [Test]
        public void JoinWithMultipleInRightTable()
        {
            var columnsLeft = new List<string>
            {
                "A",
                "B",
                "C",
            };

            var columnsRight = new List<string>
            {
                "D",
                "E",
            };

            var dataLeft = new[]
            {
                new object[] {"Carlos", "B1", "C1"},
                new object[] {"Tatiana", "B2", "C2"},
                new object[] {"Henry", "B3", "C3"},
                new object[] {"Gordon", "B4", "C4"},
            };

            var dataRight = new[]
            {
                new object[] {"D1", "Carlos"},
                new object[] {"D2", "Tatiana"},
                new object[] {"D3", "Carlos"},
                new object[] {"D4", "Tatiana"},
            };

            var tableLeft = new Table(Logger, Display, columnsLeft, "A", dataLeft);
            var tableRight = new Table(Logger, Display, columnsRight, "E", dataRight);

            var joinedTable = tableLeft.Join(tableRight);

            var expectedData = new[]
            {
                new object[] {"Carlos", "B1", "C1", "D1", "Carlos"},
                new object[] {"Carlos", "B1", "C1", "D3", "Carlos"},
                new object[] {"Tatiana", "B2", "C2", "D2", "Tatiana"},
                new object[] {"Tatiana", "B2", "C2", "D4", "Tatiana"},
            };

            var expectedColumns = new List<string>
            {
                "A",
                "B",
                "C",
                "D",
                "E"
            };

            Assert.AreEqual(expectedData, joinedTable.Data);
            Assert.AreEqual(expectedColumns, joinedTable.Columns);
        }

        [Test]
        public void JoinWithMultipleInLeftTable()
        {
            var columnsLeft = new List<string>
            {
                "A",
                "B",
                "C",
            };

            var columnsRight = new List<string>
            {
                "D",
                "E",
            };

            var dataLeft = new[]
            {
                new object[] {"Carlos", "B1", "C1"},
                new object[] {"Tatiana", "B2", "C2"},
                new object[] {"Carlos", "B3", "C3"},
                new object[] {"Gordon", "B4", "C4"},
            };

            var dataRight = new[]
            {
                new object[] {"Carlos", "D1"},
                new object[] {"Tatiana", "D2"},
                new object[] {"Henry", "D3"},
                new object[] {"Gordon", "D4"},
            };

            var tableLeft = new Table(Logger, Display, columnsLeft, "A", dataLeft);
            var tableRight = new Table(Logger, Display, columnsRight, "D", dataRight);

            var joinedTable = tableLeft.Join(tableRight);

            var expectedData = new[]
            {
                new object[] {"Carlos", "B1", "C1", "Carlos", "D1"},
                new object[] {"Tatiana", "B2", "C2", "Tatiana", "D2"},
                new object[] {"Carlos", "B3", "C3", "Carlos", "D1"},
                new object[] {"Gordon", "B4", "C4", "Gordon", "D4"},
            };

            var expectedColumns = new List<string>
            {
                "A",
                "B",
                "C",
                "D",
                "E"
            };

            Assert.AreEqual(expectedData, joinedTable.Data);
            Assert.AreEqual(expectedColumns, joinedTable.Columns);
        }

        [Test]
        public void JoinWithMultipleInLeftAndRightTable()
        {
            var columnsLeft = new List<string>
            {
                "A",
                "B",
                "C",
            };

            var columnsRight = new List<string>
            {
                "D",
                "E",
            };

            var dataLeft = new[]
            {
                new object[] {"Carlos", "B1", "C1"},
                new object[] {"Tatiana", "B2", "C2"},
                new object[] {"Carlos", "B3", "C3"},
                new object[] {"Gordon", "B4", "C4"},
            };

            var dataRight = new[]
            {
                new object[] {"Carlos", "D1"},
                new object[] {"Tatiana", "D2"},
                new object[] {"Carlos", "D3"},
                new object[] {"Gordon", "D4"},
            };

            var tableLeft = new Table(Logger, Display, columnsLeft, "A", dataLeft);
            var tableRight = new Table(Logger, Display, columnsRight, "D", dataRight);

            var joinedTable = tableLeft.Join(tableRight);

            var expectedData = new[]
            {
                new object[] {"Carlos", "B1", "C1", "Carlos", "D1"},
                new object[] {"Carlos", "B1", "C1", "Carlos", "D3"},
                new object[] {"Tatiana", "B2", "C2", "Tatiana", "D2"},
                new object[] {"Carlos", "B3", "C3", "Carlos", "D1"},
                new object[] {"Carlos", "B3", "C3", "Carlos", "D3"},
                new object[] {"Gordon", "B4", "C4", "Gordon", "D4"},
            };

            var expectedColumns = new List<string>
            {
                "A",
                "B",
                "C",
                "D",
                "E"
            };

            Assert.AreEqual(expectedData, joinedTable.Data);
            Assert.AreEqual(expectedColumns, joinedTable.Columns);
        }

        [Test]
        public void JoinEmptyLeftTable()
        {
            var columnsLeft = new List<string>
            {
                "A",
                "B",
                "C",
            };

            var columnsRight = new List<string>
            {
                "D",
                "E",
            };

            var dataLeft = new object[][] { };

            var dataRight = new[]
            {
                new object[] {"Carlos", "D1"},
                new object[] {"Tatiana", "D2"},
                new object[] {"Henry", "D3"},
            };

            var tableLeft = new Table(Logger, Display, columnsLeft, "A", dataLeft);
            var tableRight = new Table(Logger, Display, columnsRight, "D", dataRight);

            var joinedTable = tableLeft.Join(tableRight);

            var expectedData = new object[][] { };

            var expectedColumns = new List<string>
            {
                "A",
                "B",
                "C",
                "D",
                "E"
            };

            Assert.AreEqual(expectedData, joinedTable.Data);
            Assert.AreEqual(expectedColumns, joinedTable.Columns);
        }

        [Test]
        public void JoinEmptyRightTable()
        {
            var columnsLeft = new List<string>
            {
                "A",
                "B",
                "C",
            };

            var columnsRight = new List<string>
            {
                "D",
                "E",
            };

            var dataLeft = new[]
            {
                new object[] {"Carlos", "B1", "C1"},
                new object[] {"Tatiana", "B2", "C2"},
                new object[] {"Henry", "B3", "C3"},
            };

            var dataRight = new object[][] { };

            var tableLeft = new Table(Logger, Display, columnsLeft, "A", dataLeft);
            var tableRight = new Table(Logger, Display, columnsRight, "D", dataRight);

            var joinedTable = tableLeft.Join(tableRight);

            var expectedData = new object[][] { };

            var expectedColumns = new List<string>
            {
                "A",
                "B",
                "C",
                "D",
                "E"
            };

            Assert.AreEqual(expectedData, joinedTable.Data);
            Assert.AreEqual(expectedColumns, joinedTable.Columns);
        }
        
        [Test]
        public void JoinNoMatch()
        {
            var columnsLeft = new List<string>
            {
                "A",
                "B",
                "C",
            };

            var columnsRight = new List<string>
            {
                "D",
                "E",
            };

            var dataLeft = new[]
            {
                new object[] {"Carlos", "B1", "C1"},
                new object[] {"Tatiana", "B2", "C2"},
            };

            var dataRight = new[]
            {
                new object[] {"Henry", "D3"},
                new object[] {"Gordon", "D4"},
            };
            
            var tableLeft = new Table(Logger, Display, columnsLeft, "A", dataLeft);
            var tableRight = new Table(Logger, Display, columnsRight, "D", dataRight);

            var joinedTable = tableLeft.Join(tableRight);

            var expectedData = new object[][] { };

            var expectedColumns = new List<string>
            {
                "A",
                "B",
                "C",
                "D",
                "E"
            };

            Assert.AreEqual(expectedData, joinedTable.Data);
            Assert.AreEqual(expectedColumns, joinedTable.Columns);
        }
    }
}
