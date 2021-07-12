using System.Collections.Generic;
using ExcelOrderAddIn.Logging;
using NUnit.Framework;
using ExcelOrderAddIn.Model;

namespace Tests
{
    public class TestRemoveUnavailableProducts
    {
        private static readonly ILogger Logger = new TestLogger();

        private static readonly IList<string> Columns = new List<string>
        {
            "Produkt",
            "Bude k dispozici",
            "Údaj sklad 1",
        };

        [Test]
        public void RemovesUnavailableProductsWhereWillBeAvailableEqualsZeroAndNoteContainsEnded()
        {
            var data = new[]
            {
                new object[] {"Carlos", "0", "ukončeno"},
                new object[] {"Tatiana", "0", ""},
                new object[] {"Henry", "0", "ukončeno"},
                new object[] {"Gordon", "3", "ukončeno"},
            };

            var table = new Table(Logger, Columns, "Produkt", data);

            table.RemoveUnavailableProducts();

            var expected = new[]
            {
                new object[] {"Tatiana", "0", ""},
                new object[] {"Gordon", "3", "ukončeno"},
            };

            Assert.AreEqual(expected, table.Data);
        }

        [Test]
        public void RemovesUnavailableProductsWhereWillBeAvailableEqualsZeroAndNoteContainsClearance()
        {
            var data = new[]
            {
                new object[] {"Carlos", "0", "doprodej"},
                new object[] {"Tatiana", "0", ""},
                new object[] {"Henry", "0", "doprodej"},
                new object[] {"Gordon", "3", "doprodej"},
            };

            var table = new Table(Logger, Columns, "Produkt", data);

            table.RemoveUnavailableProducts();

            var expected = new[]
            {
                new object[] {"Tatiana", "0", ""},
                new object[] {"Gordon", "3", "doprodej"},
            };

            Assert.AreEqual(expected, table.Data);
        }

        [Test]
        public void RemovesUnavailableProductsWhereWillBeAvailableEqualsZeroAndNoteContainsPOS()
        {
            var data = new[]
            {
                new object[] {"Carlos", "0", "POS"},
                new object[] {"Tatiana", "0", ""},
                new object[] {"Henry", "0", "some note"},
                new object[] {"Gordon", "3", "POS"},
            };

            var table = new Table(Logger, Columns, "Produkt", data);

            table.RemoveUnavailableProducts();

            var expected = new[]
            {
                new object[] {"Tatiana", "0", ""},
                new object[] {"Henry", "0", "some note"},
            };

            Assert.AreEqual(expected, table.Data);
        }
    }
}
