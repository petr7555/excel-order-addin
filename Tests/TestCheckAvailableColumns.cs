using System.Collections.Generic;
using System.IO;
using ExcelOrderAddIn.Logging;
using NUnit.Framework;
using ExcelOrderAddIn.Model;

namespace Tests
{
    public class Tests
    {
        private static readonly ILogger Logger = new TestLogger();

        [SetUp]
        public void Setup()
        {
        }

        // CheckAvailableColumns
        // RemoveUnavailableProducts
        // Do nabídky se nemají dostat produkty, u kterých:
        //
        // „Bude k dispozici = 0“ a současně je ve sloupci „Údaj Sklad 1“ poznámka „ukončeno“ nebo „doprodej“
        // ve sloupci „Údaj Sklad 1“ je poznámka „POS“.

        // SelectColumns
        // RenameColumns
        // InsertColumns
        // Join

        // ArrayExtensions
        // IntExtensions

        // Logger -> understand it

        [Test]
        public void CheckAvailableColumnsThrowsWhenMandatoryColumnsAreMissing()
        {
            var columns = new List<string>
            {
                "Produkt",
                "Katalogové číslo",
                "Cena DMOC EUR",
                "Údaj 1",
            };
            var table = new Table(Logger, columns, "Produkt");

            var ex = Assert.Throws<InvalidDataException>(() => table.CheckAvailableColumns());
            Assert.AreEqual("Data do not contain the following columns: Cena, K dispozici, Výrobce, Údaj 2.",
                ex.Message);
        }

        [Test]
        public void CheckAvailableColumnsDoesNotThrowWhenOptionalColumnsAreMissing()
        {
            var columns = new List<string>
            {
                "Produkt",
                "Katalogové číslo",
                "Cena",
                "Cena DMOC EUR",
                "K dispozici",
                "Výrobce",
                "Údaj 1",
                "Údaj 2",
                // some optional columns
                "Údaj sklad 1",
                "Popis",
                "DODAT",
            };
            var table = new Table(Logger, columns, "Produkt");

            Assert.DoesNotThrow(() => table.CheckAvailableColumns());
        }
        
        [Test]
        public void CheckAvailableColumnsAllowsUnknownColumns()
        {
            var columns = new List<string>
            {
                "Produkt",
                "Katalogové číslo",
                "Cena",
                "Cena DMOC EUR",
                "K dispozici",
                "Výrobce",
                "Údaj 1",
                "Údaj 2",
                // unknown column
                "Unknown column",
            };
            var table = new Table(Logger, columns, "Produkt");

            Assert.DoesNotThrow(() => table.CheckAvailableColumns());
        }
    }
}
