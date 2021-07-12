using NUnit.Framework;

namespace Tests
{
    public class Tests
    {
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
        public void Test1()
        {
            new Table();
        }
    }
}
