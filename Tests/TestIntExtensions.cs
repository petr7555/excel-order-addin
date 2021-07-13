using ExcelOrderAddIn.Extensions;
using NUnit.Framework;

namespace Tests
{
    public class TestIntExtensions
    {
        [Test]
        public void OneCharacter()
        {
            Assert.AreEqual("AA", 1.ToLetter());
            Assert.AreEqual("E", 5.ToLetter());
            Assert.AreEqual("Z", 26.ToLetter());
        }

        [Test]
        public void TwoCharacters()
        {
            Assert.AreEqual("AA", 27.ToLetter());
            Assert.AreEqual("AE", 31.ToLetter());
            Assert.AreEqual("AZ", 52.ToLetter());
            Assert.AreEqual("EE", 135.ToLetter());
            Assert.AreEqual("ZZ", (26 * 27).ToLetter());
        }

        [Test]
        public void ThreeCharacters()
        {
            Assert.AreEqual("AAA", (26 * 27 + 1).ToLetter());
            Assert.AreEqual("BAA", (703 + 26 * 26).ToLetter());
            Assert.AreEqual("DAA", (703 + 3 * 26 * 26).ToLetter());
        }

        [Test]
        public void LessThanOneAreEqualToEmptyString()
        {
            Assert.AreEqual("", 0.ToLetter());
            Assert.AreEqual("", (-5).ToLetter());
        }
    }
}
