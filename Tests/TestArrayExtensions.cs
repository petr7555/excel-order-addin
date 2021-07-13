using System;
using System.Collections.Generic;
using System.Linq;
using ExcelOrderAddIn.Extensions;
using NUnit.Framework;

namespace Tests
{
    public class TestArrayExtensions
    {
        [Test]
        public void FromExcelMultidimArray()
        {
            var array = Array.CreateInstance(typeof(object), new[] {2, 3}, new[] {1, 1}) as object[,];
            Assert.NotNull(array);

            array.SetValue(1, 1, 1);
            array.SetValue(2, 1, 2);
            array.SetValue(3, 1, 3);
            array.SetValue(4, 2, 1);
            array.SetValue(5, 2, 2);
            array.SetValue(6, 2, 3);

            var expected = new[]
            {
                new object[] {1, 2, 3},
                new object[] {4, 5, 6},
            };

            Assert.AreEqual(expected, array.FromExcelMultidimArray());
        }

        [Test]
        public void FromEmptyExcelMultidimArray()
        {
            var array = Array.CreateInstance(typeof(object), new[] {0, 0}, new[] {1, 1}) as object[,];
            Assert.NotNull(array);

            var expected = new object[] { };

            Assert.AreEqual(expected, array.FromExcelMultidimArray());
        }

        [Test]
        public void ToExcelMultidimArrayFromJaggedArray()
        {
            var array = new[]
            {
                new object[] {1, 2, 3},
                new object[] {4, 5, 6},
            };

            var expected = new object[,]
            {
                {1, 2, 3},
                {4, 5, 6},
            };

            Assert.AreEqual(expected, array.ToExcelMultidimArray());
        }

        [Test]
        public void ToExcelMultidimArrayFromEmptyJaggedArray()
        {
            var array = new object[][] { };

            var expected = new object[,] { };

            Assert.AreEqual(expected, array.ToExcelMultidimArray());
        }

        [Test]
        public void ToExcelMultidimArrayFromList()
        {
            var array = new List<object> {1, 2, 3};

            var expected = new object[,]
            {
                {1, 2, 3},
            };

            Assert.AreEqual(expected, array.ToExcelMultidimArray());
        }

        [Test]
        public void ToExcelMultidimArrayFromEmptyList()
        {
            var array = new List<object>();

            var expected = new object[,] { };

            Assert.AreEqual(expected, array.ToExcelMultidimArray());
        }

        [Test]
        public void ToJaggedArrayFromIEnumerable()
        {
            var array = new[]
            {
                new object[] {1, 2, 3},
                new object[] {4, 5, 6},
            };

            var enumerable = array.Select(row => row.Select(col => col));

            var expected = new[]
            {
                new object[] {1, 2, 3},
                new object[] {4, 5, 6},
            };

            Assert.AreEqual(expected, array.ToJaggedArray());
            Assert.AreEqual(expected, enumerable.ToJaggedArray());
        }

        [Test]
        public void ToJaggedArrayFromEmptyIEnumerable()
        {
            var array = new object[][] { };

            var enumerable = array.Select(row => row.Select(col => col));

            var expected = new object[][] { };

            Assert.AreEqual(expected, array.ToJaggedArray());
            Assert.AreEqual(expected, enumerable.ToJaggedArray());
        }
    }
}
