using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Kipon.Excel.Types;

namespace Kipon.Excel.UnitTests.PublicAPI.Types
{
    [TestClass]

    public class CellTest
    {
        [TestMethod]
        public void ConstructorTest()
        {
            // Test string base constructor
            Assert.AreEqual("A2", new Cell("A2").Value);
            Assert.AreEqual("A", new Cell("A2").Column.Value);
            Assert.AreEqual(1, new Cell("A2").Row.Value);

            // Test string, line constructor
            Assert.AreEqual(new Cell("A", 1), new Cell("A1"));
            Assert.AreNotEqual(new Cell("A", 1), new Cell("A2"));

            Assert.AreEqual(new Cell("A2"), new Cell(new Column("A"), new Row(1)));

            Assert.ThrowsException<NullReferenceException>(() => new Cell(null, 1));
            Assert.ThrowsException<NullReferenceException>(() => new Cell(null, new Row(0)));

            Assert.ThrowsException<ArgumentOutOfRangeException>(() => new Cell("A", 0));
            Assert.ThrowsException<ArgumentException>(() => new Cell("", 1));
        }

        [TestMethod]
        public void CellStringBoxingTest()
        {
            Assert.IsTrue("A1" == new Cell("A1"));
            Assert.IsFalse("A1" != new Cell("A1"));

            Assert.IsTrue(new Cell("A1") == "A1");
            Assert.IsFalse(new Cell("A1") != "A1" );

        }
    }
}
