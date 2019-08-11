using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Kipon.Excel.WriterImplementation.OpenXml.Types;

namespace Kipon.Excel.UnitTests.WriterImplementation.OpenXml.Types
{
    [TestClass]

    public class CellTest
    {
        [TestMethod]
        public void ConstructorTest()
        {
            // Test string base constructor
            Assert.IsTrue("A2" == new CellBox { Cell = "A2" }.Cell);
            Assert.IsTrue("A" == new CellBox { Cell = "A2" }.Cell.Column.Value);
            Assert.IsTrue(1 == new CellBox { Cell = "A2" }.Cell.Row.Value);

            Assert.ThrowsException<NullReferenceException>(() => new CellBox { Cell = null });

            Assert.ThrowsException<ArgumentOutOfRangeException>(() => new CellBox { Cell = "A0" });
            Assert.ThrowsException<ArgumentException>(() => new CellBox { Cell = "Æ1" });
            Assert.ThrowsException<ArgumentException>(() => new CellBox { Cell = "1" });
        }

        public class CellBox
        {
            internal Cell Cell { get; set; }
        }

        internal interface IWithCellProperty
        {
            Cell Cell { get; }
        }

        internal class WithCellProperty : IWithCellProperty
        {
            public Cell Cell => "A1";
        }
    }
}
