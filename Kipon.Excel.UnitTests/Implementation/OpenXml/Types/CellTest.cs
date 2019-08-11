using NUnit.Framework;
using System;
using Kipon.Excel.WriterImplementation.OpenXml.Types;

namespace Kipon.Excel.UnitTests.WriterImplementation.OpenXml.Types
{

    public class CellTest
    {
        [Test]
        public void ConstructorTest()
        {
            // Test string base constructor
            Assert.IsTrue("A2" == new CellBox { Cell = "A2" }.Cell);
            Assert.IsTrue("A" == new CellBox { Cell = "A2" }.Cell.Column.Value);
            Assert.IsTrue(1 == new CellBox { Cell = "A2" }.Cell.Row.Value);

            Assert.Throws<NullReferenceException>(() => new CellBox { Cell = null });

            Assert.Throws<ArgumentOutOfRangeException>(() => new CellBox { Cell = "A0" });
            Assert.Throws<ArgumentException>(() => new CellBox { Cell = "Æ1" });
            Assert.Throws<ArgumentException>(() => new CellBox { Cell = "1" });
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
