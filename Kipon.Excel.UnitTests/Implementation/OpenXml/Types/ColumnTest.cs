using NUnit.Framework;
using System;
using Kipon.Excel.WriterImplementation.OpenXml.Types;

namespace Kipon.Excel.UnitTests.WriterImplementation.OpenXml.Types
{
    public class ColumnTest
    {
        [Test]
        public void ColumnConstructorTest()
        {
            // Test int based constructor
            Assert.AreEqual("A", Column.getColumn(0).Value);
            Assert.AreEqual("Z", Column.getColumn(25).Value);

            Assert.AreEqual("AA", Column.getColumn(26).Value);
            Assert.AreEqual("ZZ", Column.getColumn(701).Value);

            Assert.AreEqual("AAA", Column.getColumn(702).Value);
            Assert.AreEqual("AOE", Column.getColumn(1070).Value);
            Assert.AreEqual("BBW", Column.getColumn(1426).Value);

            Assert.Throws<ArgumentOutOfRangeException>(() => Column.getColumn(-1));
            Assert.Throws<ArgumentOutOfRangeException>(() => Column.getColumn(Column.EXCEL_TOTAL_MAXCOLUMNS));

            // Test string based constructor
            Assert.AreEqual(0, Column.getColumn("A").Index);
            Assert.AreEqual(25, Column.getColumn("Z").Index);

            Assert.AreEqual(26, Column.getColumn("AA").Index);
            Assert.AreEqual(701, Column.getColumn("ZZ").Index);

            Assert.AreEqual(702, Column.getColumn("AAA").Index);
            Assert.AreEqual(1070, Column.getColumn("AOE").Index);
            Assert.AreEqual(1426, Column.getColumn("BBW").Index);

            Assert.Throws<NullReferenceException>(() => Column.getColumn(null));
            Assert.Throws<ArgumentException>(() => Column.getColumn(string.Empty));
            Assert.Throws<ArgumentException>(() => Column.getColumn("AAAA"));
            Assert.Throws<ArgumentException>(() => Column.getColumn("a"));
            Assert.Throws<ArgumentOutOfRangeException>(() => Column.getColumn("XFE"));
        }

        [Test]
        public void ColumnStringBoxingTest()
        {
            Assert.IsTrue("AAA" == Column.getColumn("AAA"));

            var ts = new ColumnTestClass { C = "AAA" };
            Assert.AreEqual(702, ts.C.Index);
        }

        [Test]
        public void ColumnIntBoxingTest()
        {
            Assert.IsTrue(702 == Column.getColumn("AAA"));

            var ts = new ColumnTestClass { C = 702 };
            Assert.AreEqual(702, ts.C.Index);
        }

        public class ColumnTestClass
        {
            internal Column C { get; set; }
        }
    }
}
