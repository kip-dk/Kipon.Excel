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
            Assert.AreEqual("A", new Column(0).Value);
            Assert.AreEqual("Z", new Column(25).Value);

            Assert.AreEqual("AA", new Column(26).Value);
            Assert.AreEqual("ZZ", new Column(701).Value);

            Assert.AreEqual("AAA", new Column(702).Value);
            Assert.AreEqual("AOE", new Column(1070).Value);
            Assert.AreEqual("BBW", new Column(1426).Value);

            Assert.Throws<ArgumentOutOfRangeException>(() => new Column(-1));
            Assert.Throws<ArgumentOutOfRangeException>(() => new Column(Column.EXCEL_TOTAL_MAXCOLUMNS));

            // Test string based constructor
            Assert.AreEqual(0, new Column("A").Index);
            Assert.AreEqual(25, new Column("Z").Index);

            Assert.AreEqual(26, new Column("AA").Index);
            Assert.AreEqual(701, new Column("ZZ").Index);

            Assert.AreEqual(702, new Column("AAA").Index);
            Assert.AreEqual(1070, new Column("AOE").Index);
            Assert.AreEqual(1426, new Column("BBW").Index);

            Assert.Throws<NullReferenceException>(() => new Column(null));
            Assert.Throws<ArgumentException>(() => new Column(string.Empty));
            Assert.Throws<ArgumentException>(() => new Column("AAAA"));
            Assert.Throws<ArgumentException>(() => new Column("a"));
            Assert.Throws<ArgumentOutOfRangeException>(() => new Column("XFE"));
        }

        [Test]
        public void ColumnStringBoxingTest()
        {
            Assert.IsTrue("AAA" == new Column("AAA"));

            var ts = new ColumnTestClass { C = "AAA" };
            Assert.AreEqual(702, ts.C.Index);
        }

        [Test]
        public void ColumnIntBoxingTest()
        {
            Assert.IsTrue(702 == new Column("AAA"));

            var ts = new ColumnTestClass { C = 702 };
            Assert.AreEqual(702, ts.C.Index);
        }

        public class ColumnTestClass
        {
            internal Column C { get; set; }
        }
    }
}
