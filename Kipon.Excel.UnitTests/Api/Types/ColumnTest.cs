using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Kipon.Excel.Api.Types;

namespace Kipon.Excel.UnitTests.Api.Types
{
    [TestClass]
    public class ColumnTest
    {
        [TestMethod]
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

            Assert.ThrowsException<ArgumentOutOfRangeException>(() => new Column(-1));
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => new Column(Column.EXCEL_TOTAL_MAXCOLUMNS));

            // Test string based constructor
            Assert.AreEqual(0, new Column("A").Index);
            Assert.AreEqual(25, new Column("Z").Index);

            Assert.AreEqual(26, new Column("AA").Index);
            Assert.AreEqual(701, new Column("ZZ").Index);

            Assert.AreEqual(702, new Column("AAA").Index);
            Assert.AreEqual(1070, new Column("AOE").Index);
            Assert.AreEqual(1426, new Column("BBW").Index);

            Assert.ThrowsException<NullReferenceException>(() => new Column(null));
            Assert.ThrowsException<ArgumentException>(() => new Column(string.Empty));
            Assert.ThrowsException<ArgumentException>(() => new Column("AAAA"));
            Assert.ThrowsException<ArgumentException>(() => new Column("a"));
            Assert.ThrowsException<ArgumentOutOfRangeException>(() => new Column("XFE"));
        }

        [TestMethod]
        public void ColumnStringBoxingTest()
        {
            Assert.IsTrue("AAA" == new Column("AAA"));

            var ts = new ColumnTestClass { C = "AAA" };
            Assert.AreEqual(702, ts.C.Index);
        }

        [TestMethod]
        public void ColumnIntBoxingTest()
        {
            Assert.IsTrue(702 == new Column("AAA"));

            var ts = new ColumnTestClass { C = 702 };
            Assert.AreEqual(702, ts.C.Index);
        }

        public class ColumnTestClass
        {
            public Column C { get; set; }
        }
    }
}
