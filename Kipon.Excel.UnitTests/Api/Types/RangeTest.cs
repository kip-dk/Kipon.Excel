using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Kipon.Excel.Api.Types;

namespace Kipon.Excel.UnitTests.Api.Types
{
    [TestClass]

    public class RangeTest
    {
        [TestMethod]
        public void ConstructorTest()
        {
            Assert.AreEqual(new Range("A1:B2"), new Range("B2:A1"));
            Assert.AreEqual(new Range("A1", "B2"), new Range("B2", "A1"));

            Assert.ThrowsException<NullReferenceException>(() => new Range(null));
            Assert.ThrowsException<FormatException>(() => new Range("FEJL"));
        }

        [TestMethod]
        public void RangeStringBoxingTest()
        {
            Assert.IsTrue("A1:B2" == new Range("A1:B2"));
            Assert.IsFalse("A1:B2" != new Range("A1:B2"));

            Assert.IsTrue(new Range("A1:B2") == "A1:B2");
            Assert.IsFalse(new Range("A1:B2") != "A1:B2");
        }

        [TestMethod]
        public void ClassPropertyBoxingTest()
        {
            var t2 = new RangeTestClass { Range = "A1:B2" };
            Assert.AreEqual("A1:B2", t2.Range.Value);

            var t3 = new RangeTestClass { Range = "B2:A1" };
            Assert.AreEqual("A1:B2", t3.Range.Value);

            Assert.AreNotEqual("B2:A1", t3.Range.Value);

        }

        public class RangeTestClass
        {
            public Range Range { get; set; }
        }
    }
}
