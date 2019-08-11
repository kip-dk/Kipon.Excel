using NUnit.Framework;
using System;
using Kipon.Excel.WriterImplementation.OpenXml.Types;

namespace Kipon.Excel.UnitTests.WriterImplementation.OpenXml.Types
{
    public class RangeTest
    {
        [Test]
        public void ConstructorTest()
        {
            Assert.AreEqual(new Range("A1:B2"), new Range("B2:A1"));
            Assert.AreEqual(new Range("A1", "B2"), new Range("B2", "A1"));

            Assert.Throws<NullReferenceException>(() => new Range(null));
            Assert.Throws<FormatException>(() => new Range("FEJL"));
        }

        [Test]
        public void RangeStringBoxingTest()
        {
            Assert.IsTrue("A1:B2" == new Range("A1:B2"));
            Assert.IsFalse("A1:B2" != new Range("A1:B2"));

            Assert.IsTrue(new Range("A1:B2") == "A1:B2");
            Assert.IsFalse(new Range("A1:B2") != "A1:B2");
        }

        [Test]
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
            internal Range Range { get; set; }
        }
    }
}
