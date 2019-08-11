using NUnit.Framework;
using System;
using Kipon.Excel.WriterImplementation.OpenXml.Types;
using System.Xml.Serialization;

namespace Kipon.Excel.UnitTests.WriterImplementation.OpenXml.Types
{
    public class RowTest
    {
        [Test]
        public void RowConstructorTest()
        {
            Assert.AreEqual(new Row(5), 5);
        }

        [Test]
        public void RowExceptionTest()
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => new Row(-1));
            Assert.Throws<ArgumentOutOfRangeException>(() => new Row(Row.EXCEL_TOTAL_MAXROWS + 1));
        }

        [Test]
        public void RowImplicitIntOperatorsTest()
        {
            {
                var row = (int)new Row(5);
                Assert.AreEqual(row, 5);
            }

            {
                var no = 5;
                var row = (Row)no;
                Assert.IsTrue(no == row);
                Assert.IsTrue(row == no);
            }
        }


        [Test]
        public void RowImplicitUIntOperatorsTest()
        {
            var t5 = new Row(5);

            //Assert.AreEqual(5, 4); // ADD uint impl. and test
        }

        [Test]
        public void RowClassPropertyTest()
        {
            var t4 = new Row(4);
            var t5 = new Row(5);
            {
                var mytestclass = new RowTestClass(5);
                Assert.AreEqual(t5, mytestclass.Row);
                Assert.AreNotEqual(t4, mytestclass.Row);
            }

            Assert.Throws<ArgumentOutOfRangeException>(() => new RowTestClass(-5));
        }

        [Test]
        public void RowMethodParameterTest()
        {
            new RowTestClass(4).Method(5);
            Assert.Throws<ArgumentOutOfRangeException>(() => new RowTestClass(4).Method(-1));
        }

        public class RowTestClass
        {
            private Row _row;
            internal RowTestClass(Row r)
            {
                this._row = r;
            }

            internal Row Row { get { return this._row; } }

            internal void Method(Row row)
            {
            }
        }
    }
}
