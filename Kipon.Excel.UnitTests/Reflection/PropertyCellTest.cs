using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.UnitTests.Reflection
{
    public class PropertyCellTest
    {
        [Test]
        public void HasCellTest()
        {
            Assert.IsTrue(Kipon.Excel.Reflection.PropertyCell.HasCell(typeof(DuckSheet)));
            Assert.IsTrue(Kipon.Excel.Reflection.PropertyCell.HasCell(typeof(DecoratedSheet)));

            Assert.IsFalse(Kipon.Excel.Reflection.PropertyCell.HasCell(typeof(object)));
        }

        [Test]
        public void IsCellTest()
        {
            Assert.IsTrue(Kipon.Excel.Reflection.PropertyCell.IsCell(typeof(int)));
            Assert.IsTrue(Kipon.Excel.Reflection.PropertyCell.IsCell(typeof(int?)));
            Assert.IsFalse(Kipon.Excel.Reflection.PropertyCell.IsCell(typeof(object)));
        }


        public class DuckSheet
        {
            public int Property { get; set; }
        }


        public class DecoratedSheet
        {
            [Kipon.Excel.Attributes.Column(Title = "A property")]
            public object Property { get; set; }
        }
    }
}
