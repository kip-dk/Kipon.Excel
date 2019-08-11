using Kipon.Excel.Api;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Attributes;

namespace Kipon.Excel.UnitTests.Reflection
{
    [TestClass]
    public class PropertySheetTest
    {
        [TestMethod]
        public void ForTypeTest()
        {
            var t1 = Kipon.Excel.Reflection.PropertySheet.ForType(typeof(DuckSheet));
            var t2 = Kipon.Excel.Reflection.PropertySheet.ForType(typeof(DuckSheet));
            Assert.AreEqual(t2, t1);
        }

        [TestMethod]
        public void IsPropertySheetTest()
        {
            Assert.IsTrue(Kipon.Excel.Reflection.PropertySheet.IsPropertySheet(typeof(DuckSheet)));
            Assert.IsTrue(Kipon.Excel.Reflection.PropertySheet.IsPropertySheet(typeof(DecoratedSheet)));
            Assert.IsFalse(Kipon.Excel.Reflection.PropertySheet.IsPropertySheet(typeof(SheetImpl)));
            Assert.IsFalse(Kipon.Excel.Reflection.PropertySheet.IsPropertySheet(typeof(object)));
        }


        [TestMethod]
        public void PropertiesTest()
        {
            var ps = Kipon.Excel.Reflection.PropertySheet.ForType(typeof(DuckSheet));
            Assert.AreEqual(6, ps.Properties.Count());

            // sort
            Assert.AreEqual(11, ps.Properties.First().sort);
            Assert.AreEqual(99, ps.Properties.Last().sort);

            // title
            Assert.AreEqual(nameof(DuckSheet.Number), ps.Properties.First().title);
            Assert.AreEqual("The null number", ps.Properties.Last().title);

            // Maxlength
            Assert.AreEqual(32, ps[nameof(DuckSheet.StringField)].maxlength);

            // Decimals
            Assert.AreEqual(2, ps[nameof(DuckSheet.Amount)].decimals);
            Assert.IsFalse(ps[nameof(DuckSheet.Amount)].isReadonly);
            Assert.IsFalse(ps[nameof(DuckSheet.Amount)].isHidden);

            // Default readonly
            Assert.IsTrue(ps[nameof(DuckSheet.Id)].isReadonly);

            // Explicit readonly
            Assert.IsTrue(ps[nameof(DuckSheet.StringField)].isReadonly);

            // Hidden
            Assert.IsTrue(ps[nameof(DuckSheet.HideMe)].isHidden);
        }

        public class DecoratedSheet
        {
            [Kipon.Excel.Attributes.Column(Title="A field")]
            public object Field { get; set; }
        }

        public class DuckSheet
        {
            [Sort(11)]
            public int Number { get; set; }

            [Sort(12)]
            [MaxLength(32)]
            [Readonly]
            public string StringField { get; set; }

            [Sort(13)]
            [Decimals(2)]
            public decimal Amount { get; set; }

            [Sort(20)]
            public Guid Id { get; private set; }

            [Hidden]
            [Sort(21)]
            public bool HideMe { get; set; }

            [Title("The null number")]
            [Sort(99)]
            public int? NullNumber { get; set; }
        }

        public class SheetImpl : Kipon.Excel.Api.ISheet
        {
            public string Title => throw new NotImplementedException();

            public IEnumerable<ICell> Cells => throw new NotImplementedException();
        }
    }
}
