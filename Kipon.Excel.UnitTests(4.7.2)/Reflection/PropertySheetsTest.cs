using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kipon.Excel.Attributes;
using Kipon.Excel.Api;

namespace Kipon.Excel.UnitTests.Reflection
{
    [TestClass]
    public class PropertySheetsTest
    {
        [TestMethod]
        public void ForTypeTest()
        {
            {
                var sh1 = Kipon.Excel.Reflection.PropertySheets.ForType(typeof(DecoratedSheets));
                var sh2 = Kipon.Excel.Reflection.PropertySheets.ForType(typeof(DecoratedSheets));
                Assert.AreEqual(sh1, sh2);

                Assert.AreEqual(2, sh1.Properties.Count());
            }

            {
                var dc1 = Kipon.Excel.Reflection.PropertySheets.ForType(typeof(DuckSheets));
                Assert.AreEqual(2, dc1.Properties.Count());

                Assert.AreEqual(2, dc1.Properties.Last().Sort);
                Assert.AreEqual("Sheet1", dc1.Properties.First().Title);
                Assert.AreEqual("Second sheet", dc1.Properties.Last().Title);

                Assert.AreEqual(typeof(PropertySheetTest.DecoratedSheet), dc1.Properties.First().ElementType);
                Assert.AreEqual(typeof(PropertySheetTest.DecoratedSheet), dc1.Properties.First().ElementType);
            }
        }

        [TestMethod]
        public void IsPropertySheetsTest()
        {
            Assert.IsTrue(Kipon.Excel.Reflection.PropertySheets.IsPropertySheets(typeof(DecoratedSheets)));
            Assert.IsTrue(Kipon.Excel.Reflection.PropertySheets.IsPropertySheets(typeof(DuckSheets)));

            Assert.IsFalse(Kipon.Excel.Reflection.PropertySheets.IsPropertySheets(typeof(SheetsImplementation)));
            Assert.IsFalse(Kipon.Excel.Reflection.PropertySheets.IsPropertySheets(typeof(object)));
            Assert.IsFalse(Kipon.Excel.Reflection.PropertySheets.IsPropertySheets(typeof(HasNoPublicProperties)));
        }


        public class HasNoPublicProperties
        {
        }

        public class SheetsImplementation : Kipon.Excel.Api.ISpreadsheet
        {
            public IEnumerable<ISheet> Sheets => throw new NotImplementedException();
        }


        public class DecoratedSheets
        {
            [Sheet(1)]
            public List<PropertySheetTest.DecoratedSheet> Sheet1 { get; set; }

            [Sheet(2)]
            public List<PropertySheetTest.DuckSheet> Sheet2 { get; set; }

            public List<object> NotASheet { get; set; }
        }

        public class DuckSheets
        {
            [Sort(1)]
            public List<PropertySheetTest.DecoratedSheet> Sheet1 { get; set; }

            [Sort(2)]
            [Title("Second sheet")]
            public PropertySheetTest.DuckSheet[] Sheet2 { get; set; }

            public List<object> NotASheet { get; set; }
        }
    }
}
