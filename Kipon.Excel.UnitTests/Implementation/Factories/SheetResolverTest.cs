using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.UnitTests.WriterImplementation.Factories
{
    public class SheetResolverTest
    {
        [Test]
        public void DecoratedSheetTest()
        {
            var sheetResolver = new Kipon.Excel.WriterImplementation.Factories.SheetResolver();

            /*
            {
                DecoratedSheet[] data = null;
                var sheet = sheetResolver.Resolve(data);
                Assert.AreEqual("DecoratedSheet", sheet.Title);
                Assert.AreEqual(3, sheet.Cells.Count());
            }
            */

            // empty array te
            {
                var data = new DecoratedSheet[0];

                var sheet = sheetResolver.Resolve(data);
                Assert.AreEqual("DecoratedSheet", sheet.Title);
                Assert.AreEqual(3, sheet.Cells.Count());
            }

            // single row array test
            {
                var data = new DecoratedSheet[]
                {
                    new DecoratedSheet{ }
                };

                var sheet = sheetResolver.Resolve(data);
                Assert.AreEqual("DecoratedSheet", sheet.Title);
                Assert.AreEqual(6, sheet.Cells.Count());
            }
        }

        public class DecoratedSheet
        {
            [Kipon.Excel.Attributes.Column(Sort = 1)]
            public int Number { get; set; }

            [Kipon.Excel.Attributes.Column(Sort = 2, Title = "Text title")]
            public string Text { get; set; }

            [Kipon.Excel.Attributes.Column(Sort = 3, Title = "Bool field")]
            public bool BoolValue { get; }

            public string UndecoratedProperty { get; set; }
        }
    }
}
