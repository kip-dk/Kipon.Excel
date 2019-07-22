using Kipon.Excel.Api;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.UnitTests.Implementation.Factories
{
    [TestClass]
    public class SheetsResolverTest
    {
        [TestMethod]
        public void SheetAttributeDecoratePropertyTest()
        {
            var sheets = new DecoratedSheets();

            var resolver = new Kipon.Excel.Implementation.Factories.SheetsResolver<Kipon.Excel.Implementation.Models.Sheets.AbstractBaseSheets>();
            var impl = resolver.Resolve(sheets);
            Assert.IsInstanceOfType(impl, typeof(IEnumerable<Kipon.Excel.Api.ISheet>));
            Assert.AreEqual(2, impl.Count());
        }

        public class DecoratedSheets
        {
            [Kipon.Excel.Attributes.Sheet(1)]
            public TestSheet Sheet1 => new TestSheet();

            [Kipon.Excel.Attributes.Sheet(2)]
            public TestSheet Sheet2 => new TestSheet();

            public object NotDecorated => throw new NotImplementedException();
        }

        public class TestSheet : Kipon.Excel.Api.ISheet
        {
            public string Title => throw new NotImplementedException();

            public IEnumerable<ICell> Cells => throw new NotImplementedException();
        }
    }
}
