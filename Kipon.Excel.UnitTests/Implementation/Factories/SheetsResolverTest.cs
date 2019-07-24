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
        #region decorated sheets
        [TestMethod]
        public void SheetAttributeDecoratePropertyTest()
        {
            var sheets = new DecoratedSheets();

            var resolver = new Kipon.Excel.Implementation.Factories.SheetsResolver();
            var impl = resolver.Resolve(sheets);
            Assert.IsInstanceOfType(impl, typeof(IEnumerable<Kipon.Excel.Api.ISheet>));
            Assert.AreEqual(2, impl.Count());
        }

        [TestMethod]
        public void DuckSheetsTest()
        {
            var sheets = new DuckSheets();

            var resolver = new Kipon.Excel.Implementation.Factories.SheetsResolver();
            var impl = resolver.Resolve(sheets);
            Assert.IsInstanceOfType(impl, typeof(IEnumerable<Kipon.Excel.Api.ISheet>));
            Assert.AreEqual(3, impl.Count());
        }

        public class DecoratedSheets
        {
            [Kipon.Excel.Attributes.Sheet(1)]
            public TestSheet Sheet1 => new TestSheet();

            [Kipon.Excel.Attributes.Sheet(2)]
            public TestSheet Sheet2 => new TestSheet();

            public object NotDecorated => throw new NotImplementedException();
        }
        #endregion

        public class TestSheet : Kipon.Excel.Api.ISheet
        {
            public string Title => throw new NotImplementedException();

            public IEnumerable<ICell> Cells => throw new NotImplementedException();
        }

        public class DuckSheets
        {
            public List<DuckSheet> FirstSheet => new List<DuckSheet>();
            public DuckSheet[] SecondSheet => new DuckSheet[0];

            public TestSheet ThirdSheet => new TestSheet();

            public string IgnoreString { get; set; }
            public int IgnoreInt { get; set; }
        }

        public class DuckSheet
        {
            public string Field1 { get; set; }
            public int Field2 { get; set; }
        }
    }
}
