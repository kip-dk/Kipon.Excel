﻿using Kipon.Excel.Api;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.UnitTests.WriterImplementation.Factories
{
    [TestClass]
    public class SheetsResolverTest
    {
        #region decorated sheets
        [TestMethod]
        public void SheetAttributeDecoratePropertyTest()
        {
            var sheets = new DecoratedSheets();

            var resolver = new Kipon.Excel.WriterImplementation.Factories.SheetsResolver();
            var impl = resolver.Resolve(sheets);

            Assert.IsInstanceOfType(impl, typeof(IEnumerable<Kipon.Excel.Api.ISheet>));
            Assert.AreEqual(2, impl.Count());
            Assert.AreEqual("First sheet", impl.First().Title);
            Assert.AreEqual("Last sheet", impl.Last().Title);
        }
        #endregion

        #region multi sheet, duck type
        [TestMethod]
        public void DuckSheetsTest()
        {
            var sheets = new DuckSheets();

            var resolver = new Kipon.Excel.WriterImplementation.Factories.SheetsResolver();
            var impl = resolver.Resolve(sheets);
            Assert.IsInstanceOfType(impl, typeof(IEnumerable<Kipon.Excel.Api.ISheet>));
            Assert.AreEqual(3, impl.Count());

            Assert.AreEqual(nameof(DuckSheets.FirstSheet), impl.First().Title);
            Assert.AreEqual(nameof(DuckSheets.SecondSheet), impl.Skip(1).First().Title);
            Assert.AreEqual("Sheet title", impl.Last().Title);
        }
        #endregion

        #region single array duck type test
        [TestMethod]
        public void SingleSheetDuckArrayTest()
        {
            var data = new DuckSheet[0];
            var resolver = new Kipon.Excel.WriterImplementation.Factories.SheetsResolver();
            var impl = resolver.Resolve(data);
            Assert.IsInstanceOfType(impl, typeof(IEnumerable<Kipon.Excel.Api.ISheet>));
            Assert.AreEqual(1, impl.Count());
            Assert.AreEqual(nameof(DuckSheet), impl.First().Title);
        }
        #endregion

        #region single list duck type test
        [TestMethod]
        public void SingleSheetDuckListTest()
        {
            var data = new List<DuckSheet>();
            var resolver = new Kipon.Excel.WriterImplementation.Factories.SheetsResolver();
            var impl = resolver.Resolve(data);
            Assert.IsInstanceOfType(impl, typeof(IEnumerable<Kipon.Excel.Api.ISheet>));
            Assert.AreEqual(1, impl.Count());
            Assert.AreEqual(nameof(DuckSheet), impl.First().Title);
        }
        #endregion

        #region single ISheet actual implementation
        [TestMethod]
        public void SingleSheetImplementingISheet()
        {
            var data = new TestSheet("Test sheet");

            var resolver = new Kipon.Excel.WriterImplementation.Factories.SheetsResolver();
            var impl = resolver.Resolve(data);
            Assert.IsInstanceOfType(impl, typeof(IEnumerable<Kipon.Excel.Api.ISheet>));
            Assert.AreEqual(1, impl.Count());
            Assert.AreEqual("Test sheet", impl.First().Title);

        }
        #endregion

        #region exception test
        [TestMethod]
        public void UnresolveableTypeExceptionTest()
        {
            var x = new object();
            var resolver = new Kipon.Excel.WriterImplementation.Factories.SheetsResolver();
            Assert.ThrowsException<Kipon.Excel.Exceptions.UnresolveableTypeException>(() => resolver.Resolve(x));
        }
        #endregion

        #region null instance test
        [TestMethod]
        public void NullInstanceExceptionTest()
        {
            var resolver = new Kipon.Excel.WriterImplementation.Factories.SheetsResolver();
            object x = null;
            Assert.ThrowsException<Kipon.Excel.Exceptions.NullInstanceException>(() => resolver.Resolve(x));
        }
        #endregion

        #region helper impl
        public class DecoratedSheets
        {
            [Kipon.Excel.Attributes.Sheet(1)]
            public TestSheet Sheet1 => new TestSheet("First sheet");

            [Kipon.Excel.Attributes.Sheet(2)]
            public TestSheet Sheet2 => new TestSheet("Last sheet");

            public object NotDecorated => throw new NotImplementedException();
        }

        public class TestSheet : Kipon.Excel.Api.ISheet
        {
            private string _title;
            public TestSheet()
            {
            }

            public TestSheet(string title)
            {
                this._title = title;
            }

            public string Title => this._title;

            public IEnumerable<ICell> Cells => throw new NotImplementedException();
        }

        public class DuckSheets
        {
            public List<DuckSheet> FirstSheet => new List<DuckSheet>();
            public DuckSheet[] SecondSheet => new DuckSheet[0];

            public TestSheet ThirdSheet => new TestSheet("Sheet title");

            public string IgnoreString { get; set; }
            public int IgnoreInt { get; set; }
        }

        public class DuckSheet
        {
            public string Field1 { get; set; }
            public int Field2 { get; set; }
        }
        #endregion
    }
}
