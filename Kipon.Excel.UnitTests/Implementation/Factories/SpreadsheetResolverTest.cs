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
    public class SpreadsheetResolverTest
    {
        [TestMethod]
        public void ResolveTest()
        {
            {
                var impl = new SpreadsheetTest();
                var resolver = new Kipon.Excel.Implementation.Factories.SpreadsheetResolver();
                var other = resolver.Resolve(impl);
                Assert.AreEqual(impl, other);
            }
        }

        public class SpreadsheetTest : Kipon.Excel.Api.ISpreadsheet
        {
            public IEnumerable<ISheet> Sheets => throw new NotImplementedException();
        }
    }
}
