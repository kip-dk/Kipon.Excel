using NUnit.Framework;
using Kipon.Excel.Api;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.UnitTests.WriterImplementation.Factories
{
    public class SpreadsheetResolverTest
    {
        [Test]
        public void ResolveTest()
        {
            {
                var impl = new SpreadsheetTest();
                var resolver = new Kipon.Excel.WriterImplementation.Factories.SpreadsheetResolver();
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
