using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.UnitTests.Implementation.Factories
{
    [TestClass]
    public class AbstractBaseResolverTest
    {
        [TestMethod]
        public void ResolveTest()
        {
            var impl = new List<Kipon.Excel.Api.ISheet>();
        }


        internal class Resolver<T> : Kipon.Excel.Implementation.Factories.AbstractBaseResolver<T> where T: class
        {
        }
    }
}
