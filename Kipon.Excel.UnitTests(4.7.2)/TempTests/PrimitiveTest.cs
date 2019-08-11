using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.UnitTests.TempTests
{
    [TestClass]
    public class PrimitiveTest
    {
        [TestMethod]
        public void IsPrimitiveTest()
        {
            Assert.AreEqual(true, typeof(int).IsPrimitive);
        }
    }
}
