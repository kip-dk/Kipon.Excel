using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.UnitTests.TempTests
{
    public class PrimitiveTest
    {
        [Test]
        public void IsPrimitiveTest()
        {
            Assert.AreEqual(true, typeof(int).IsPrimitive);
        }
    }
}
