using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace Kipon.Excel.UnitTests.Lang
{
    public class ObjectDictionaryTest
    {
        [Test]
        public void ContainerKeyTest()
        {
            var dic = new Dictionary<int, string>();

            var x = 10;

            dic.Add(10, "halo");


            Assert.IsTrue(dic.ContainsKey(x));
        }
    }
}
