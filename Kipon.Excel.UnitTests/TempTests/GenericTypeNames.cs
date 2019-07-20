using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.UnitTests.TempTests
{
    [TestClass]
    public class GenericTypeNames
    {
        [TestMethod]
        public void ShowClassName()
        {
            var list = new List<Data>();

            var fullname = list.GetType().FullName;

        }


        public class Data
        {
            public string Noget { get; set; }
            public string Mere { get; set; }
        }
    }
}
