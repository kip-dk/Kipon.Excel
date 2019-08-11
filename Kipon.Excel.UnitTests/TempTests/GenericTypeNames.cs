using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kipon.Excel.UnitTests.TempTests
{
    public class GenericTypeNames
    {
        [Test]
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
